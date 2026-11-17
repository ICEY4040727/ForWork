/**
 * DOCX mapping between OOXML documents and Univer `IDocumentData`.
 *
 * SUPPORTED: paragraph order and paragraph breaks, bold/italic/underline text
 * runs, tab characters inside runs, named heading paragraph styles
 * (`Heading1` through `Heading5`), and flow tables. Table support covers rows,
 * cells, cell text, grid column widths, preferred cell widths, and the
 * paragraph/section-break pair each cell requires.
 *
 * The installed Univer runtime tokens define the stream conventions used here:
 * `\r` ends a paragraph and `\n` ends a section. Tables use `\x1A` (table
 * start), `\x1B` (row start), `\x1C` (cell start), `\x1D` (cell end), `\x0E`
 * (row end), and `\x0F` (table end). `IParagraph.startIndex` points at its
 * paragraph break, `ISectionBreak.startIndex` points at its section break, and
 * `ITextRun.st`/`ed` are inclusive/exclusive UTF-16 offsets into `dataStream`.
 * A table range in `body.tables` ends immediately after the table end token and
 * its structural rows live in `tableSource[tableId]`.
 *
 * UNSUPPORTED: images, charts, drawings, shapes, comments, tracked changes,
 * headers, footers, macros, footnotes and endnotes, hyperlinks, lists and
 * numbering, sections and page layout, theme styles, fonts, colors, sizes,
 * alignment, strikethrough, highlighting, paragraph spacing, indentation,
 * borders, merged cells with spans, nested tables, and every style
 * other than the heading styles above. A read/write round-trip does not
 * preserve these. Macro-bearing packages are read without executing or
 * retaining macros.
 */
import { XMLParser } from 'fast-xml-parser'
import { SyntaxValidator } from 'fast-xml-validator'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import type { Unzipped, Zippable } from 'fflate'
import type {
  BooleanNumber,
  ICustomTable,
  IDocumentData,
  IParagraph,
  ISectionBreak,
  ITable,
  ITableCell,
  ITableCellMargin,
  ITableRow,
  ITables,
  ITextDecoration,
  ITextRun,
  ITextStyle,
  IWidthInTableSize,
  LocaleType,
  NamedStyleType,
  TableAlignmentType,
  TableRowHeightRule,
  TableSizeType,
  TableTextWrapType,
  TextDecoration,
} from '@univerjs/core'
import { type OfficeExchangeConfig, resolveConfig } from '../config.ts'
import { OfficeExchangeError } from '../errors.ts'

const DOCUMENT_XML = 'word/document.xml'
const PARAGRAPH = '\r'
const SECTION_BREAK = '\n'
const TABLE_START = '\x1A'
const TABLE_ROW_START = '\x1B'
const TABLE_CELL_START = '\x1C'
const TABLE_CELL_END = '\x1D'
const TABLE_ROW_END = '\x0E'
const TABLE_END = '\x0F'
const TAB = '\t'

const ENTITY_COUNT_LIMIT = 10_000
const EXPANSION_BYTE_LIMIT = 200 * 1024 * 1024
const TWIPS_PER_PIXEL = 15
const DEFAULT_COLUMN_WIDTH = 160
const DEFAULT_ROW_HEIGHT = 30
const DETERMINISTIC_DATE = new Date('2026-01-01T00:00:00.000Z')

const BOOLEAN_TRUE: BooleanNumber = 1
const TEXT_DECORATION_SINGLE: TextDecoration = 12
const TEXT_DECORATION_DOUBLE: TextDecoration = 10
const LOCALE_EN_US = 'enUS' as LocaleType
const TABLE_ALIGN_START: TableAlignmentType = 0
const TABLE_WRAP_NONE: TableTextWrapType = 0
const TABLE_SIZE_SPECIFIED: TableSizeType = 1
const TABLE_ROW_HEIGHT_AUTO: TableRowHeightRule = 0
const RELATIVE_H_PAGE = 0
const RELATIVE_V_PAGE = 0

const HEADING_STYLE_IDS: readonly string[] = ['Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5']

const NAMED_STYLE_BY_ID: Readonly<Record<string, NamedStyleType>> = {
  Heading1: 4,
  Heading2: 5,
  Heading3: 6,
  Heading4: 7,
  Heading5: 8,
}

const XML_PARSER_OPTIONS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: true,
  preserveOrder: true,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
}

const TABLE_BORDERS = '<w:tblBorders>'
  + '<w:top w:val="single" w:sz="4" w:color="auto"/>'
  + '<w:left w:val="single" w:sz="4" w:color="auto"/>'
  + '<w:bottom w:val="single" w:sz="4" w:color="auto"/>'
  + '<w:right w:val="single" w:sz="4" w:color="auto"/>'
  + '<w:insideH w:val="single" w:sz="4" w:color="auto"/>'
  + '<w:insideV w:val="single" w:sz="4" w:color="auto"/>'
  + '</w:tblBorders>'

/** A dynamically shaped node produced by `fast-xml-parser` in preserved order. */
interface ParserNode {
  readonly [key: string]: unknown
}

/** One XML element with its already-validated children and attributes. */
interface XmlElement {
  /** Local element name with its namespace prefix removed. */
  readonly tag: string
  /** Child nodes in source order, including text nodes. */
  readonly children: readonly ParserNode[]
  /** Attribute values keyed by the parser's attribute prefix. */
  readonly attributes: ParserNode
}

/** Accumulates a document stream and the Univer metadata addressing it. */
interface ReadContext {
  readonly parts: string[]
  length: number
  readonly paragraphs: IParagraph[]
  readonly textRuns: ITextRun[]
  readonly sectionBreaks: ISectionBreak[]
  readonly tables: ICustomTable[]
  readonly tableSource: ITables
  readonly bounds: OfficeExchangeConfig
  cells: number
}

/** A table cell's content range inside `dataStream`. */
interface ParsedCell {
  /** Inclusive offset of the first content character. */
  readonly start: number
  /** Exclusive offset of the cell end token. */
  readonly end: number
}

/** A table parsed from `dataStream` together with its exclusive end. */
interface ParsedTable {
  /** Cell content ranges grouped by row. */
  readonly rows: ReadonlyArray<ReadonlyArray<ParsedCell>>
  /** Exclusive offset immediately after the table end token. */
  readonly end: number
}

/**
 * Read `.docx` bytes into a Univer document snapshot. Paragraph text, paragraph
 * breaks, bold/italic/underline runs, heading styles, and flow tables are
 * mapped; the unsupported features listed in the module note are not.
 * @param bytes - Complete DOCX package bytes.
 * @param config - Bounds applied before parsing; defaults when omitted.
 * @returns Document snapshot with Univer stream tokens, run styles, and table metadata.
 * @throws {OfficeExchangeError} `too-large` when a byte, package-entry,
 * expansion, or table-cell bound is exceeded; `conversion-failed` when the
 * input is not a readable Word document.
 */
export function readDocument(bytes: Uint8Array, config?: OfficeExchangeConfig): Promise<IDocumentData> {
  return Promise.resolve().then(() => readDocumentData(bytes, config))
}

function readDocumentData(bytes: Uint8Array, config?: OfficeExchangeConfig): IDocumentData {
  const bounds = resolveConfig(config ?? {})
  if (bytes.byteLength > bounds.maxFileBytes) {
    throw new OfficeExchangeError(
      'too-large',
      `Document is ${bytes.byteLength} bytes; the limit is ${bounds.maxFileBytes}`,
    )
  }
  const documentXml = readPackage(bytes)
  if (documentXml === undefined) {
    throw new OfficeExchangeError('conversion-failed', 'Package has no word/document.xml part')
  }
  const body = parseDocumentBody(documentXml)
  const context = createReadContext(bounds)
  for (const child of childElements(body)) {
    if (child.tag === 'p') appendParagraph(context, child)
    if (child.tag === 'tbl') appendTable(context, child)
  }
  finalizeBody(context)
  return {
    id: 'office-exchange-document',
    locale: LOCALE_EN_US,
    title: '',
    body: {
      dataStream: context.parts.join(''),
      textRuns: context.textRuns,
      paragraphs: context.paragraphs,
      sectionBreaks: context.sectionBreaks,
      tables: context.tables,
    },
    documentStyle: {},
    tableSource: context.tableSource,
    resources: [],
  }
}

/**
 * Write a Univer document snapshot to `.docx` bytes. The inverse of
 * `readDocument` for the supported features; the unsupported features listed in
 * the module note are absent from the output.
 * @param data - Document snapshot from `readDocument` or the editor.
 * @param config - Bounds validated before writing; defaults when omitted.
 * @returns A deterministic DOCX package that `readDocument` can reopen.
 * @throws {OfficeExchangeError} `too-large` when the table-cell bound is
 * exceeded; `conversion-failed` when the stream is structurally invalid or
 * serialization fails.
 */
export function writeDocument(data: IDocumentData, config?: OfficeExchangeConfig): Promise<Uint8Array> {
  return Promise.resolve().then(() => writeDocumentBytes(data, config))
}

function writeDocumentBytes(data: IDocumentData, config?: OfficeExchangeConfig): Uint8Array {
  const bounds = resolveConfig(config ?? {})
  const stream = data.body?.dataStream ?? ''
  const paragraphs = new Map<number, IParagraph>()
  for (const paragraph of data.body?.paragraphs ?? []) paragraphs.set(paragraph.startIndex, paragraph)
  const bodyXml = writeBody(stream, paragraphs, data, bounds)
  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(contentTypesXml()),
    '_rels/.rels': strToU8(rootRelsXml()),
    'word/_rels/document.xml.rels': strToU8(documentRelsXml()),
    [DOCUMENT_XML]: strToU8(documentXml(bodyXml)),
    'word/styles.xml': strToU8(stylesXml()),
  }
  const zippable: Zippable = {}
  for (const name of Object.keys(entries).sort()) {
    const content = entries[name]
    /* v8 ignore next -- Object.keys came from this record, so every lookup resolves. */
    if (content === undefined) continue
    zippable[name] = [content, { mtime: DETERMINISTIC_DATE }]
  }
  try {
    return zipSync(zippable, { level: 6 })
  } catch (error) {
    throw new OfficeExchangeError('conversion-failed', `Cannot write document: ${messageOf(error)}`)
  }
}

function createReadContext(bounds: OfficeExchangeConfig): ReadContext {
  return {
    parts: [],
    length: 0,
    paragraphs: [],
    textRuns: [],
    sectionBreaks: [],
    tables: [],
    tableSource: {},
    bounds,
    cells: 0,
  }
}

function append(context: ReadContext, text: string): void {
  if (text.length === 0) return
  context.parts.push(text)
  context.length += text.length
}

function lastCharacter(context: ReadContext): string | undefined {
  const last = context.parts[context.parts.length - 1]
  return last === undefined || last.length === 0 ? undefined : last[last.length - 1]
}

function appendParagraphBreak(context: ReadContext, styleId?: string): void {
  append(context, PARAGRAPH)
  const paragraph: IParagraph = {
    startIndex: context.length - 1,
    paragraphId: `paragraph-${context.paragraphs.length + 1}`,
  }
  if (styleId !== undefined) paragraph.styleId = styleId
  const namedStyle = styleId === undefined ? undefined : NAMED_STYLE_BY_ID[styleId]
  if (namedStyle !== undefined) paragraph.paragraphStyle = { namedStyleType: namedStyle }
  context.paragraphs.push(paragraph)
}

function appendParagraph(context: ReadContext, element: XmlElement): void {
  const properties = findChild(element, 'pPr')
  const style = properties === undefined ? undefined : findChild(properties, 'pStyle')
  const styleId = style === undefined ? undefined : attribute(style, '@_val')
  for (const child of childElements(element)) {
    if (child.tag !== 'r') continue
    appendRun(context, child)
  }
  appendParagraphBreak(context, styleId)
}

function appendRun(context: ReadContext, element: XmlElement): void {
  const properties = findChild(element, 'rPr')
  const style = properties === undefined ? undefined : readTextStyle(properties)
  const start = context.length
  let text = ''
  for (const child of childElements(element)) {
    if (child.tag === 't') text += textContent(child)
    if (child.tag === 'tab') text += TAB
  }
  append(context, text)
  if (style !== undefined && text.length > 0) {
    context.textRuns.push({ st: start, ed: context.length, ts: style })
  }
}

function readTextStyle(properties: XmlElement): ITextStyle | undefined {
  const style: ITextStyle = {}
  if (isEnabled(properties, 'b') === true) style.bl = BOOLEAN_TRUE
  if (isEnabled(properties, 'i') === true) style.it = BOOLEAN_TRUE
  const underline = readUnderline(properties)
  if (underline !== undefined) style.ul = underline
  return Object.keys(style).length === 0 ? undefined : style
}

function readUnderline(properties: XmlElement): ITextDecoration | undefined {
  const underline = findChild(properties, 'u')
  if (underline === undefined) return undefined
  const value = attribute(underline, '@_val')
  if (value === 'none') return undefined
  return {
    s: BOOLEAN_TRUE,
    t: value === 'double' ? TEXT_DECORATION_DOUBLE : TEXT_DECORATION_SINGLE,
  }
}

function isEnabled(properties: XmlElement, tag: string): boolean | undefined {
  const element = findChild(properties, tag)
  if (element === undefined) return undefined
  const value = attribute(element, '@_val')
  return value === undefined || (value !== '0' && value !== 'false' && value !== 'off')
}

function appendTable(context: ReadContext, element: XmlElement): void {
  if (lastCharacter(context) !== PARAGRAPH) appendParagraphBreak(context)
  const tableId = `table-${context.tables.length + 1}`
  const startIndex = context.length
  append(context, TABLE_START)
  const rows: ITableRow[] = []
  const gridWidths = readGridWidths(element)
  let columnCount = 1
  for (const row of childElements(element)) {
    if (row.tag !== 'tr') continue
    append(context, TABLE_ROW_START)
    const cells: ITableCell[] = []
    for (const cell of childElements(row)) {
      if (cell.tag !== 'tc') continue
      countCell(context)
      append(context, TABLE_CELL_START)
      for (const block of childElements(cell)) {
        if (block.tag === 'p') appendParagraph(context, block)
      }
      if (lastCharacter(context) !== PARAGRAPH) appendParagraphBreak(context)
      context.sectionBreaks.push({
        sectionId: `section-${context.sectionBreaks.length + 1}`,
        startIndex: context.length,
      })
      append(context, `${SECTION_BREAK}${TABLE_CELL_END}`)
      const width = readCellWidth(cell)
      cells.push(width === undefined ? { margin: cellMargin() } : { margin: cellMargin(), size: width })
    }
    append(context, TABLE_ROW_END)
    columnCount = Math.max(columnCount, cells.length)
    rows.push({ tableCells: cells, trHeight: { val: { v: DEFAULT_ROW_HEIGHT }, hRule: TABLE_ROW_HEIGHT_AUTO } })
  }
  append(context, TABLE_END)
  context.tables.push({ startIndex, endIndex: context.length, tableId })
  context.tableSource[tableId] = createTable(tableId, rows, columnCount, gridWidths)
}

function countCell(context: ReadContext): void {
  context.cells += 1
  if (context.cells > context.bounds.maxCells) {
    throw new OfficeExchangeError('too-large', `Document has more than ${context.bounds.maxCells} table cells`)
  }
}

function readGridWidths(element: XmlElement): readonly number[] {
  const grid = findChild(element, 'tblGrid')
  if (grid === undefined) return []
  const widths: number[] = []
  for (const column of childElements(grid)) {
    if (column.tag !== 'gridCol') continue
    const twips = positiveNumber(attribute(column, '@_w'))
    if (twips !== undefined) widths.push(twips / TWIPS_PER_PIXEL)
  }
  return widths
}

function readCellWidth(element: XmlElement): IWidthInTableSize | undefined {
  const properties = findChild(element, 'tcPr')
  if (properties === undefined) return undefined
  const width = findChild(properties, 'tcW')
  if (width === undefined) return undefined
  const twips = positiveNumber(attribute(width, '@_w'))
  if (twips === undefined) return undefined
  return { type: TABLE_SIZE_SPECIFIED, width: { v: twips / TWIPS_PER_PIXEL } }
}

function createTable(
  tableId: string,
  rows: readonly ITableRow[],
  columnCount: number,
  gridWidths: readonly number[],
): ITable {
  const widths: number[] = []
  for (let index = 0; index < columnCount; index += 1) {
    widths.push(gridWidths[index] ?? DEFAULT_COLUMN_WIDTH)
  }
  const totalWidth = widths.reduce((total, width) => total + width, 0)
  return {
    tableRows: [...rows],
    tableColumns: widths.map(width => ({
      size: { type: TABLE_SIZE_SPECIFIED, width: { v: width } },
    })),
    tableId,
    align: TABLE_ALIGN_START,
    indent: { v: 0 },
    textWrap: TABLE_WRAP_NONE,
    position: {
      positionH: { relativeFrom: RELATIVE_H_PAGE, posOffset: 0 },
      positionV: { relativeFrom: RELATIVE_V_PAGE, posOffset: 0 },
    },
    dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
    cellMargin: cellMargin(),
    size: { type: TABLE_SIZE_SPECIFIED, width: { v: totalWidth } },
  }
}

function cellMargin(): ITableCellMargin {
  return {
    start: { v: 10 },
    end: { v: 10 },
    top: { v: 5 },
    bottom: { v: 5 },
  }
}

function finalizeBody(context: ReadContext): void {
  if (lastCharacter(context) !== PARAGRAPH) appendParagraphBreak(context)
  context.sectionBreaks.push({
    sectionId: `section-${context.sectionBreaks.length + 1}`,
    startIndex: context.length,
  })
  append(context, SECTION_BREAK)
}

function writeBody(
  stream: string,
  paragraphs: ReadonlyMap<number, IParagraph>,
  data: IDocumentData,
  bounds: OfficeExchangeConfig,
): string {
  const cells = { count: 0 }
  let xml = ''
  let cursor = 0
  let index = 0
  while (index < stream.length) {
    const token = stream[index]
    if (token === TABLE_START) {
      const table = readTableRange(stream, index)
      xml += writeTable(stream, index, table, paragraphs, data, bounds, cells)
      index = table.end
      cursor = index
    } else if (token === PARAGRAPH) {
      xml += writeParagraph(stream, cursor, index, paragraphs.get(index), data)
      cursor = index + 1
      index += 1
    } else {
      index += 1
    }
  }
  const remainder = stream.slice(cursor).replaceAll(SECTION_BREAK, '')
  if (remainder.length > 0) {
    throw new OfficeExchangeError('conversion-failed', 'Document stream has text outside a paragraph')
  }
  return xml
}

function readTableRange(stream: string, start: number): ParsedTable {
  let index = start + 1
  const rows: ParsedCell[][] = []
  while (stream[index] === TABLE_ROW_START) {
    index += 1
    const cells: ParsedCell[] = []
    while (stream[index] === TABLE_CELL_START) {
      index += 1
      const end = findCellEnd(stream, index)
      if (end === undefined) {
        throw new OfficeExchangeError('conversion-failed', 'Document stream has an unterminated table cell')
      }
      cells.push({ start: index, end })
      index = end + 1
    }
    if (stream[index] !== TABLE_ROW_END) {
      throw new OfficeExchangeError('conversion-failed', 'Document stream has an unterminated table row')
    }
    index += 1
    rows.push(cells)
  }
  if (stream[index] !== TABLE_END) {
    throw new OfficeExchangeError('conversion-failed', 'Document stream has an unterminated table')
  }
  return { rows, end: index + 1 }
}

function findCellEnd(stream: string, start: number): number | undefined {
  let depth = 0
  for (let index = start; index < stream.length; index += 1) {
    const token = stream[index]
    if (token === TABLE_CELL_START) depth += 1
    if (token !== TABLE_CELL_END) continue
    if (depth === 0) return index
    depth -= 1
  }
  return undefined
}

function writeTable(
  stream: string,
  start: number,
  table: ParsedTable,
  paragraphs: ReadonlyMap<number, IParagraph>,
  data: IDocumentData,
  bounds: OfficeExchangeConfig,
  cells: { count: number },
): string {
  const source = tableSourceAt(data, start)
  const sourceColumns = source?.tableColumns ?? []
  const columnCount = Math.max(
    1,
    ...table.rows.map(row => row.length),
    sourceColumns.length,
  )
  const widths: number[] = []
  for (let index = 0; index < columnCount; index += 1) {
    widths.push(sourceColumns[index]?.size.width.v ?? DEFAULT_COLUMN_WIDTH)
  }
  let xml = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
    + TABLE_BORDERS
    + '</w:tblPr><w:tblGrid>'
  for (const width of widths) {
    xml += `<w:gridCol w:w="${String(Math.max(1, Math.round(width * TWIPS_PER_PIXEL)))}"/>`
  }
  xml += '</w:tblGrid>'
  for (const [rowIndex, row] of table.rows.entries()) {
    xml += '<w:tr>'
    for (const [cellIndex, cell] of row.entries()) {
      cells.count += 1
      if (cells.count > bounds.maxCells) {
        throw new OfficeExchangeError('too-large', `Document has more than ${bounds.maxCells} table cells`)
      }
      const sourceCell = source?.tableRows[rowIndex]?.tableCells[cellIndex]
      /* v8 ignore next -- columnCount sizes widths to cover every cell index. */
      const width = sourceCell?.size?.width.v ?? widths[cellIndex] ?? DEFAULT_COLUMN_WIDTH
      xml += '<w:tc><w:tcPr>'
        + `<w:tcW w:w="${String(Math.max(1, Math.round(width * TWIPS_PER_PIXEL)))}" w:type="dxa"/>`
        + `</w:tcPr>${writeCell(stream, cell, paragraphs, data)}</w:tc>`
    }
    xml += '</w:tr>'
  }
  return `${xml}</w:tbl>`
}

function tableSourceAt(data: IDocumentData, startIndex: number): ITable | undefined {
  const range = data.body?.tables?.find(table => table.startIndex === startIndex)
  return range === undefined ? undefined : data.tableSource?.[range.tableId]
}

function writeCell(
  stream: string,
  cell: ParsedCell,
  paragraphs: ReadonlyMap<number, IParagraph>,
  data: IDocumentData,
): string {
  let xml = ''
  let cursor = cell.start
  let index = cell.start
  while (index < cell.end) {
    const token = stream[index]
    if (token === PARAGRAPH) {
      xml += writeParagraph(stream, cursor, index, paragraphs.get(index), data)
      cursor = index + 1
      index += 1
    } else if (token === TABLE_START) {
      const nested = readTableRange(stream, index)
      index = nested.end
      cursor = index
    } else {
      index += 1
    }
  }
  const remainder = stream.slice(cursor, cell.end).replaceAll(SECTION_BREAK, '')
  if (remainder.length > 0) {
    throw new OfficeExchangeError('conversion-failed', 'Table cell text must end with a paragraph break')
  }
  return xml.length === 0 ? writeParagraph(stream, cell.start, cell.start, undefined, data) : xml
}

function writeParagraph(
  stream: string,
  start: number,
  end: number,
  paragraph: IParagraph | undefined,
  data: IDocumentData,
): string {
  const properties = paragraphProperties(paragraph)
  let xml = `<w:p>${properties}`
  const runs = runSlices(data.body?.textRuns ?? [], start, end)
  if (runs.length === 0) {
    xml += writeRun(stream.slice(start, end), undefined)
  } else {
    for (const run of runs) {
      xml += writeRun(stream.slice(run.start, run.end), run.style)
    }
  }
  return `${xml}</w:p>`
}

interface RunSlice {
  readonly start: number
  readonly end: number
  readonly style: ITextStyle | undefined
}

function runSlices(textRuns: readonly ITextRun[], start: number, end: number): readonly RunSlice[] {
  const boundaries = new Set<number>([start, end])
  const covered: RunSlice[] = []
  for (const run of textRuns) {
    const clippedStart = Math.max(start, run.st)
    const clippedEnd = Math.min(end, run.ed)
    if (clippedEnd <= clippedStart) continue
    boundaries.add(clippedStart)
    boundaries.add(clippedEnd)
    covered.push({ start: clippedStart, end: clippedEnd, style: run.ts })
  }
  if (covered.length === 0) return []
  const sorted = [...boundaries].sort((left, right) => left - right)
  const slices: RunSlice[] = []
  for (let index = 0; index < sorted.length - 1; index += 1) {
    const sliceStart = sorted[index]
    const sliceEnd = sorted[index + 1]
    /* v8 ignore next -- Boundaries come in sorted pairs, so both offsets are always present. */
    if (sliceStart === undefined || sliceEnd === undefined) continue
    const style = covered.find(run => run.start <= sliceStart && run.end >= sliceEnd)?.style
    slices.push({ start: sliceStart, end: sliceEnd, style })
  }
  return slices
}

function paragraphProperties(paragraph: IParagraph | undefined): string {
  const styleId = paragraph?.styleId ?? styleIdForNamedStyle(paragraph?.paragraphStyle?.namedStyleType)
  return styleId === undefined ? '' : `<w:pPr><w:pStyle w:val="${escapeXml(styleId)}"/></w:pPr>`
}

function styleIdForNamedStyle(namedStyle: NamedStyleType | undefined): string | undefined {
  if (namedStyle === undefined) return undefined
  return HEADING_STYLE_IDS[namedStyle - 4]
}

function writeRun(text: string, style: ITextStyle | undefined): string {
  const content = writeRunText(text)
  return content.length === 0 ? '' : `<w:r>${writeRunProperties(style)}${content}</w:r>`
}

function writeRunProperties(style: ITextStyle | undefined): string {
  if (style === undefined) return ''
  let properties = ''
  if (style.bl === BOOLEAN_TRUE) properties += '<w:b/>'
  if (style.it === BOOLEAN_TRUE) properties += '<w:i/>'
  if (style.ul?.s === BOOLEAN_TRUE) {
    properties += `<w:u w:val="${style.ul.t === TEXT_DECORATION_DOUBLE ? 'double' : 'single'}"/>`
  }
  return properties.length === 0 ? '' : `<w:rPr>${properties}</w:rPr>`
}

function writeRunText(text: string): string {
  let xml = ''
  let buffer = ''
  const flush = (): void => {
    if (buffer.length === 0) return
    xml += `<w:t xml:space="preserve">${escapeXml(buffer)}</w:t>`
    buffer = ''
  }
  for (const character of text) {
    if (character === TAB) {
      flush()
      xml += '<w:tab/>'
      continue
    }
    if (character.charCodeAt(0) >= 0x20) buffer += character
  }
  flush()
  return xml
}

function escapeXml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

function readPackage(bytes: Uint8Array): Uint8Array | undefined {
  let entryCount = 0
  let expandedBytes = 0
  let entries: Unzipped
  try {
    entries = unzipSync(bytes, {
      filter: (file) => {
        entryCount += 1
        if (entryCount > ENTITY_COUNT_LIMIT) {
          throw new OfficeExchangeError('too-large', `Document package contains more than ${ENTITY_COUNT_LIMIT} entries`)
        }
        expandedBytes += file.originalSize
        /* v8 ignore next 6 -- A 200 MB expansion is not testable; the entry guard covers this failure class. */
        if (expandedBytes > EXPANSION_BYTE_LIMIT) {
          throw new OfficeExchangeError(
            'too-large',
            `Document package expands beyond ${EXPANSION_BYTE_LIMIT} bytes`,
          )
        }
        return file.name === DOCUMENT_XML
      },
    })
  } catch (error) {
    if (error instanceof OfficeExchangeError) throw error
    throw new OfficeExchangeError('conversion-failed', `Cannot read OOXML package: ${messageOf(error)}`)
  }
  return entries[DOCUMENT_XML]
}

function parseDocumentBody(documentXml: Uint8Array): XmlElement {
  const xml = strFromU8(documentXml)
  let parsed: readonly ParserNode[]
  try {
    SyntaxValidator.validate(xml)
    parsed = new XMLParser(XML_PARSER_OPTIONS).parse(xml) as readonly ParserNode[]
  } catch (error) {
    throw new OfficeExchangeError('conversion-failed', `Cannot read document XML: ${messageOf(error)}`)
  }
  const root = parsed.map(toElement).find(element => element?.tag === 'document')
  if (root === undefined) {
    throw new OfficeExchangeError('conversion-failed', 'Document XML has no w:document element')
  }
  const body = findChild(root, 'body')
  if (body === undefined) {
    throw new OfficeExchangeError('conversion-failed', 'Document XML has no w:body element')
  }
  return body
}

function toElement(node: ParserNode): XmlElement | undefined {
  for (const key of Object.keys(node)) {
    if (key === ':@' || key === '#text') continue
    const children = node[key]
    /* v8 ignore next -- Every non-attribute parser key maps to an element's child array. */
    if (!isParserNodeArray(children)) continue
    const attributes = node[':@']
    return { tag: key, children, attributes: isParserNode(attributes) ? attributes : {} }
  }
  return undefined
}

function childElements(element: XmlElement): readonly XmlElement[] {
  const elements: XmlElement[] = []
  for (const child of element.children) {
    const converted = toElement(child)
    if (converted !== undefined) elements.push(converted)
  }
  return elements
}

function findChild(element: XmlElement, tag: string): XmlElement | undefined {
  for (const child of element.children) {
    const converted = toElement(child)
    if (converted?.tag === tag) return converted
  }
  return undefined
}

function textContent(element: XmlElement): string {
  let text = ''
  for (const child of element.children) {
    const value = child['#text']
    /* v8 ignore next -- A w:t element contains only text children. */
    if (typeof value === 'string') text += value
  }
  return text
}

function attribute(element: XmlElement, name: string): string | undefined {
  const value = element.attributes[name]
  return typeof value === 'string' ? value : undefined
}

function isParserNode(value: unknown): value is ParserNode {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function isParserNodeArray(value: unknown): value is readonly ParserNode[] {
  return Array.isArray(value) && value.every(isParserNode)
}

function positiveNumber(value: string | undefined): number | undefined {
  if (value === undefined) return undefined
  const parsed = Number(value)
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined
}

function documentXml(body: string): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + `<w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>`
    + '</w:document>'
}

function contentTypesXml(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '</Types>'
}

function rootRelsXml(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1"'
    + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
    + ' Target="word/document.xml"/>'
    + '</Relationships>'
}

function documentRelsXml(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1"'
    + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
    + ' Target="styles.xml"/>'
    + '</Relationships>'
}

function stylesXml(): string {
  const headingStyles = HEADING_STYLE_IDS.map((styleId, index) => {
    const level = index + 1
    return `<w:style w:type="paragraph" w:styleId="${styleId}">`
      + `<w:name w:val="heading ${String(level)}"/>`
      + '<w:basedOn w:val="Normal"/>'
      + '<w:next w:val="Normal"/>'
      + '<w:qFormat/>'
      + `<w:pPr><w:outlineLvl w:val="${String(index)}"/></w:pPr>`
      + '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>'
      + '</w:style>'
  }).join('')
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:docDefaults><w:rPrDefault><w:rPr>'
    + '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/>'
    + '</w:rPr></w:rPrDefault></w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    + headingStyles
    + '</w:styles>'
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
