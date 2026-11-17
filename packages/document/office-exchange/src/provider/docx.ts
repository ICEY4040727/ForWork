/**
 * DOCX mapping between OOXML documents and Univer `IDocumentData`.
 *
 * SUPPORTED: paragraph order and paragraph breaks, bold, italic, underline,
 * font family, font size, text color, highlight or shading, strike, vertical
 * alignment, caps, small caps, and run borders; tab characters inside runs;
 * paragraph alignment, spacing, line spacing, indentation, tab stops, outline
 * level, shading, and paragraph borders; document font and paragraph defaults
 * from `word/styles.xml`; named heading paragraph styles (`Heading1` through
 * `Heading5`); and flow tables. Table support covers rows, cells, cell text,
 * grid column widths, preferred cell widths, and the paragraph/section-break
 * pair each cell requires.
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
 * numbering, sections and page layout, theme styles and theme-derived colors,
 * named style definitions other than the heading styles above, cell borders,
 * merged cells with spans, nested tables, floating and inline objects, and
 * property values the Univer model has no field for (for example a run
 * border's line size or a decimal tab stop's alignment). A read/write
 * round-trip does not preserve these. Macro-bearing packages are read without
 * executing or retaining macros.
 */
import { XMLParser } from 'fast-xml-parser'
import { SyntaxValidator } from 'fast-xml-validator'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import type { Unzipped, Zippable } from 'fflate'
import {
  BaselineOffset,
  BorderStyleTypes,
  BooleanNumber,
  DashStyleType,
  HorizontalAlign,
  SpacingRule,
  TabStopAlignment,
  TabStopLeader,
  TextDecoration,
} from '@univerjs/core'
import type {
  IBorderData,
  IBorderStyleData,
  ICustomTable,
  IDocumentData,
  IDocumentStyle,
  IParagraph,
  IParagraphBorder,
  IParagraphStyle,
  ISectionBreak,
  ITable,
  ITableCell,
  ITableCellMargin,
  ITableRow,
  ITables,
  ITabStop,
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
} from '@univerjs/core'
import { type OfficeExchangeConfig, resolveConfig } from '../config.ts'
import { OfficeExchangeError } from '../errors.ts'

const DOCUMENT_XML = 'word/document.xml'
const STYLES_XML = 'word/styles.xml'
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
const PIXELS_PER_POINT = 4 / 3
const POINTS_PER_PIXEL = 0.75
const BORDER_UNITS_PER_PIXEL = 6
const AUTO_LINE_SPACING_UNITS = 240
const HALF_POINTS_PER_POINT = 2
const DEFAULT_BORDER_SIZE = 4
const RUN_BORDER_SIZE = 8
const DEFAULT_BORDER_COLOR = '000000'
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

/** Defaults written when a snapshot declares no document text style. */
const DEFAULT_TEXT_STYLE: ITextStyle = { ff: 'Calibri', fs: 11 }

/** OOXML `w:highlight` names and their RGB values. */
const HIGHLIGHT_COLORS: Readonly<Record<string, string>> = {
  black: '000000',
  blue: '0000FF',
  cyan: '00FFFF',
  darkBlue: '000080',
  darkCyan: '008080',
  darkGray: '808080',
  darkGreen: '008000',
  darkMagenta: '800080',
  darkRed: '800000',
  darkYellow: '808000',
  green: '00FF00',
  lightGray: 'C0C0C0',
  magenta: 'FF00FF',
  red: 'FF0000',
  white: 'FFFFFF',
  yellow: 'FFFF00',
}

/** OOXML `w:u` values and their Univer text decoration. */
const UNDERLINE_VALUES: readonly (readonly [string, TextDecoration])[] = [
  ['single', TextDecoration.SINGLE],
  ['words', TextDecoration.WORDS],
  ['double', TextDecoration.DOUBLE],
  ['thick', TextDecoration.THICK],
  ['dotted', TextDecoration.DOTTED],
  ['dottedHeavy', TextDecoration.DOTTED_HEAVY],
  ['dash', TextDecoration.DASH],
  ['dashedHeavy', TextDecoration.DASHED_HEAVY],
  ['dashLong', TextDecoration.DASH_LONG],
  ['dashLongHeavy', TextDecoration.DASH_LONG_HEAVY],
  ['dotDash', TextDecoration.DOT_DASH],
  ['dotDotDash', TextDecoration.DOT_DOT_DASH],
  ['dashDotHeavy', TextDecoration.DASH_DOT_HEAVY],
  ['dashDotDotHeavy', TextDecoration.DASH_DOT_DOT_HEAVY],
  ['wave', TextDecoration.WAVE],
  ['wavyDouble', TextDecoration.WAVY_DOUBLE],
  ['wavyHeavy', TextDecoration.WAVY_HEAVY],
  ['singleAccounting', TextDecoration.SINGLE_ACCOUNTING],
  ['doubleAccounting', TextDecoration.DOUBLE_ACCOUNTING],
]

/** OOXML border values and their Univer paragraph-border dash style. */
const BORDER_DASH_STYLES: readonly (readonly [string, DashStyleType])[] = [
  ['single', DashStyleType.SOLID],
  ['thick', DashStyleType.SOLID],
  ['double', DashStyleType.SOLID],
  ['medium', DashStyleType.SOLID],
  ['dotted', DashStyleType.DOT],
  ['hair', DashStyleType.DOT],
  ['dashed', DashStyleType.DASH],
  ['dashDot', DashStyleType.DASH],
  ['dashDotDot', DashStyleType.DASH],
  ['dashLong', DashStyleType.DASH],
  ['dashLongHeavy', DashStyleType.DASH],
  ['mediumDashed', DashStyleType.DASH],
  ['mediumDashDot', DashStyleType.DASH],
  ['mediumDashDotDot', DashStyleType.DASH],
  ['slantDashDot', DashStyleType.DASH],
]

/** OOXML border values and the Univer run-border weight that stands in for them. */
const BORDER_WEIGHTS: readonly (readonly [string, BorderStyleTypes])[] = [
  ['single', BorderStyleTypes.THIN],
  ['hair', BorderStyleTypes.HAIR],
  ['dotted', BorderStyleTypes.DOTTED],
  ['dashed', BorderStyleTypes.DASHED],
  ['dashDot', BorderStyleTypes.DASH_DOT],
  ['dashDotDot', BorderStyleTypes.DASH_DOT_DOT],
  ['double', BorderStyleTypes.DOUBLE],
  ['medium', BorderStyleTypes.MEDIUM],
  ['mediumDashed', BorderStyleTypes.MEDIUM_DASHED],
  ['mediumDashDot', BorderStyleTypes.MEDIUM_DASH_DOT],
  ['mediumDashDotDot', BorderStyleTypes.MEDIUM_DASH_DOT_DOT],
  ['slantDashDot', BorderStyleTypes.SLANT_DASH_DOT],
  ['thick', BorderStyleTypes.THICK],
  ['threeDEmboss', BorderStyleTypes.MEDIUM],
  ['threeDEngrave', BorderStyleTypes.MEDIUM],
  ['inset', BorderStyleTypes.MEDIUM],
  ['outset', BorderStyleTypes.MEDIUM],
]

/** OOXML `w:jc` values and their Univer horizontal alignment. */
const HORIZONTAL_ALIGNMENTS: readonly (readonly [string, HorizontalAlign])[] = [
  ['left', HorizontalAlign.LEFT],
  ['start', HorizontalAlign.LEFT],
  ['center', HorizontalAlign.CENTER],
  ['right', HorizontalAlign.RIGHT],
  ['end', HorizontalAlign.RIGHT],
  ['both', HorizontalAlign.BOTH],
  ['justified', HorizontalAlign.JUSTIFIED],
  ['distribute', HorizontalAlign.DISTRIBUTED],
]

/** OOXML `w:tab` alignment values and their Univer tab stop alignment. */
const TAB_STOP_ALIGNMENTS: readonly (readonly [string, TabStopAlignment])[] = [
  ['left', TabStopAlignment.START],
  ['start', TabStopAlignment.START],
  ['center', TabStopAlignment.CENTER],
  ['right', TabStopAlignment.END],
  ['end', TabStopAlignment.END],
]

/** OOXML `w:tab` leader values and their Univer tab stop leader. */
const TAB_STOP_LEADERS: readonly (readonly [string, TabStopLeader])[] = [
  ['none', TabStopLeader.NONE],
  ['dot', TabStopLeader.DOT],
  ['hyphen', TabStopLeader.HYPHEN],
  ['underscore', TabStopLeader.UNDERSCORE],
  ['heavy', TabStopLeader.HEAVY],
  ['middleDot', TabStopLeader.MIDDLE_DOT],
]

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
 * breaks, run and paragraph formatting, `word/styles.xml` defaults, heading
 * styles, and flow tables are mapped; the unsupported features listed in the
 * module note are not.
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
  const parts = readPackage(bytes)
  if (parts.document === undefined) {
    throw new OfficeExchangeError('conversion-failed', 'Package has no word/document.xml part')
  }
  const body = parseDocumentBody(parts.document)
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
    documentStyle: readDocumentStyle(parts.styles),
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
    [STYLES_XML]: strToU8(stylesXml(data.documentStyle)),
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

function appendParagraphBreak(context: ReadContext, styleId?: string, formatting?: IParagraphStyle): void {
  append(context, PARAGRAPH)
  const paragraph: IParagraph = {
    startIndex: context.length - 1,
    paragraphId: `paragraph-${context.paragraphs.length + 1}`,
  }
  if (styleId !== undefined) paragraph.styleId = styleId
  const namedStyle = styleId === undefined ? undefined : NAMED_STYLE_BY_ID[styleId]
  if (namedStyle !== undefined) paragraph.paragraphStyle = { ...formatting, namedStyleType: namedStyle }
  else if (formatting !== undefined) paragraph.paragraphStyle = formatting
  context.paragraphs.push(paragraph)
}

function appendParagraph(context: ReadContext, element: XmlElement): void {
  const properties = findChild(element, 'pPr')
  const style = properties === undefined ? undefined : findChild(properties, 'pStyle')
  const styleId = style === undefined ? undefined : attribute(style, '@_val')
  const formatting = properties === undefined ? undefined : readParagraphStyle(properties)
  for (const child of childElements(element)) {
    if (child.tag !== 'r') continue
    appendRun(context, child)
  }
  appendParagraphBreak(context, styleId, formatting)
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
  if (isEnabled(properties, 'caps') === true) style.caps = true
  if (isEnabled(properties, 'smallCaps') === true) style.smallCaps = true
  const strike = readStrike(properties)
  if (strike !== undefined) style.st = strike
  const fonts = findChild(properties, 'rFonts')
  if (fonts !== undefined) {
    const latin = attribute(fonts, '@_ascii') ?? attribute(fonts, '@_hAnsi') ?? attribute(fonts, '@_cs')
    if (latin !== undefined) style.ff = latin
    const eastAsia = attribute(fonts, '@_eastAsia')
    if (eastAsia !== undefined) style.eastAsiaFontFamily = eastAsia
  }
  const size = readFontSize(properties)
  if (size !== undefined) style.fs = size
  const color = rgbOf(findChild(properties, 'color'))
  if (color !== undefined) style.cl = { rgb: color }
  const background = readHighlight(properties) ?? shadingColor(findChild(properties, 'shd'))
  if (background !== undefined) style.bg = { rgb: background }
  const underline = readUnderline(properties)
  if (underline !== undefined) style.ul = underline
  const vertical = readVerticalAlign(properties)
  if (vertical !== undefined) style.va = vertical
  const border = readRunBorder(properties)
  if (border !== undefined) style.bd = border
  return Object.keys(style).length === 0 ? undefined : style
}

function readParagraphStyle(properties: XmlElement): IParagraphStyle | undefined {
  const style: IParagraphStyle = {}
  const alignment = optionalAttribute(findChild(properties, 'jc'), '@_val')
  const horizontal = alignment === undefined ? undefined : lookup(HORIZONTAL_ALIGNMENTS, alignment)
  if (horizontal !== undefined) style.horizontalAlign = horizontal
  const spacing = findChild(properties, 'spacing')
  if (spacing !== undefined) readSpacing(spacing, style)
  const indent = findChild(properties, 'ind')
  if (indent !== undefined) readIndent(indent, style)
  const tabs = findChild(properties, 'tabs')
  if (tabs !== undefined) readTabStops(tabs, style)
  const borders = findChild(properties, 'pBdr')
  if (borders !== undefined) readParagraphBorders(borders, style)
  const outline = numberAttribute(findChild(properties, 'outlineLvl'), '@_val')
  if (outline !== undefined) style.outlineLevel = outline
  const shading = shadingColor(findChild(properties, 'shd'))
  if (shading !== undefined) style.shading = { backgroundColor: { rgb: shading } }
  return Object.keys(style).length === 0 ? undefined : style
}

function readSpacing(element: XmlElement, style: IParagraphStyle): void {
  const before = numberAttribute(element, '@_before')
  if (before !== undefined) style.spaceAbove = { v: before / TWIPS_PER_PIXEL }
  const after = numberAttribute(element, '@_after')
  if (after !== undefined) style.spaceBelow = { v: after / TWIPS_PER_PIXEL }
  const line = numberAttribute(element, '@_line')
  if (line === undefined) return
  const rule = attribute(element, '@_lineRule')
  if (rule === undefined || rule === 'auto') {
    style.lineSpacing = line / AUTO_LINE_SPACING_UNITS
    style.spacingRule = SpacingRule.AUTO
    return
  }
  style.lineSpacing = line / TWIPS_PER_PIXEL
  style.spacingRule = rule === 'atLeast' ? SpacingRule.AT_LEAST : SpacingRule.EXACT
}

function readIndent(element: XmlElement, style: IParagraphStyle): void {
  const start = numberAttribute(element, '@_start') ?? numberAttribute(element, '@_left')
  if (start !== undefined) style.indentStart = { v: start / TWIPS_PER_PIXEL }
  const end = numberAttribute(element, '@_end') ?? numberAttribute(element, '@_right')
  if (end !== undefined) style.indentEnd = { v: end / TWIPS_PER_PIXEL }
  const firstLine = numberAttribute(element, '@_firstLine')
  if (firstLine !== undefined) style.indentFirstLine = { v: firstLine / TWIPS_PER_PIXEL }
  const hanging = numberAttribute(element, '@_hanging')
  if (hanging !== undefined) style.hanging = { v: hanging / TWIPS_PER_PIXEL }
}

function readTabStops(element: XmlElement, style: IParagraphStyle): void {
  const stops: ITabStop[] = []
  for (const child of childElements(element)) {
    if (child.tag !== 'tab') continue
    const stop = readTabStop(child)
    if (stop !== undefined) stops.push(stop)
  }
  if (stops.length > 0) style.tabStops = stops
}

function readTabStop(element: XmlElement): ITabStop | undefined {
  const position = numberAttribute(element, '@_pos')
  if (position === undefined) return undefined
  const value = attribute(element, '@_val')
  const stop: ITabStop = {
    offset: position / TWIPS_PER_PIXEL,
    alignment: (value === undefined ? undefined : lookup(TAB_STOP_ALIGNMENTS, value)) ?? TabStopAlignment.START,
  }
  if (value === 'clear') stop.clear = true
  const leaderName = attribute(element, '@_leader')
  const leader = leaderName === undefined ? undefined : lookup(TAB_STOP_LEADERS, leaderName)
  if (leader !== undefined) stop.leader = leader
  return stop
}

function readParagraphBorders(element: XmlElement, style: IParagraphStyle): void {
  const top = readParagraphBorder(findChild(element, 'top'))
  if (top !== undefined) style.borderTop = top
  const bottom = readParagraphBorder(findChild(element, 'bottom'))
  if (bottom !== undefined) style.borderBottom = bottom
  const left = readParagraphBorder(findChild(element, 'left'))
  if (left !== undefined) style.borderLeft = left
  const right = readParagraphBorder(findChild(element, 'right'))
  if (right !== undefined) style.borderRight = right
  const between = readParagraphBorder(findChild(element, 'between'))
  if (between !== undefined) style.borderBetween = between
}

function readParagraphBorder(element: XmlElement | undefined): IParagraphBorder | undefined {
  if (element === undefined) return undefined
  const value = attribute(element, '@_val')
  if (value === undefined || value === 'none') return undefined
  const size = positiveNumber(attribute(element, '@_sz'))
  const space = numberAttribute(element, '@_space')
  const color = hexColor(attribute(element, '@_color')) ?? DEFAULT_BORDER_COLOR
  return {
    color: { rgb: `#${color}` },
    width: (size ?? DEFAULT_BORDER_SIZE) / BORDER_UNITS_PER_PIXEL,
    dashStyle: lookup(BORDER_DASH_STYLES, value) ?? DashStyleType.SOLID,
    padding: (space ?? 0) * PIXELS_PER_POINT,
  }
}

function readUnderline(properties: XmlElement): ITextDecoration | undefined {
  const underline = findChild(properties, 'u')
  if (underline === undefined) return undefined
  const value = attribute(underline, '@_val')
  if (value === 'none') return undefined
  const decoration = value === undefined ? undefined : lookup(UNDERLINE_VALUES, value)
  const style: ITextDecoration = {
    s: BOOLEAN_TRUE,
    t: decoration ?? TEXT_DECORATION_SINGLE,
  }
  const color = hexColor(attribute(underline, '@_color'))
  if (color !== undefined) style.cl = { rgb: `#${color}` }
  return style
}

function readStrike(properties: XmlElement): ITextDecoration | undefined {
  if (isEnabled(properties, 'dstrike') === true) {
    return { s: BOOLEAN_TRUE, t: TEXT_DECORATION_DOUBLE }
  }
  if (isEnabled(properties, 'strike') === true) {
    return { s: BOOLEAN_TRUE, t: TEXT_DECORATION_SINGLE }
  }
  return undefined
}

function readFontSize(properties: XmlElement): number | undefined {
  const size = findChild(properties, 'sz') ?? findChild(properties, 'szCs')
  if (size === undefined) return undefined
  const halfPoints = positiveNumber(attribute(size, '@_val'))
  return halfPoints === undefined ? undefined : halfPoints / HALF_POINTS_PER_POINT
}

function readHighlight(properties: XmlElement): string | undefined {
  const highlight = findChild(properties, 'highlight')
  if (highlight === undefined) return undefined
  const value = attribute(highlight, '@_val')
  if (value === undefined || value === 'none') return undefined
  const color = HIGHLIGHT_COLORS[value]
  return color === undefined ? undefined : `#${color}`
}

function readVerticalAlign(properties: XmlElement): BaselineOffset | undefined {
  const value = optionalAttribute(findChild(properties, 'vertAlign'), '@_val')
  if (value === 'superscript') return BaselineOffset.SUPERSCRIPT
  if (value === 'subscript') return BaselineOffset.SUBSCRIPT
  if (value === 'baseline') return BaselineOffset.NORMAL
  return undefined
}

function readRunBorder(properties: XmlElement): IBorderData | undefined {
  const border = findChild(properties, 'bdr')
  if (border === undefined) return undefined
  const side = readBorderStyle(border)
  return side === undefined ? undefined : { t: side, r: side, b: side, l: side }
}

function readBorderStyle(element: XmlElement): IBorderStyleData | undefined {
  const value = attribute(element, '@_val')
  if (value === undefined || value === 'none') return undefined
  const color = hexColor(attribute(element, '@_color')) ?? DEFAULT_BORDER_COLOR
  return {
    s: lookup(BORDER_WEIGHTS, value) ?? BorderStyleTypes.THIN,
    cl: { rgb: `#${color}` },
  }
}

function isEnabled(properties: XmlElement, tag: string): boolean | undefined {
  const element = findChild(properties, tag)
  if (element === undefined) return undefined
  const value = attribute(element, '@_val')
  return value === undefined || (value !== '0' && value !== 'false' && value !== 'off')
}

/**
 * Read the document defaults declared by `word/styles.xml`.
 * @param stylesXml - Decoded styles part, when the package contains one.
 * @returns Document-level text and paragraph defaults; empty when the part is absent.
 * @throws {OfficeExchangeError} `conversion-failed` when the styles XML is not well formed.
 */
function readDocumentStyle(stylesXml: Uint8Array | undefined): IDocumentStyle {
  if (stylesXml === undefined) return {}
  const root = parseStylesDocument(stylesXml)
  if (root === undefined) return {}
  const defaults = findChild(root, 'docDefaults')
  const defaultText = defaults === undefined ? undefined : defaultsTextStyle(defaults)
  const defaultParagraph = defaults === undefined ? undefined : defaultsParagraphStyle(defaults)
  const normal = findNormalStyle(root)
  const textStyle = mergeStyles(
    defaultText,
    normal === undefined ? undefined : readElementProperties(normal, 'rPr', readTextStyle),
  )
  const paragraphStyle = mergeStyles(
    defaultParagraph,
    normal === undefined ? undefined : readElementProperties(normal, 'pPr', readParagraphStyle),
  )
  const documentStyle: IDocumentStyle = {}
  if (textStyle !== undefined) documentStyle.textStyle = textStyle
  if (paragraphStyle !== undefined) documentStyle.defaultParagraphStyle = paragraphStyle
  return documentStyle
}

function defaultsTextStyle(defaults: XmlElement): ITextStyle | undefined {
  const container = findChild(defaults, 'rPrDefault')
  return container === undefined ? undefined : readElementProperties(container, 'rPr', readTextStyle)
}

function defaultsParagraphStyle(defaults: XmlElement): IParagraphStyle | undefined {
  const container = findChild(defaults, 'pPrDefault')
  return container === undefined ? undefined : readElementProperties(container, 'pPr', readParagraphStyle)
}

function readElementProperties<Properties>(
  element: XmlElement,
  tag: string,
  read: (properties: XmlElement) => Properties | undefined,
): Properties | undefined {
  const properties = findChild(element, tag)
  return properties === undefined ? undefined : read(properties)
}

function mergeStyles<Value extends object>(
  base: Value | undefined,
  override: Value | undefined,
): Value | undefined {
  if (base === undefined) return override
  if (override === undefined) return base
  return { ...base, ...override }
}

function findNormalStyle(root: XmlElement): XmlElement | undefined {
  for (const child of childElements(root)) {
    if (child.tag !== 'style') continue
    if (attribute(child, '@_type') !== 'paragraph') continue
    if (attribute(child, '@_default') === '1' || attribute(child, '@_styleId') === 'Normal') return child
  }
  return undefined
}

function parseStylesDocument(stylesXml: Uint8Array): XmlElement | undefined {
  const xml = strFromU8(stylesXml)
  let parsed: readonly ParserNode[]
  try {
    SyntaxValidator.validate(xml)
    parsed = new XMLParser(XML_PARSER_OPTIONS).parse(xml) as readonly ParserNode[]
  } catch (error) {
    throw new OfficeExchangeError('conversion-failed', `Cannot read styles XML: ${messageOf(error)}`)
  }
  return parsed.map(toElement).find(element => element?.tag === 'styles')
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
  const formatting = paragraph?.paragraphStyle === undefined
    ? ''
    : writeParagraphFormatting(paragraph.paragraphStyle)
  if (styleId === undefined && formatting.length === 0) return ''
  const named = styleId === undefined ? '' : `<w:pStyle w:val="${escapeXml(styleId)}"/>`
  return `<w:pPr>${named}${formatting}</w:pPr>`
}

function writeParagraphFormatting(style: IParagraphStyle): string {
  let properties = ''
  if (style.horizontalAlign !== undefined) {
    properties += `<w:jc w:val="${horizontalAlignName(style.horizontalAlign)}"/>`
  }
  properties += writeParagraphSpacing(style)
  properties += writeParagraphIndent(style)
  properties += writeTabStops(style.tabStops)
  properties += writeParagraphBorders(style)
  const shading = style.shading === undefined ? undefined : hexColor(style.shading.backgroundColor.rgb)
  if (shading !== undefined) properties += `<w:shd w:val="clear" w:color="auto" w:fill="${shading}"/>`
  if (style.outlineLevel !== undefined) properties += `<w:outlineLvl w:val="${String(style.outlineLevel)}"/>`
  return properties
}

function horizontalAlignName(alignment: HorizontalAlign): string {
  if (alignment === HorizontalAlign.CENTER) return 'center'
  if (alignment === HorizontalAlign.RIGHT) return 'right'
  if (alignment === HorizontalAlign.BOTH || alignment === HorizontalAlign.JUSTIFIED) return 'both'
  if (alignment === HorizontalAlign.DISTRIBUTED) return 'distribute'
  return 'left'
}

function writeParagraphSpacing(style: IParagraphStyle): string {
  const attributes: string[] = []
  if (style.spaceAbove !== undefined) {
    attributes.push(`w:before="${String(Math.round(style.spaceAbove.v * TWIPS_PER_PIXEL))}"`)
  }
  if (style.spaceBelow !== undefined) {
    attributes.push(`w:after="${String(Math.round(style.spaceBelow.v * TWIPS_PER_PIXEL))}"`)
  }
  if (style.lineSpacing !== undefined) {
    if (style.spacingRule === SpacingRule.AT_LEAST || style.spacingRule === SpacingRule.EXACT) {
      attributes.push(`w:line="${String(Math.round(style.lineSpacing * TWIPS_PER_PIXEL))}"`)
      attributes.push(`w:lineRule="${style.spacingRule === SpacingRule.AT_LEAST ? 'atLeast' : 'exact'}"`)
    } else {
      attributes.push(`w:line="${String(Math.round(style.lineSpacing * AUTO_LINE_SPACING_UNITS))}"`)
      attributes.push('w:lineRule="auto"')
    }
  }
  return attributes.length === 0 ? '' : `<w:spacing ${attributes.join(' ')}/>`
}

function writeParagraphIndent(style: IParagraphStyle): string {
  const attributes: string[] = []
  if (style.indentStart !== undefined) attributes.push(`w:left="${twipsOf(style.indentStart.v)}"`)
  if (style.indentEnd !== undefined) attributes.push(`w:right="${twipsOf(style.indentEnd.v)}"`)
  if (style.indentFirstLine !== undefined) attributes.push(`w:firstLine="${twipsOf(style.indentFirstLine.v)}"`)
  if (style.hanging !== undefined) attributes.push(`w:hanging="${twipsOf(style.hanging.v)}"`)
  return attributes.length === 0 ? '' : `<w:ind ${attributes.join(' ')}/>`
}

function writeTabStops(stops: readonly ITabStop[] | undefined): string {
  if (stops === undefined || stops.length === 0) return ''
  let xml = ''
  for (const stop of stops) {
    const value = stop.clear === true
      ? 'clear'
      : nameOf(TAB_STOP_ALIGNMENTS, stop.alignment) ?? 'left'
    xml += `<w:tab w:val="${value}" w:pos="${twipsOf(stop.offset)}"`
    const leader = stop.leader === undefined ? undefined : nameOf(TAB_STOP_LEADERS, stop.leader)
    if (leader !== undefined) xml += ` w:leader="${leader}"`
    xml += '/>'
  }
  return `<w:tabs>${xml}</w:tabs>`
}

function writeParagraphBorders(style: IParagraphStyle): string {
  let xml = ''
  xml += writeParagraphBorder('top', style.borderTop)
  xml += writeParagraphBorder('bottom', style.borderBottom)
  xml += writeParagraphBorder('left', style.borderLeft)
  xml += writeParagraphBorder('right', style.borderRight)
  xml += writeParagraphBorder('between', style.borderBetween)
  return xml.length === 0 ? '' : `<w:pBdr>${xml}</w:pBdr>`
}

function writeParagraphBorder(tag: string, border: IParagraphBorder | undefined): string {
  if (border === undefined) return ''
  const color = hexColor(border.color.rgb) ?? DEFAULT_BORDER_COLOR
  const dash = nameOf(BORDER_DASH_STYLES, border.dashStyle) ?? 'single'
  const size = Math.max(1, Math.round(border.width * BORDER_UNITS_PER_PIXEL))
  const space = Math.max(0, Math.round(border.padding * POINTS_PER_PIXEL))
  return `<w:${tag} w:val="${dash}" w:sz="${String(size)}" w:space="${String(space)}" w:color="${color}"/>`
}

function twipsOf(pixels: number): string {
  return String(Math.round(pixels * TWIPS_PER_PIXEL))
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
  const family = style.ff ?? undefined
  const eastAsia = style.eastAsiaFontFamily ?? undefined
  if (family !== undefined || eastAsia !== undefined) {
    properties += '<w:rFonts'
    if (family !== undefined) properties += ` w:ascii="${escapeXml(family)}" w:hAnsi="${escapeXml(family)}"`
    if (eastAsia !== undefined) properties += ` w:eastAsia="${escapeXml(eastAsia)}"`
    properties += '/>'
  }
  if (style.bl === BOOLEAN_TRUE) properties += '<w:b/>'
  if (style.it === BOOLEAN_TRUE) properties += '<w:i/>'
  if (style.caps === true) properties += '<w:caps/>'
  if (style.smallCaps === true) properties += '<w:smallCaps/>'
  if (style.st?.s === BOOLEAN_TRUE) {
    properties += style.st.t === TEXT_DECORATION_DOUBLE ? '<w:dstrike/>' : '<w:strike/>'
  }
  const foreground = hexColor(style.cl?.rgb)
  if (foreground !== undefined) properties += `<w:color w:val="${foreground}"/>`
  if (style.fs !== undefined) {
    const halfPoints = String(Math.round(style.fs * HALF_POINTS_PER_POINT))
    properties += `<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>`
  }
  const background = hexColor(style.bg?.rgb)
  if (background !== undefined) {
    const highlight = highlightName(background)
    properties += highlight === undefined
      ? `<w:shd w:val="clear" w:color="auto" w:fill="${background}"/>`
      : `<w:highlight w:val="${highlight}"/>`
  }
  if (style.ul?.s === BOOLEAN_TRUE) {
    properties += `<w:u w:val="${underlineName(style.ul.t)}"`
    const underlineColor = hexColor(style.ul.cl?.rgb)
    if (underlineColor !== undefined) properties += ` w:color="${underlineColor}"`
    properties += '/>'
  }
  const vertical = style.va === null || style.va === undefined ? undefined : verticalAlignName(style.va)
  if (vertical !== undefined) properties += `<w:vertAlign w:val="${vertical}"/>`
  const border = style.bd?.t ?? style.bd?.r ?? style.bd?.b ?? style.bd?.l
  if (border !== null && border !== undefined) {
    const color = hexColor(border.cl?.rgb) ?? DEFAULT_BORDER_COLOR
    properties += `<w:bdr w:val="${nameOf(BORDER_WEIGHTS, border.s) ?? 'single'}"`
      + ` w:sz="${String(RUN_BORDER_SIZE)}" w:space="1" w:color="${color}"/>`
  }
  return properties.length === 0 ? '' : `<w:rPr>${properties}</w:rPr>`
}

function underlineName(decoration: TextDecoration | undefined): string {
  return decoration === undefined ? 'single' : nameOf(UNDERLINE_VALUES, decoration) ?? 'single'
}

function verticalAlignName(vertical: BaselineOffset): string | undefined {
  if (vertical === BaselineOffset.SUPERSCRIPT) return 'superscript'
  if (vertical === BaselineOffset.SUBSCRIPT) return 'subscript'
  if (vertical === BaselineOffset.NORMAL) return 'baseline'
  return undefined
}

function highlightName(color: string): string | undefined {
  for (const [name, value] of Object.entries(HIGHLIGHT_COLORS)) {
    if (value === color) return name
  }
  return undefined
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

/** XML parts a document read needs from the package. */
interface PackageParts {
  /** Decoded `word/document.xml`, when present. */
  readonly document: Uint8Array | undefined
  /** Decoded `word/styles.xml`, when present. */
  readonly styles: Uint8Array | undefined
}

function readPackage(bytes: Uint8Array): PackageParts {
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
        return file.name === DOCUMENT_XML || file.name === STYLES_XML
      },
    })
  } catch (error) {
    if (error instanceof OfficeExchangeError) throw error
    throw new OfficeExchangeError('conversion-failed', `Cannot read OOXML package: ${messageOf(error)}`)
  }
  return { document: entries[DOCUMENT_XML], styles: entries[STYLES_XML] }
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

function optionalAttribute(element: XmlElement | undefined, name: string): string | undefined {
  return element === undefined ? undefined : attribute(element, name)
}

function numberAttribute(element: XmlElement | undefined, name: string): number | undefined {
  const value = optionalAttribute(element, name)
  if (value === undefined) return undefined
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : undefined
}

function lookup<Value>(table: readonly (readonly [string, Value])[], name: string): Value | undefined {
  for (const [key, value] of table) {
    if (key === name) return value
  }
  return undefined
}

function nameOf<Value>(table: readonly (readonly [string, Value])[], value: Value): string | undefined {
  for (const [key, candidate] of table) {
    if (candidate === value) return key
  }
  return undefined
}

function hexColor(value: string | null | void): string | undefined {
  if (typeof value !== 'string') return undefined
  const hex = value.startsWith('#') ? value.slice(1) : value
  return /^[0-9A-Fa-f]{6}$/u.test(hex) ? hex.toUpperCase() : undefined
}

function rgbOf(element: XmlElement | undefined): string | undefined {
  const color = hexColor(optionalAttribute(element, '@_val'))
  return color === undefined ? undefined : `#${color}`
}

function shadingColor(element: XmlElement | undefined): string | undefined {
  const color = hexColor(optionalAttribute(element, '@_fill'))
  return color === undefined ? undefined : `#${color}`
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

function stylesXml(documentStyle: IDocumentStyle): string {
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
  const textStyle = writeRunProperties(documentStyle.textStyle ?? DEFAULT_TEXT_STYLE)
  const paragraphStyle = documentStyle.defaultParagraphStyle === undefined
    ? ''
    : writeParagraphFormatting(documentStyle.defaultParagraphStyle)
  const paragraphDefault = paragraphStyle.length === 0
    ? ''
    : `<w:pPrDefault><w:pPr>${paragraphStyle}</w:pPr></w:pPrDefault>`
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + `<w:docDefaults><w:rPrDefault>${textStyle}</w:rPrDefault>${paragraphDefault}</w:docDefaults>`
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    + headingStyles
    + '</w:styles>'
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
