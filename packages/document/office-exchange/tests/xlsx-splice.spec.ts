/** In-place XLSX save: cell span splicing, row/column shifts, and refusals. */
import {
  BaselineOffset,
  BooleanNumber,
  BorderStyleTypes,
  HorizontalAlign,
  TextDecoration,
  VerticalAlign,
  type ICellData,
  type IStyleData,
  type IWorksheetData,
} from '@univerjs/core'
import { strFromU8, strToU8, zipSync } from 'fflate'
import { describe, expect, it, vi } from 'vitest'
import { resolveConfig } from '../src/config.ts'
import { OfficeExchangeError } from '../src/errors.ts'
import { readWorkbook, readWorkbookWithSource, writeWorkbookFromSource } from '../src/provider/xlsx.ts'
import { OfficePackage, open } from '../src/package/ooxml-package.ts'
import type { OfficeSourceState, UnitAnchor } from '../src/source/source-state.ts'
import type { OfficeWriteResult } from '../src/snapshots.ts'
import { comparePackages, formatDiffReport, unzipEntries, validateOpc } from './ooxml-package.ts'

const SPREADSHEET = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
const SHEET = 'xl/worksheets/sheet1.xml'
const SHARED_STRINGS = 'xl/sharedStrings.xml'
const STYLES = 'xl/styles.xml'
const WORKBOOK_RELS = 'xl/_rels/workbook.xml.rels'

const DEFAULT_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${SPREADSHEET}">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`

interface XlsxFixture {
  readonly sheetData: string
  readonly dimension?: string
  readonly extra?: string
  readonly sharedStrings?: readonly string[]
  readonly richSharedStrings?: readonly string[]
  readonly styles?: string
  readonly includeStyles?: boolean
  readonly stylesUnreferenced?: string
  readonly sharedStringsUnreferenced?: string
  readonly sheetRels?: string
  readonly workbookSheets?: string
  readonly workbookRelsExtra?: string
  readonly entriesExtra?: Readonly<Record<string, Uint8Array>>
  readonly contentTypesExtra?: string
}

function xlsxPackage(fixture: XlsxFixture): Uint8Array {
  const sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<worksheet xmlns="${SPREADSHEET}" xmlns:r="${REL}">`
    + (fixture.dimension === undefined ? '' : `<dimension ref="${fixture.dimension}"/>`)
    + `<sheetData>${fixture.sheetData}</sheetData>`
    + (fixture.extra ?? '')
    + '</worksheet>'
  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(contentTypes(fixture)),
    '_rels/.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument"`
      + ' Target="xl/workbook.xml"/></Relationships>',
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<workbook xmlns="${SPREADSHEET}" xmlns:r="${REL}">`
      + `<sheets>${fixture.workbookSheets ?? '<sheet name="Data" sheetId="1" r:id="rId1"/>'}</sheets>`
      + '</workbook>',
    ),
    [WORKBOOK_RELS]: strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<Relationships xmlns="${PKG_REL}">`
      + `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`
      + (fixture.includeStyles === false ? '' : `<Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/>`)
      + (fixture.sharedStrings === undefined && fixture.richSharedStrings === undefined
        ? ''
        : `<Relationship Id="rId3" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`)
      + (fixture.workbookRelsExtra ?? '')
      + '</Relationships>',
    ),
    [SHEET]: strToU8(sheet),
    ...(fixture.stylesUnreferenced !== undefined
      ? { [STYLES]: strToU8(fixture.stylesUnreferenced) }
      : fixture.includeStyles === false
        ? {}
        : { [STYLES]: strToU8(fixture.styles ?? DEFAULT_STYLES) }),
    ...(fixture.sharedStringsUnreferenced === undefined
      ? {}
      : { 'xl/sharedStrings.xml': strToU8(fixture.sharedStringsUnreferenced) }),
    ...(fixture.entriesExtra ?? {}),
  }
  if (fixture.sharedStrings !== undefined || fixture.richSharedStrings !== undefined) {
    entries[SHARED_STRINGS] = strToU8(sharedStringsXml(fixture))
  }
  if (fixture.sheetRels !== undefined) {
    entries['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<Relationships xmlns="${PKG_REL}">${fixture.sheetRels}</Relationships>`,
    )
  }
  return zipSync(entries, { level: 6 })
}

function contentTypes(fixture: XlsxFixture): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<Types xmlns="${CT}">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + '<Override PartName="/xl/worksheets/sheet1.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    + (fixture.includeStyles === false
      ? ''
      : '<Override PartName="/xl/styles.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    + (fixture.sharedStrings === undefined && fixture.richSharedStrings === undefined
      ? ''
      : '<Override PartName="/xl/sharedStrings.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>')
    + (fixture.contentTypesExtra ?? '')
    + '</Types>'
}

function sharedStringsXml(fixture: XlsxFixture): string {
  const plain = fixture.sharedStrings ?? []
  const rich = fixture.richSharedStrings ?? []
  const entries = [
    ...plain.map(text => `<si><t>${text}</t></si>`),
    ...rich,
  ]
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<sst xmlns="${SPREADSHEET}" count="${String(entries.length)}" uniqueCount="${String(entries.length)}">`
    + entries.join('')
    + '</sst>'
}

interface SidecarCarrier {
  officeExchange?: {
    readonly source?: unknown
  }
}

function styleOf(cell: ICellData): IStyleData {
  return typeof cell.s === 'object' && cell.s !== null ? cell.s : {}
}

function sourceStateOf(data: object): OfficeSourceState | undefined {
  return (data as SidecarCarrier).officeExchange?.source as OfficeSourceState | undefined
}

function replaceSourceState(data: object, state: OfficeSourceState): void {
  ;(data as SidecarCarrier).officeExchange = { source: state }
}

function withSourceState(data: object, mutate: (state: OfficeSourceState) => OfficeSourceState): void {
  const state = sourceStateOf(data)
  if (state === undefined) throw new Error('fixture has no source state')
  replaceSourceState(data, mutate(state))
}

function corruptEntryCrc(source: Uint8Array, name: string): Uint8Array {
  const bytes = source.slice()
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let eocd = bytes.byteLength - 22
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x0605_4b50) eocd -= 1
  if (eocd < 0) throw new Error('package has no end-of-central-directory record')
  let cursor = view.getUint32(eocd + 16, true)
  const count = view.getUint16(eocd + 10, true)
  for (let index = 0; index < count; index += 1) {
    if (view.getUint32(cursor, true) !== 0x0201_4b50) throw new Error('malformed central directory')
    const nameLength = view.getUint16(cursor + 28, true)
    const extraLength = view.getUint16(cursor + 30, true)
    const commentLength = view.getUint16(cursor + 32, true)
    const entryName = new TextDecoder().decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength))
    if (entryName === name) {
      const localOffset = view.getUint32(cursor + 42, true)
      view.setUint32(localOffset + 14, view.getUint32(localOffset + 14, true) ^ 0xff, true)
      view.setUint32(cursor + 16, view.getUint32(cursor + 16, true) ^ 0xff, true)
      return bytes
    }
    cursor += 46 + nameLength + extraLength + commentLength
  }
  throw new Error(`package has no entry ${name}`)
}

function anchorsOf(data: object): readonly UnitAnchor[] {
  const state = sourceStateOf(data)
  if (state?.anchorMode !== 'per-unit') throw new Error('fixture has no per-unit source anchors')
  return state.anchors
}

function anchorOf(data: object, key: string): UnitAnchor {
  const anchor = anchorsOf(data).find(candidate => candidate.key === key)
  if (anchor === undefined) throw new Error(`fixture has no anchor ${key}`)
  return anchor
}

function sheetOf(data: { sheets: Record<string, Partial<IWorksheetData>> }, id = 'sheet-1'): Partial<IWorksheetData> {
  const sheet = data.sheets[id]
  if (sheet === undefined) throw new Error(`fixture has no sheet ${id}`)
  return sheet
}

function cellAt(sheet: Partial<IWorksheetData>, row: number, column: number): ICellData | undefined {
  return sheet.cellData?.[row]?.[column]
}

function entryOf(entries: ReadonlyMap<string, Uint8Array>, name: string): Uint8Array {
  const bytes = entries.get(name)
  if (bytes === undefined) throw new Error(`package has no ${name}`)
  return bytes
}

function expectOnlySpan(
  sourcePart: Uint8Array,
  writtenPart: Uint8Array,
  span: { readonly start: number; readonly end: number },
): void {
  expect(writtenPart.subarray(0, span.start)).toEqual(sourcePart.subarray(0, span.start))
  const tail = sourcePart.subarray(span.end)
  expect(writtenPart.subarray(writtenPart.byteLength - tail.byteLength)).toEqual(tail)
}

function indexOfSequence(haystack: Uint8Array, needle: Uint8Array, from: number): number {
  if (needle.byteLength === 0) return from
  const last = haystack.byteLength - needle.byteLength
  for (let index = from; index <= last; index += 1) {
    let match = true
    for (let offset = 0; offset < needle.byteLength; offset += 1) {
      if (haystack[index + offset] !== needle[offset]) {
        match = false
        break
      }
    }
    if (match) return index
  }
  return -1
}

function expectSourceOutsideSpans(
  sourcePart: Uint8Array,
  writtenPart: Uint8Array,
  spans: readonly { readonly start: number; readonly end: number }[],
): void {
  const sorted = [...spans].sort((left, right) => left.start - right.start || left.end - right.end)
  if (sorted.length === 0) {
    expect(writtenPart).toEqual(sourcePart)
    return
  }
  let sourceCursor = 0
  let writtenCursor = 0
  for (const [index, span] of sorted.entries()) {
    const gap = sourcePart.subarray(sourceCursor, span.start)
    const found = indexOfSequence(writtenPart, gap, writtenCursor)
    expect(found, `gap before span ${String(index)} is absent`).toBeGreaterThanOrEqual(0)
    if (index === 0) expect(found).toBe(0)
    writtenCursor = found + gap.byteLength
    sourceCursor = span.end
  }
  const tail = sourcePart.subarray(sourceCursor)
  const foundTail = indexOfSequence(writtenPart, tail, writtenCursor)
  expect(foundTail).toBeGreaterThanOrEqual(0)
  expect(foundTail + tail.byteLength).toBe(writtenPart.byteLength)
}

function expectDiff(
  source: Uint8Array,
  written: Uint8Array,
  changed: readonly string[],
): ReturnType<typeof comparePackages> {
  const diff = comparePackages(source, written, { changed })
  expect(diff.ok, formatDiffReport(diff)).toBe(true)
  return diff
}

function byteOffsetOf(text: string, index: number): number {
  return strToU8(text.slice(0, index)).byteLength
}

function rowCellReferences(xml: string): readonly string[] {
  return [...xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>|<c\b[^>]*\br="([A-Z]+\d+)"/gu)]
    .map(match => match[1] === undefined ? match[2] ?? '' : `row ${match[1]}`)
}

function cellSpanOf(xml: string, reference: string): { readonly start: number; readonly end: number } {
  const pattern = new RegExp(`<c\\b[^>]*\\br="${reference}"[^>]*(?:/>|>[\\s\\S]*?</c>)`, 'u')
  const match = pattern.exec(xml)
  if (match === null) throw new Error(`fixture has no cell ${reference}`)
  return {
    start: byteOffsetOf(xml, match.index),
    end: byteOffsetOf(xml, match.index + match[0].length),
  }
}

function expectCellSpanDiff(
  source: Uint8Array,
  written: Uint8Array,
  reference: string,
): void {
  const sourceEntries = unzipEntries(source)
  const writtenEntries = unzipEntries(written)
  const sourceXml = strFromU8(entryOf(sourceEntries, SHEET))
  expectOnlySpan(entryOf(sourceEntries, SHEET), entryOf(writtenEntries, SHEET), cellSpanOf(sourceXml, reference))
}

function xmlAttributeValueSpan(
  xml: string,
  from: number,
  name: string,
): { readonly start: number; readonly end: number } {
  const marker = `${name}="`
  const index = xml.indexOf(marker, from)
  if (index < 0) throw new Error(`fixture has no ${name} attribute`)
  const valueStart = index + marker.length
  const valueEnd = xml.indexOf('"', valueStart)
  if (valueEnd < 0) throw new Error(`fixture has an unterminated ${name} attribute`)
  return { start: byteOffsetOf(xml, valueStart), end: byteOffsetOf(xml, valueEnd) }
}

function collectionCountSpan(
  xml: string,
  collection: string,
): { readonly start: number; readonly end: number } {
  const index = xml.indexOf(`<${collection} `)
  if (index < 0) throw new Error(`fixture has no ${collection} collection`)
  return xmlAttributeValueSpan(xml, index, 'count')
}

function closingTagSpan(
  xml: string,
  qname: string,
): { readonly start: number; readonly end: number } {
  const index = xml.indexOf(`</${qname}>`)
  if (index < 0) throw new Error(`fixture has no closing ${qname} tag`)
  const start = byteOffsetOf(xml, index)
  return { start, end: start }
}

function sourceRowSpans(xml: string): readonly { readonly row: number; readonly start: number; readonly end: number }[] {
  const spans: Array<{ row: number; start: number; end: number }> = []
  const opener = /<row\b[^>]*>/gu
  for (;;) {
    const match = opener.exec(xml)
    if (match === null) break
    const reference = /\br="(\d+)"/u.exec(match[0])
    const row = reference === null ? 0 : Number(reference[1])
    const close = xml.indexOf('</row>', match.index + match[0].length)
    if (close < 0) break
    spans.push({
      row,
      start: byteOffsetOf(xml, match.index),
      end: byteOffsetOf(xml, close + '</row>'.length),
    })
  }
  return spans
}

function dimensionRefSpan(xml: string): { readonly start: number; readonly end: number } | undefined {
  const match = /<dimension\b[^>]*\bref="([^"]*)"/u.exec(xml)
  if (match === null) return undefined
  const reference = match[1]
  if (reference === undefined) throw new Error('dimension has no ref attribute value')
  const valueStart = match.index + match[0].length - reference.length - 1
  return {
    start: byteOffsetOf(xml, valueStart),
    end: byteOffsetOf(xml, valueStart + reference.length),
  }
}

function expectStructuralDiff(
  result: OfficeWriteResult,
  source: Uint8Array,
  expectedSpans: readonly { readonly start: number; readonly end: number }[],
): void {
  expectDiff(source, result.bytes, [SHEET])
  expect(result.disposition.spliced.names).toEqual([SHEET])
  const sourceEntries = unzipEntries(source)
  const writtenEntries = unzipEntries(result.bytes)
  expectSourceOutsideSpans(entryOf(sourceEntries, SHEET), entryOf(writtenEntries, SHEET), expectedSpans)
  expect(validateOpc(writtenEntries)).toEqual([])
}

async function refusalOf(result: Promise<unknown>): Promise<OfficeExchangeError> {
  try {
    await result
  } catch (error) {
    if (error instanceof OfficeExchangeError) return error
    throw error
  }
  throw new Error('Expected the writer to refuse')
}

const SOURCE_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${SPREADSHEET}">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="0.000"/>
    <numFmt numFmtId="165" formatCode="0.000"/>
  </numFmts>
  <fonts count="6">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><i/><strike/><u val="double"/><vertAlign val="subscript"/><sz val="14"/><name val="Arial"/><color rgb="FF112233"/></font>
    <font><u val="none"/><vertAlign val="superscript"/></font>
    <font><sz val="0"/><name/></font>
    <font><u val="single"/><color rgb="445566"/></font>
    <font><i/><strike/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF445566"/></patternFill></fill>
    <fill><gradientFill type="linear" degree="90"/></fill>
    <fill><patternFill patternType="solid"/></fill>
    <fill><patternFill patternType="darkGrid"/></fill>
  </fills>
  <borders count="7">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FF778899"/></left><right style="hair"/></border>
    <border diagonalUp="1"/>
    <border><diagonal style="thin"/></border>
    <border><left style="medium"><color rgb="zzz"/></left></border>
    <border/>
    <border><left style="medium"/><right style="thick"><color rgb="FFAABBCC"/></right></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="12">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1" shrinkToFit="1" textRotation="45"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="2" xfId="0" applyFont="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="3" xfId="0"><alignment horizontal="right" vertical="middle"/></xf>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="4" xfId="0"><alignment horizontal="justify" vertical="bottom"/></xf>
    <xf numFmtId="0" fontId="5" fillId="0" borderId="5" xfId="0"><alignment horizontal="centerContinuous" vertical="justify"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment horizontal="distributed" vertical="distributed"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment horizontal="fill" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1" pivotButton="1"><protection locked="0" hidden="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="0" indent="0" readingOrder="0"/></xf>
  </cellXfs>
</styleSheet>`

const STYLE_SHEET_DATA = Array.from({ length: 12 }, (_, index) => {
  const row = index + 1
  const style = index
  return `<row r="${String(row)}"><c r="A${String(row)}" s="${String(style)}"><v>${String(row)}</v></c></row>`
}).join('')

describe('writeWorkbookFromSource', () => {
  it('writes an unchanged snapshot back as the identical source package', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>',
      sharedStrings: ['Alpha', 'Beta'],
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const result = await writeWorkbookFromSource(data, source)
    expect(result.bytes).toEqual(source)
    expect(result.disposition.spliced).toEqual({ count: 0, names: [] })
    expect(result.disposition.verbatimCount).toBe(unzipEntries(source).size)
    const repeated = await writeWorkbookFromSource(data, source)
    expect(repeated.bytes).toEqual(result.bytes)
  })

  it('does not rewrite a stale dimension when nothing changed', async () => {
    const source = xlsxPackage({
      dimension: 'A1:C10',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>'
        + '<row r="2"><c r="A2"><v>3</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const result = await writeWorkbookFromSource(data, source)
    expect(result.bytes).toEqual(source)
  })

  it('splices one numeric cell and keeps every other entry byte-identical', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>',
      sharedStrings: ['Alpha', 'Beta'],
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const anchor = anchorOf(data, 'sheet-1!0:0')
    const sheet = sheetOf(data)
    const cell = cellAt(sheet, 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 10
    const result = await writeWorkbookFromSource(data, source)
    expect(result.disposition.spliced).toEqual({ count: 1, names: [SHEET] })
    const written = unzipEntries(result.bytes)
    expectOnlySpan(entryOf(unzipEntries(source), SHEET), entryOf(written, SHEET), anchor)
    expect(strFromU8(entryOf(written, SHEET))).toContain('<c r="A1"><v>10</v></c>')
    expectDiff(source, result.bytes, [SHEET])
    expect(validateOpc(written)).toEqual([])
    const reread = await readWorkbook(result.bytes)
    const rereadSheet = sheetOf(reread)
    expect(cellAt(rereadSheet, 0, 0)?.v).toBe(10)
    expect(cellAt(rereadSheet, 0, 1)?.v).toBe('Alpha')
    expect(entryOf(written, SHARED_STRINGS)).toEqual(entryOf(unzipEntries(source), SHARED_STRINGS))
  })

  it('uses the known-different replacement path for a changed worksheet part', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>',
      sharedStrings: ['Alpha', 'Beta'],
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    const knownDifferent = vi.spyOn(OfficePackage.prototype, 'replacePartKnownDifferent')
    const equalityChecked = vi.spyOn(OfficePackage.prototype, 'replacePart')
    try {
      const unchanged = await writeWorkbookFromSource(data, source)
      expect(unchanged.bytes).toEqual(source)
      expect(knownDifferent).not.toHaveBeenCalled()
      expect(equalityChecked).not.toHaveBeenCalled()

      cell.v = 10
      const result = await writeWorkbookFromSource(data, source)
      expect(knownDifferent).toHaveBeenCalledTimes(1)
      expect(knownDifferent.mock.calls[0]?.[0]).toBe(SHEET)
      expect(equalityChecked).not.toHaveBeenCalled()
      expect(cellAt(sheetOf(await readWorkbook(result.bytes)), 0, 0)?.v).toBe(10)
    } finally {
      knownDifferent.mockRestore()
      equalityChecked.mockRestore()
    }
  })

  it('splices a formula cell and preserves its cached result representation', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B1',
      sheetData: '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>4</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    const cell = cellAt(sheet, 0, 1)
    if (cell === undefined) throw new Error('fixture has no B1')
    cell.f = '=A1*3'
    cell.v = 6
    cell.t = 2
    const result = await writeWorkbookFromSource(data, source)
    expectDiff(source, result.bytes, [SHEET])
    expectCellSpanDiff(source, result.bytes, 'B1')
    const sheetXml = strFromU8(entryOf(unzipEntries(result.bytes), SHEET))
    expect(sheetXml).toContain('<f>A1*3</f><v>6</v>')
    const reread = await readWorkbook(result.bytes)
    const rereadCell = cellAt(sheetOf(reread), 0, 1)
    expect(rereadCell?.f).toBe('=A1*3')
    expect(rereadCell?.v).toBe(6)
  })

  it('appends an xf for a style-only change and preserves every other style byte', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    const cell = cellAt(sheet, 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { bl: BooleanNumber.TRUE }
    const result = await writeWorkbookFromSource(data, source)
    expectDiff(source, result.bytes, [STYLES, SHEET])
    expect(result.disposition.spliced.names).toEqual([STYLES, SHEET])
    const sourceStyles = strFromU8(entryOf(unzipEntries(source), STYLES))
    const writtenStyles = strFromU8(entryOf(unzipEntries(result.bytes), STYLES))
    expectSourceOutsideSpans(
      entryOf(unzipEntries(source), STYLES),
      entryOf(unzipEntries(result.bytes), STYLES),
      [
        collectionCountSpan(sourceStyles, 'fonts'),
        closingTagSpan(sourceStyles, 'fonts'),
        collectionCountSpan(sourceStyles, 'cellXfs'),
        closingTagSpan(sourceStyles, 'cellXfs'),
      ],
    )
    expect(writtenStyles).toContain('<fonts count="3"')
    expect(writtenStyles).toContain('<font><b/></font>')
    expect(writtenStyles).toContain('<cellXfs count="4"')
    expect(writtenStyles).toContain('fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"')
    expectCellSpanDiff(source, result.bytes, 'A1')
    expect(strFromU8(entryOf(unzipEntries(result.bytes), SHEET))).toContain('<c r="A1" s="3"><v>1</v></c>')
    expect(validateOpc(unzipEntries(result.bytes))).toEqual([])
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.s).toEqual({ bl: BooleanNumber.TRUE })
  })

  it('reuses an existing xf for a style change without touching the style table', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { bl: BooleanNumber.TRUE, fs: 11, ff: 'Calibri' }
    const result = await writeWorkbookFromSource(data, source)
    expectDiff(source, result.bytes, [SHEET])
    expect(result.disposition.spliced.names).toEqual([SHEET])
    expect(entryOf(unzipEntries(result.bytes), STYLES)).toEqual(entryOf(unzipEntries(source), STYLES))
    expectCellSpanDiff(source, result.bytes, 'A1')
    expect(strFromU8(entryOf(unzipEntries(result.bytes), SHEET))).toContain('<c r="A1" s="2"><v>1</v></c>')
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.s).toEqual({ bl: BooleanNumber.TRUE, fs: 11, ff: 'Calibri' })
  })

  it('appends a custom number format for a style-only change', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { n: { pattern: '0.000' } }
    const result = await writeWorkbookFromSource(data, source)
    expectDiff(source, result.bytes, [STYLES, SHEET])
    const sourceStyles = strFromU8(entryOf(unzipEntries(source), STYLES))
    const writtenStyles = strFromU8(entryOf(unzipEntries(result.bytes), STYLES))
    expectSourceOutsideSpans(
      entryOf(unzipEntries(source), STYLES),
      entryOf(unzipEntries(result.bytes), STYLES),
      [
        collectionCountSpan(sourceStyles, 'numFmts'),
        closingTagSpan(sourceStyles, 'numFmts'),
        collectionCountSpan(sourceStyles, 'cellXfs'),
        closingTagSpan(sourceStyles, 'cellXfs'),
      ],
    )
    expect(writtenStyles).toContain('<numFmts count="2"')
    expect(writtenStyles).toContain('<numFmt numFmtId="165" formatCode="0.000"/>')
    expect(writtenStyles).toContain('<cellXfs count="4"')
    expect(writtenStyles).toContain('numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"')
    expectCellSpanDiff(source, result.bytes, 'A1')
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.s).toMatchObject({ n: { pattern: '0.000' } })
  })

  it('refuses a model style component the writer cannot express', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { bg: { th: 5 } }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('theme color reference')
  })

  it('writes a changed shared string as an inline string and leaves sharedStrings untouched', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B1',
      sheetData: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>2</v></c></row>',
      sharedStrings: ['Alpha', 'Beta'],
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 'Gamma'
    const result = await writeWorkbookFromSource(data, source)
    expectDiff(source, result.bytes, [SHEET])
    expectCellSpanDiff(source, result.bytes, 'A1')
    const written = unzipEntries(result.bytes)
    const sheetXml = strFromU8(entryOf(written, SHEET))
    expect(sheetXml).toContain('<c r="A1" t="inlineStr"><is><t xml:space="preserve">Gamma</t></is></c>')
    expect(sheetXml).not.toContain('<c r="A1" t="s">')
    expect(entryOf(written, SHARED_STRINGS)).toEqual(entryOf(unzipEntries(source), SHARED_STRINGS))
    expect(validateOpc(written)).toEqual([])
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe('Gamma')
  })

  it('splices the master cell of a merged range without touching the merge list', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B2',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>Merged</t></is></c></row>'
        + '<row r="2"><c r="A2"><v>1</v></c></row>',
      extra: '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 'Updated'
    const result = await writeWorkbookFromSource(data, source)
    expectDiff(source, result.bytes, [SHEET])
    expectCellSpanDiff(source, result.bytes, 'A1')
    const sheetXml = strFromU8(entryOf(unzipEntries(result.bytes), SHEET))
    expect(sheetXml).toContain('<mergeCell ref="A1:B1"/>')
    expect(sheetXml).toContain('Updated')
    const reread = await readWorkbook(result.bytes)
    expect(sheetOf(reread).mergeData).toEqual([{ startRow: 0, startColumn: 0, endRow: 1, endColumn: 2 }])
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe('Updated')
  })

  it('splices a date cell and preserves its date number format index', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" s="1"><v>45000</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    expect(cell.v).toBe(45000)
    cell.v = 45001
    const result = await writeWorkbookFromSource(data, source)
    expectDiff(source, result.bytes, [SHEET])
    expectCellSpanDiff(source, result.bytes, 'A1')
    const sheetXml = strFromU8(entryOf(unzipEntries(result.bytes), SHEET))
    expect(sheetXml).toContain('<c r="A1" s="1"><v>45001</v></c>')
    const reread = await readWorkbook(result.bytes)
    const rereadCell = cellAt(sheetOf(reread), 0, 0)
    expect(rereadCell?.v).toBe(45001)
    expect(rereadCell?.s).toEqual(cell.s)
  })

  it('keeps the source s style index when a styled cell value changes (SC-012)', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" s="2" t="inlineStr"><is><t>Head</t></is></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 'Head2'
    const result = await writeWorkbookFromSource(data, source)
    expectDiff(source, result.bytes, [SHEET])
    expectCellSpanDiff(source, result.bytes, 'A1')
    const written = unzipEntries(result.bytes)
    const sheetXml = strFromU8(entryOf(written, SHEET))
    expect(sheetXml).toContain('<c r="A1" s="2" t="inlineStr">')
    expect(sheetXml).toContain('Head2')
    expect(entryOf(written, STYLES)).toEqual(entryOf(unzipEntries(source), STYLES))
    const reread = await readWorkbook(result.bytes)
    const rereadCell = cellAt(sheetOf(reread), 0, 0)
    expect(rereadCell?.v).toBe('Head2')
    expect(rereadCell?.s).toEqual(cell.s)
  })

  it('inserts a worksheet row and shifts the rows below it', async () => {
    const source = xlsxPackage({
      dimension: 'A1:A3',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>one</t></is></c></row>'
        + '<row r="2"><c r="A2" t="inlineStr"><is><t>two</t></is></c></row>'
        + '<row r="3"><c r="A3" t="inlineStr"><is><t>three</t></is></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    sheet.cellData = {
      0: sheet.cellData?.[0] ?? {},
      1: { 0: { v: 'inserted', t: 1 } },
      2: sheet.cellData?.[1] ?? {},
      3: sheet.cellData?.[2] ?? {},
    }
    const result = await writeWorkbookFromSource(data, source)
    const sourceXml = strFromU8(entryOf(unzipEntries(source), SHEET))
    const dimension = dimensionRefSpan(sourceXml)
    if (dimension === undefined) throw new Error('fixture has no dimension ref')
    expectStructuralDiff(result, source, [
      ...sourceRowSpans(sourceXml).filter(span => span.row >= 2),
      dimension,
    ])
    const reread = await readWorkbook(result.bytes)
    const rereadSheet = sheetOf(reread)
    expect(cellAt(rereadSheet, 0, 0)?.v).toBe('one')
    expect(cellAt(rereadSheet, 1, 0)?.v).toBe('inserted')
    expect(cellAt(rereadSheet, 2, 0)?.v).toBe('two')
    expect(cellAt(rereadSheet, 3, 0)?.v).toBe('three')
    expect(strFromU8(entryOf(unzipEntries(result.bytes), SHEET))).toContain('<dimension ref="A1:A4"/>')
  })

  it('rewrites every shifted cell reference after a row insert', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B3',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>one-a</t></is></c>'
        + '<c r="B1" t="inlineStr"><is><t>one-b</t></is></c></row>'
        + '<row r="2"><c r="A2" t="inlineStr"><is><t>two-a</t></is></c>'
        + '<c r="B2" t="inlineStr"><is><t>two-b</t></is></c></row>'
        + '<row r="3"><c r="A3" t="inlineStr"><is><t>three-a</t></is></c>'
        + '<c r="B3" t="inlineStr"><is><t>three-b</t></is></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    const original = sheet.cellData
    sheet.cellData = {
      0: original?.[0] ?? {},
      1: { 0: { v: 'inserted-a', t: 1 }, 1: { v: 'inserted-b', t: 1 } },
      2: original?.[1] ?? {},
      3: original?.[2] ?? {},
    }
    const result = await writeWorkbookFromSource(data, source)
    const sheetXml = strFromU8(entryOf(unzipEntries(result.bytes), SHEET))
    expect(rowCellReferences(sheetXml)).toEqual([
      'row 1', 'A1', 'B1',
      'row 2', 'A2', 'B2',
      'row 3', 'A3', 'B3',
      'row 4', 'A4', 'B4',
    ])

    const rereadSheet = sheetOf(await readWorkbook(result.bytes))
    const expected = [
      ['one-a', 'one-b'],
      ['inserted-a', 'inserted-b'],
      ['two-a', 'two-b'],
      ['three-a', 'three-b'],
    ]
    for (const [row, values] of expected.entries()) {
      for (const [column, value] of values.entries()) {
        expect(cellAt(rereadSheet, row, column)?.v).toBe(value)
      }
    }
  })

  it('deletes a worksheet row and shifts the rows below it', async () => {
    const source = xlsxPackage({
      dimension: 'A1:A3',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>one</t></is></c></row>'
        + '<row r="2"><c r="A2" t="inlineStr"><is><t>two</t></is></c></row>'
        + '<row r="3"><c r="A3" t="inlineStr"><is><t>three</t></is></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    sheet.cellData = {
      0: sheet.cellData?.[1] ?? {},
      1: sheet.cellData?.[2] ?? {},
    }
    const result = await writeWorkbookFromSource(data, source)
    const sourceXml = strFromU8(entryOf(unzipEntries(source), SHEET))
    const dimension = dimensionRefSpan(sourceXml)
    if (dimension === undefined) throw new Error('fixture has no dimension ref')
    expectStructuralDiff(result, source, [...sourceRowSpans(sourceXml), dimension])
    const reread = await readWorkbook(result.bytes)
    const rereadSheet = sheetOf(reread)
    expect(cellAt(rereadSheet, 0, 0)?.v).toBe('two')
    expect(cellAt(rereadSheet, 1, 0)?.v).toBe('three')
    expect(cellAt(rereadSheet, 2, 0)).toBeUndefined()
    expect(strFromU8(entryOf(unzipEntries(result.bytes), SHEET))).toContain('<dimension ref="A1:A2"/>')
  })

  it('inserts a worksheet column and rewrites the shifted cell references', async () => {
    const source = xlsxPackage({
      dimension: 'A1:C1',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c>'
        + '<c r="B1" t="inlineStr"><is><t>b</t></is></c>'
        + '<c r="C1" t="inlineStr"><is><t>c</t></is></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    const row = sheet.cellData?.[0]
    if (row === undefined) throw new Error('fixture has no row 1')
    sheet.cellData = {
      0: {
        0: row[0] ?? {},
        1: { v: 'x', t: 1 },
        2: row[1] ?? {},
        3: row[2] ?? {},
      },
    }
    const result = await writeWorkbookFromSource(data, source)
    const sourceXml = strFromU8(entryOf(unzipEntries(source), SHEET))
    const dimension = dimensionRefSpan(sourceXml)
    if (dimension === undefined) throw new Error('fixture has no dimension ref')
    expectStructuralDiff(result, source, [...sourceRowSpans(sourceXml), dimension])
    const sheetXml = strFromU8(entryOf(unzipEntries(result.bytes), SHEET))
    expect(sheetXml).toContain('<c r="B1" t="inlineStr"><is><t xml:space="preserve">x</t></is></c>')
    expect(sheetXml).toContain('<c r="C1" t="inlineStr"><is><t>b</t></is></c>')
    expect(sheetXml).toContain('<c r="D1" t="inlineStr"><is><t>c</t></is></c>')
    expect(sheetXml).toContain('<dimension ref="A1:D1"/>')
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe('a')
    expect(cellAt(sheetOf(reread), 0, 1)?.v).toBe('x')
    expect(cellAt(sheetOf(reread), 0, 2)?.v).toBe('b')
    expect(cellAt(sheetOf(reread), 0, 3)?.v).toBe('c')
  })

  it('deletes a worksheet column and rewrites the shifted cell references', async () => {
    const source = xlsxPackage({
      dimension: 'A1:C1',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c>'
        + '<c r="B1" t="inlineStr"><is><t>b</t></is></c>'
        + '<c r="C1" t="inlineStr"><is><t>c</t></is></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    const row = sheet.cellData?.[0]
    if (row === undefined) throw new Error('fixture has no row 1')
    sheet.cellData = {
      0: {
        0: row[0] ?? {},
        1: row[2] ?? {},
      },
    }
    const result = await writeWorkbookFromSource(data, source)
    const sourceXml = strFromU8(entryOf(unzipEntries(source), SHEET))
    const dimension = dimensionRefSpan(sourceXml)
    if (dimension === undefined) throw new Error('fixture has no dimension ref')
    expectStructuralDiff(result, source, [...sourceRowSpans(sourceXml), dimension])
    const sheetXml = strFromU8(entryOf(unzipEntries(result.bytes), SHEET))
    expect(sheetXml).toContain('<c r="A1" t="inlineStr"><is><t>a</t></is></c>')
    expect(sheetXml).toContain('<c r="B1" t="inlineStr"><is><t>c</t></is></c>')
    expect(sheetXml).not.toContain('>b<')
    expect(sheetXml).toContain('<dimension ref="A1:B1"/>')
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe('a')
    expect(cellAt(sheetOf(reread), 0, 1)?.v).toBe('c')
    expect(cellAt(sheetOf(reread), 0, 2)).toBeUndefined()
  })

  it('re-derives anchors in sheet-summary mode before splicing', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>'
        + '<row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>',
    })
    const config = resolveConfig({ maxSourceAnchors: 1 })
    const data = await readWorkbookWithSource(source, 'xlsx', config)
    const state = sourceStateOf(data)
    expect(state?.anchorMode).toBe('sheet-summary')
    const cell = cellAt(sheetOf(data), 1, 1)
    if (cell === undefined) throw new Error('fixture has no B2')
    cell.v = 40
    const result = await writeWorkbookFromSource(data, source, config)
    expectDiff(source, result.bytes, [SHEET])
    expectCellSpanDiff(source, result.bytes, 'B2')
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe(1)
    expect(cellAt(sheetOf(reread), 1, 1)?.v).toBe(40)
  })

  it('writes a changed boolean in its existing t="b" representation', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" t="b"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = false
    const result = await writeWorkbookFromSource(data, source)
    expectDiff(source, result.bytes, [SHEET])
    expectCellSpanDiff(source, result.bytes, 'A1')
    expect(strFromU8(entryOf(unzipEntries(result.bytes), SHEET))).toContain('<c r="A1" t="b"><v>0</v></c>')
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe(false)
  })

  it('changes a string cell to a number while keeping the t attribute present', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>7</t></is></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 8
    cell.t = 2
    const result = await writeWorkbookFromSource(data, source)
    expectDiff(source, result.bytes, [SHEET])
    expectCellSpanDiff(source, result.bytes, 'A1')
    expect(strFromU8(entryOf(unzipEntries(result.bytes), SHEET))).toContain('<c r="A1" t="n"><v>8</v></c>')
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe(8)
  })

  it('writes a changed error in its existing t="e" representation', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = '#N/A'
    const result = await writeWorkbookFromSource(data, source)
    expectDiff(source, result.bytes, [SHEET])
    expectCellSpanDiff(source, result.bytes, 'A1')
    expect(strFromU8(entryOf(unzipEntries(result.bytes), SHEET))).toContain('<c r="A1" t="e"><v>#N/A</v></c>')
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe('#N/A')
  })

  it('refuses a rich-text shared string cell whose text changes', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" t="s"><v>0</v></c></row>',
      richSharedStrings: ['<si><r><rPr><b/></rPr><t>Rich</t></r></si>'],
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 'Plain'
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('rich-text')
  })

  it('refuses a shared formula cell whose formula changes', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 1)
    if (cell === undefined) throw new Error('fixture has no B1')
    cell.f = '=A1*3'
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('shared')
  })

  it('refuses a snapshot without source anchors', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    delete (data as SidecarCarrier).officeExchange
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('no source anchors')
  })

  it('refuses a source larger than the file bound', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const error = await refusalOf(writeWorkbookFromSource(data, source, resolveConfig({ maxFileBytes: 1 })))
    expect(error.code).toBe('too-large')
  })

  it('refuses a sidecar naming another document format', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    withSourceState(data, state => ({ ...state, format: 'docx' }))
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-format')
  })

  it('re-anchors a source whose bytes changed after the snapshot was read', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    withSourceState(data, state => ({ ...state, packageDigest: 'a'.repeat(64) }))
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 2
    const result = await writeWorkbookFromSource(data, source)
    expect(strFromU8(entryOf(unzipEntries(result.bytes), SHEET))).toContain('<v>2</v>')
  })

  it('refuses a source that changed and cannot be re-anchored', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      extra: '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        + '<mc:Choice Requires="x14"><extLst xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"/></mc:Choice>'
        + '</mc:AlternateContent>',
    })
    const clean = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(clean, 'xlsx')
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('re-anchored')
  })

  it('refuses a malformed officeExchange carrier', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    ;(data as { officeExchange?: unknown }).officeExchange = 'malformed'
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('officeExchange state is malformed')
  })

  it('refuses an unreadable source package', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const error = await refusalOf(writeWorkbookFromSource(data, new Uint8Array([1, 2, 3])))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('Cannot open the source package')
  })

  it('refuses an unreadable worksheet part', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const corrupted = corruptEntryCrc(source, SHEET)
    const archive = open(corrupted, { maxEntries: 100 })
    withSourceState(data, state => ({
      ...state,
      packageDigest: archive.source.digest,
      byteLength: archive.source.bytes.byteLength,
    }))
    const error = await refusalOf(writeWorkbookFromSource(data, corrupted))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('Cannot read worksheet part')
  })

  it('refuses an anchor from another unit kind', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    withSourceState(data, (state) => {
      if (state.anchorMode !== 'per-unit') throw new Error('fixture has no per-unit anchors')
      return {
        ...state,
        anchors: state.anchors.map(anchor => ({ ...anchor, kind: 'paragraph' as const })),
      }
    })
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('names a paragraph')
  })

  it('refuses a sidecar that omits the resolved worksheet part', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    withSourceState(data, state => ({ ...state, sourceParts: ['xl/styles.xml'] }))
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('does not name worksheet part')
  })

  it('refuses a summary state missing the sheet summary', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx', resolveConfig({ maxSourceAnchors: 1 }))
    withSourceState(data, (state) => {
      if (state.anchorMode !== 'sheet-summary') throw new Error('fixture has no summaries')
      return { ...state, sheets: state.sheets.map(summary => ({ ...summary, sheetId: 'sheet-9' })) }
    })
    const error = await refusalOf(writeWorkbookFromSource(data, source, resolveConfig({ maxSourceAnchors: 1 })))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('no summary for sheet')
  })

  it('refuses a summary naming another part', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx', resolveConfig({ maxSourceAnchors: 1 }))
    withSourceState(data, (state) => {
      if (state.anchorMode !== 'sheet-summary') throw new Error('fixture has no summaries')
      return {
        ...state,
        sourceParts: [...state.sourceParts, STYLES],
        sheets: state.sheets.map(summary => ({ ...summary, part: STYLES })),
      }
    })
    const error = await refusalOf(writeWorkbookFromSource(data, source, resolveConfig({ maxSourceAnchors: 1 })))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('summary names')
  })

  it('refuses a summary whose digest does not match the source', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx', resolveConfig({ maxSourceAnchors: 1 }))
    withSourceState(data, (state) => {
      if (state.anchorMode !== 'sheet-summary') throw new Error('fixture has no summaries')
      return { ...state, sheets: state.sheets.map(summary => ({ ...summary, digest: '0'.repeat(64) })) }
    })
    const error = await refusalOf(writeWorkbookFromSource(data, source, resolveConfig({ maxSourceAnchors: 1 })))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('does not match its recorded summary')
  })

  it('refuses a model with more cells than the bound', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    sheet.cellData = { 0: { 0: { v: 1 }, 1: { v: 2 } } }
    const error = await refusalOf(writeWorkbookFromSource(data, source, resolveConfig({ maxCells: 1 })))
    expect(error.code).toBe('too-large')
  })

  it('aligns rows when cells move without changing the cell count', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    sheetOf(data).cellData = { 0: { 0: { v: 1 }, 2: { v: 3 } } }
    const result = await writeWorkbookFromSource(data, source)
    expect(result.disposition.spliced.names).toEqual([SHEET])
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 2)?.v).toBe(3)
  })

  it('ignores model cells that carry no value, formula, style, or document', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    sheetOf(data).cellData = { 0: { 0: { v: 1 }, 1: {} } }
    const result = await writeWorkbookFromSource(data, source)
    expect(result.bytes).toEqual(source)
  })

  it('maps every source style component while appending one number format', async () => {
    const source = xlsxPackage({
      dimension: 'A1:A12',
      sheetData: STYLE_SHEET_DATA,
      styles: SOURCE_STYLES,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    for (let row = 0; row < 12; row += 1) {
      const cell = cellAt(sheet, row, 0)
      if (cell === undefined) throw new Error(`fixture has no A${String(row + 1)}`)
      cell.s = { ...styleOf(cell), n: { pattern: '0.0000' } }
    }
    const result = await writeWorkbookFromSource(data, source)
    expect(result.disposition.spliced.names).toEqual([STYLES, SHEET])
    const styles = strFromU8(entryOf(unzipEntries(result.bytes), STYLES))
    expect(styles).toContain('<numFmt numFmtId="166" formatCode="0.0000"/>')
    expect(styles).toContain('quotePrefix="1"')
    expect(styles).toContain('pivotButton="1"')
    expect(styles).toContain('<protection locked="0" hidden="1"/>')
    expect(styles).toContain('<alignment horizontal="left" vertical="top" wrapText="1" shrinkToFit="1" textRotation="45"/>')
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 2, 0)?.s).toMatchObject({
      bl: BooleanNumber.TRUE,
      it: BooleanNumber.TRUE,
      st: { s: BooleanNumber.TRUE },
      va: 2,
      fs: 14,
      ff: 'Arial',
    })
  })

  it('writes every modeled style component into one appended xf', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = {
      bl: BooleanNumber.TRUE,
      it: BooleanNumber.TRUE,
      st: { s: BooleanNumber.TRUE },
      ul: { s: BooleanNumber.TRUE, t: TextDecoration.DOUBLE },
      va: 2,
      fs: 14,
      ff: 'Arial',
      cl: { rgb: '#112233' },
      bg: { rgb: '#445566' },
      bd: {
        t: { s: 1, cl: { rgb: '#778899' } },
        r: { s: 2, cl: { rgb: '#000000' } },
        b: { s: 0, cl: { rgb: '#000000' } },
        l: { s: 13, cl: { rgb: '#000000' } },
      },
      n: { pattern: '0.0000' },
      ht: 1,
      vt: 3,
      tb: 3,
      stf: BooleanNumber.TRUE,
      tr: { a: -30 },
    }
    const result = await writeWorkbookFromSource(data, source)
    expect(result.disposition.spliced.names).toEqual([STYLES, SHEET])
    const styles = strFromU8(entryOf(unzipEntries(result.bytes), STYLES))
    expect(styles).toContain('<b/>')
    expect(styles).toContain('<i/>')
    expect(styles).toContain('<strike/>')
    expect(styles).toContain('<u val="double"/>')
    expect(styles).toContain('<vertAlign val="subscript"/>')
    expect(styles).toContain('<color rgb="FF112233"/>')
    expect(styles).toContain('<sz val="14"/>')
    expect(styles).toContain('<name val="Arial"/>')
    expect(styles).toContain('patternType="solid"')
    expect(styles).toContain('<fgColor rgb="FF445566"/>')
    expect(styles).toContain('style="thin"')
    expect(styles).toContain('style="hair"')
    expect(styles).toContain('style="thick"')
    expect(styles).toContain('<numFmt numFmtId="165" formatCode="0.0000"/>')
    expect(styles).toContain('horizontal="left"')
    expect(styles).toContain('vertical="bottom"')
    expect(styles).toContain('wrapText="1"')
    expect(styles).toContain('shrinkToFit="1"')
    expect(styles).toContain('textRotation="120"')
    expect(styles).toContain('applyFont="1"')
    expect(styles).toContain('applyFill="1"')
    expect(styles).toContain('applyBorder="1"')
    expect(styles).toContain('applyAlignment="1"')
    expect(styles).toContain('applyNumberFormat="1"')
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.s).toMatchObject({
      bl: BooleanNumber.TRUE,
      bg: { rgb: '#445566' },
      n: { pattern: '0.0000' },
      ht: 1,
      vt: 3,
    })
  })

  it('defaults a colorless border edge to black when appending a style', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { bd: { t: { s: 1, cl: {} }, b: { s: 2, cl: { rgb: '#112233' } } } }
    const result = await writeWorkbookFromSource(data, source)
    expect(result.disposition.spliced.names).toEqual([STYLES, SHEET])
    const styles = strFromU8(entryOf(unzipEntries(result.bytes), STYLES))
    expect(styles).toContain('<top style="thin"><color rgb="FF000000"/></top>')
    expect(styles).toContain('<bottom style="hair"><color rgb="FF112233"/></bottom>')
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.s).toMatchObject({
      bd: {
        t: { s: 1, cl: { rgb: '#000000' } },
        b: { s: 2, cl: { rgb: '#112233' } },
      },
    })
  })

  const unmodeledNumberFormat = { pattern: '0' }
  Reflect.set(unmodeledNumberFormat, 'pattern', 5)
  const styleRefusals: readonly (readonly [string, ICellData['s'], string])[] = [
    ['an East Asian font family', { eastAsiaFontFamily: 'Arial' }, 'East Asian font family'],
    ['a bottom border line', { bbl: { s: BooleanNumber.TRUE } }, 'bottom border line'],
    ['an overline', { ol: { s: BooleanNumber.TRUE } }, 'overline'],
    ['a text direction', { td: 2 }, 'text direction'],
    ['cell padding', { pd: { t: 1, b: 1, l: 1, r: 1 } }, 'cell padding'],
    ['an invalid color', { bg: { rgb: '#fff' } }, 'color value'],
    ['a theme color', { cl: { th: 5 } }, 'theme color reference'],
    ['an unmodeled number format', { n: unmodeledNumberFormat }, 'number format'],
    ['an underline color', { ul: { s: BooleanNumber.TRUE, c: BooleanNumber.FALSE, cl: { rgb: '#112233' } } }, 'underline color'],
    ['an underline style', { ul: { s: BooleanNumber.TRUE, t: TextDecoration.DASH_LONG_HEAVY } }, 'underline style'],
    ['a strikeout color', { st: { s: BooleanNumber.TRUE, c: BooleanNumber.FALSE, cl: { rgb: '#112233' } } }, 'strikeout color'],
    ['an invalid baseline', { va: 4 as BaselineOffset }, 'baseline'],
    ['a zero font size', { fs: 0 }, 'font size'],
    ['a non-finite font size', { fs: Number.NaN }, 'font size'],
    ['a diagonal border', { bd: { tl_br: { s: 1, cl: { rgb: '#000000' } } } }, 'diagonal border'],
    ['an unknown border style', { bd: { t: { s: 99 as BorderStyleTypes, cl: { rgb: '#000000' } } } }, 'border style'],
    ['an unknown horizontal alignment', { ht: 9 as HorizontalAlign }, 'horizontal alignment'],
    ['an unknown vertical alignment', { vt: 9 as VerticalAlign }, 'vertical alignment'],
    ['an obtuse text rotation', { tr: { a: 100 } }, 'text rotation'],
    ['a non-finite text rotation', { tr: { a: Number.NaN } }, 'text rotation'],
  ]

  it.each(styleRefusals)('refuses %s in a model style', async (_label, style, message) => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = style
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain(message)
  })

  it('accepts unmodeled style values that are not set', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = {
      eastAsiaFontFamily: '',
      td: 0,
      bbl: { s: BooleanNumber.FALSE },
      ol: { s: BooleanNumber.FALSE },
      pd: {},
    }
    const result = await writeWorkbookFromSource(data, source)
    expect(result.bytes).toEqual(source)
  })

  const sourceXfRefusals: readonly (readonly [string, string, string])[] = [
    ['alignment indentation', '<alignment indent="1"/>', 'alignment indentation'],
    ['alignment reading order', '<alignment readingOrder="1"/>', 'alignment reading order'],
    ['an extra alignment attribute', '<alignment foo="bar"/>', 'alignment attribute foo'],
    ['a cell style extension', '<extLst/>', 'cell style extension'],
    ['an unmodeled cell style attribute', '', 'cell style attribute'],
  ]

  it.each(sourceXfRefusals)('refuses a source xf carrying %s when its style changes', async (_label, xfContent, message) => {
    const attribute = xfContent === '' ? ' arbitrary="1"' : ''
    const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<styleSheet xmlns="${SPREADSHEET}">`
      + '<numFmts count="0"/>'
      + '<fonts count="1"><font/></fonts>'
      + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
      + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
      + '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
      + `<cellXfs count="2"><xf/><xf${attribute}>${xfContent}</xf></cellXfs>`
      + '</styleSheet>'
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" s="1"><v>1</v></c></row>',
      styles,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { ...styleOf(cell), ht: 1 as HorizontalAlign, bl: BooleanNumber.TRUE }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain(message)
  })

  it('appends default style collections when the source omits them', async () => {
    const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<styleSheet xmlns="${SPREADSHEET}"><cellXfs count="1"><xf/></cellXfs></styleSheet>`
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      styles,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = {
      bl: BooleanNumber.TRUE,
      bg: { rgb: '#112233' },
      bd: { t: { s: 1, cl: { rgb: '#000000' } } },
      n: { pattern: '0.0000' },
    }
    const result = await writeWorkbookFromSource(data, source)
    const written = strFromU8(entryOf(unzipEntries(result.bytes), STYLES))
    expect(written).toContain('<fonts count="2"')
    expect(written).toContain('<fills count="2"')
    expect(written).toContain('<borders count="2"')
    expect(written).toContain('<numFmts count="1"')
    expect(written).toContain('<cellXfs count="2"')
  })

  it('replaces self-closing style collections and leaves uncounted ones uncounted', async () => {
    const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<styleSheet xmlns="${SPREADSHEET}">`
      + '<numFmts count="0"/><fonts count="0"/><fills count="0"/><borders count="0"/>'
      + '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
      + '<cellXfs><xf/></cellXfs>'
      + '</styleSheet>'
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      styles,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { bl: BooleanNumber.TRUE, bg: { rgb: '#112233' }, n: { pattern: '0.0000' } }
    const result = await writeWorkbookFromSource(data, source)
    const written = strFromU8(entryOf(unzipEntries(result.bytes), STYLES))
    expect(written).toContain('<fonts count="2"><font/><font><b/></font></fonts>')
    expect(written).toContain('<fills count="2">')
    expect(written).toContain('<numFmts count="1">')
    expect(written).toContain('<cellXfs><xf/><xf')
  })

  it('refuses an empty style table root', async () => {
    const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<styleSheet xmlns="${SPREADSHEET}"/>`
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      styles,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    const cell = cellAt(sheet, 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    sheet.cellData = { 0: { 0: cell }, 1: { 0: { v: 'styled', s: { bl: BooleanNumber.TRUE } } } }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('self-closing')
  })

  const STYLES_HEAD = '<?xml version="1.0" encoding="UTF-8"?>'
    + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'

  const malformedStyles: readonly (readonly [string, string, string])[] = [
    ['a non-style root', '<notStyleSheet/>', 'is not a style sheet'],
    ['a root in another namespace', '<styleSheet xmlns="urn:wrong"/>', 'is not in the SpreadsheetML namespace'],
    ['an unterminated comment', 'STYLES_HEAD<!-- unterminated', 'unterminated XML comment'],
    ['an unterminated CDATA section', 'STYLES_HEAD<![CDATA[unterminated', 'unterminated CDATA section'],
    ['an unterminated processing instruction', 'STYLES_HEAD<?pi unterminated', 'unterminated processing instruction'],
    ['an unterminated markup declaration', 'STYLES_HEAD<!DOCTYPE', 'unterminated markup declaration'],
    ['an unterminated element tag', 'STYLES_HEAD<fonts', 'unterminated element tag'],
    ['a malformed end tag', 'STYLES_HEAD</>', 'malformed end tag'],
    ['a malformed start tag without a value', 'STYLES_HEAD<fonts a=></fonts></styleSheet>', 'malformed start tag'],
    ['a tag with an unterminated quote', 'STYLES_HEAD<fonts a="x></fonts></styleSheet>', 'unterminated element tag'],
    ['mismatched end tags', 'STYLES_HEAD<fonts><font></fonts></styleSheet>', 'does not match'],
    ['an unmatched end tag', 'STYLES_HEAD</styleSheet></styleSheet>', 'no matching start tag'],
  ]

  it.each(malformedStyles)('refuses a styles part with %s', async (_label, xml, message) => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      includeStyles: false,
      stylesUnreferenced: xml.replace('STYLES_HEAD', STYLES_HEAD),
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { bl: BooleanNumber.TRUE }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('Cannot parse xl/styles.xml')
    expect(error.message).toContain(message)
  })

  it.each([
    ['a prefixed root', '<x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<x:fonts count="0"/><x:cellXfs count="1"><x:xf/></x:cellXfs></x:styleSheet>', '<x:font>'],
    ['a prefixed root without namespace declarations', '<x:styleSheet>'
      + '<x:fonts count="0"/><x:cellXfs count="1"><x:xf/></x:cellXfs></x:styleSheet>', '<x:font>'],
    ['unknown nested elements', '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<foo><bar/></foo><fonts count="0"/><fills count="0"/><borders count="0"/>'
      + '<cellXfs count="1"><xf/></cellXfs></styleSheet>', '<font><b/></font>'],
    ['nested entry elements', '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<fonts count="1"><font><b><i/></b></font></fonts><fills count="0"/><borders count="0"/>'
      + '<cellXfs count="1"><xf/></cellXfs></styleSheet>', '<b><i/></b>'],
  ])('accepts %s in an unreferenced styles part', async (_label, xml, expected) => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      includeStyles: false,
      stylesUnreferenced: xml,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { bl: BooleanNumber.TRUE }
    const result = await writeWorkbookFromSource(data, source)
    expect(strFromU8(entryOf(unzipEntries(result.bytes), STYLES))).toContain(expected)
  })

  it('refuses a style change when the source has no styles part', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      includeStyles: false,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { bl: BooleanNumber.TRUE }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('does not carry')
  })

  it('refuses a style change when cellXfs is missing or empty', async () => {
    for (const collections of ['', '<cellXfs count="0"/>']) {
      const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + `<styleSheet xmlns="${SPREADSHEET}">`
        + '<fonts count="1"><font/></fonts>'
        + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        + '<borders count="1"><border/></borders>'
        + collections
        + '</styleSheet>'
      const source = xlsxPackage({
        dimension: 'A1',
        sheetData: '<row r="1"><c r="A1" s="0"><v>1</v></c></row>',
        styles,
      })
      const data = await readWorkbookWithSource(source, 'xlsx')
      const cell = cellAt(sheetOf(data), 0, 0)
      if (cell === undefined) throw new Error('fixture has no A1')
      cell.s = { bl: BooleanNumber.TRUE }
      const error = await refusalOf(writeWorkbookFromSource(data, source))
      expect(error.code).toBe('conversion-failed')
      expect(error.message).toContain('cellXfs')
    }
  })

  it('refuses an out-of-range source style index', async () => {
    for (const index of ['9', 'abc']) {
      const source = xlsxPackage({
        dimension: 'A1',
        sheetData: `<row r="1"><c r="A1" s="${index}"><v>1</v></c></row>`,
      })
      const data = await readWorkbookWithSource(source, 'xlsx')
      const cell = cellAt(sheetOf(data), 0, 0)
      if (cell === undefined) throw new Error('fixture has no A1')
      cell.s = { bl: BooleanNumber.TRUE }
      const error = await refusalOf(writeWorkbookFromSource(data, source))
      expect(error.code).toBe('conversion-failed')
      expect(error.message).toContain('style')
    }
  })

  it('resolves a named snapshot style and refuses a missing one', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    data.styles = { named: { bl: BooleanNumber.TRUE } }
    cell.s = 'named'
    const result = await writeWorkbookFromSource(data, source)
    expect(strFromU8(entryOf(unzipEntries(result.bytes), STYLES))).toContain('<b/>')
    cell.s = 'missing'
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('missing workbook style')
  })

  it('uses model sheet metadata when a snapshot omits cell data', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    delete sheet.cellData
    const result = await writeWorkbookFromSource(data, source)
    expect(result.disposition.spliced.names).toEqual([SHEET])
  })

  it('treats style-only and document-only cells as non-empty model cells', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    sheetOf(data).cellData = {
      0: { 0: { v: 1 }, 1: { v: 2 } },
      1: {
        0: { s: { bl: BooleanNumber.TRUE } },
        1: { p: { id: 'doc', documentStyle: {}, body: { dataStream: 'text' } } },
      },
    }
    const result = await writeWorkbookFromSource(data, source)
    expect(result.disposition.spliced.names).toEqual([STYLES, SHEET])
  })

  it.each([
    ['a key outside the sheet address space', 'sheet-1!0:-1', 'is not a sheet-1 cell key'],
    ['a key without the sheet prefix', 'other!0:0', 'is not a sheet-1 cell key'],
    ['a key without a separator', 'sheet-1', 'is not a sheet-1 cell key'],
    ['a key for a cell the source lacks', 'sheet-1!9:9', 'does not carry'],
  ])('refuses an anchor with %s', async (_label, key, message) => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    withSourceState(data, (state) => {
      if (state.anchorMode !== 'per-unit') throw new Error('fixture has no per-unit anchors')
      return { ...state, anchors: state.anchors.map(anchor => ({ ...anchor, key })) }
    })
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain(message)
  })

  it('refuses an anchor whose span or digest no longer matches the source', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const shifts = [
      (anchor: UnitAnchor): UnitAnchor => ({ ...anchor, start: anchor.start + 1 }),
      (anchor: UnitAnchor): UnitAnchor => ({ ...anchor, end: anchor.end + 1 }),
      (anchor: UnitAnchor): UnitAnchor => ({ ...anchor, digest: 'b'.repeat(64) }),
    ]
    for (const shift of shifts) {
      const data = await readWorkbookWithSource(source, 'xlsx')
      withSourceState(data, (state) => {
        if (state.anchorMode !== 'per-unit') throw new Error('fixture has no per-unit anchors')
        return { ...state, anchors: state.anchors.map(shift) }
      })
      const error = await refusalOf(writeWorkbookFromSource(data, source))
      expect(error.code).toBe('conversion-failed')
      expect(error.message).toContain('does not match')
    }
  })

  it('skips summary anchors the source model does not carry', async () => {
    const columns = Array.from({ length: 20 }, (_, index) => {
      const column = String.fromCharCode(0x41 + index)
      return `<c r="${column}1"><v>${String(index + 1)}</v></c>`
    }).join('')
    const source = xlsxPackage({
      dimension: 'A1:U1',
      sheetData: `<row r="1">${columns}<c r="U1"/></row>`,
    })
    const bounds = resolveConfig({ maxSourceAnchorBytes: 1500 })
    const data = await readWorkbookWithSource(source, 'xlsx', bounds)
    expect(sourceStateOf(data)?.anchorMode).toBe('sheet-summary')
    const result = await writeWorkbookFromSource(data, source, bounds)
    expect(result.bytes.byteLength).toBeGreaterThan(0)
  })

  it('removes a source style and writes a builtin number format', async () => {    const source = xlsxPackage({
    dimension: 'A1:B1',
    sheetData: '<row r="1"><c r="A1" s="2"><v>1</v></c><c r="B1"><v>2</v></c></row>',
  })
  const data = await readWorkbookWithSource(source, 'xlsx')
  const sheet = sheetOf(data)
  const first = cellAt(sheet, 0, 0)
  const second = cellAt(sheet, 0, 1)
  if (first === undefined || second === undefined) throw new Error('fixture has no cells')
  first.s = null
  second.s = { n: { pattern: '0.00' } }
  const result = await writeWorkbookFromSource(data, source)
  const styles = strFromU8(entryOf(unzipEntries(result.bytes), STYLES))
  expect(styles).toContain('numFmtId="2"')
  })

  it('reuses a component appended for an earlier cell', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    for (const column of [0, 1]) {
      const cell = cellAt(sheet, 0, column)
      if (cell === undefined) throw new Error('fixture has no cell')
      cell.s = { bl: BooleanNumber.TRUE }
    }
    const result = await writeWorkbookFromSource(data, source)
    const styles = strFromU8(entryOf(unzipEntries(result.bytes), STYLES))
    expect(styles.match(/<font><b\/><\/font>/gu)).toHaveLength(1)
  })

  it('writes empty fills, empty protection, and default underline styles', async () => {
    const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<styleSheet xmlns="${SPREADSHEET}">`
      + '<numFmts count="1"><numFmt numFmtId="170" formatCode="0.000"/><numFmt numFmtId="2" formatCode="#,##0"/></numFmts>'
      + '<fonts count="2"><font/><font><u/><vertAlign val="baseline"/><unknown/></font></fonts>'
      + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill/></fills>'
      + '<borders count="2"><border><left style="thin"/></border><border diagonalUp="1"/></borders>'
      + '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
      + '<cellXfs count="3"><xf/><xf numFmtId="170" fontId="1" fillId="1" borderId="1"><alignment horizontal="weird" vertical="weird"/><protection/></xf><xf fontId="0" fillId="0" borderId="0"/></cellXfs>'
      + '</styleSheet>'
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" s="1"><v>1</v></c></row>',
      styles,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { ...styleOf(cell), n: { pattern: '0.0000' } }
    const result = await writeWorkbookFromSource(data, source)
    const written = strFromU8(entryOf(unzipEntries(result.bytes), STYLES))
    expect(written).toContain('<u/>')
    expect(written).toContain('<protection locked="1"/>')
  })

  it('writes empty style collections before a closing root', async () => {
    const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<styleSheet xmlns="${SPREADSHEET}"></styleSheet>`
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      styles,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    const cell = cellAt(sheet, 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    sheet.cellData = { 0: { 0: cell }, 1: { 0: { v: 'styled', s: { bl: BooleanNumber.TRUE } } } }
    const result = await writeWorkbookFromSource(data, source)
    const written = strFromU8(entryOf(unzipEntries(result.bytes), STYLES))
    expect(written).toContain('<fonts count="2">')
    expect(written).toContain('<cellXfs count="2">')
  })

  it('refuses an extLst child that would be dropped by a value edit', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v><extLst><ext uri="x"/></extLst></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 2
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('extLst')
  })

  it('writes cached formula results of every supported scalar type', async () => {
    const source = xlsxPackage({
      dimension: 'A1:E1',
      sheetData: '<row r="1">'
        + '<c r="A1" t="b"><f>TRUE()</f><v>1</v></c>'
        + '<c r="B1" t="e"><f>NA()</f><v>#N/A</v></c>'
        + '<c r="C1" t="str"><f>TEXT(1,"0")</f><v>1</v></c>'
        + '<c r="D1"><f>1+1</f><v>2</v></c>'
        + '<c r="E1"><f>1+2</f></c>'
        + '</row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    const expectations: readonly (readonly [number, boolean | string | number | undefined])[] = [
      [0, 1], [1, '#N/A'], [2, '1'], [3, 2], [4, undefined],
    ]
    for (const [column, value] of expectations) {
      const cell = cellAt(sheet, 0, column)
      if (cell === undefined) throw new Error('fixture has no cached cell')
      cell.v = value
      cell.f = `${cell.f ?? ''}+1`
    }
    const result = await writeWorkbookFromSource(data, source)
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.t).toBe(3)
    expect(cellAt(sheetOf(reread), 0, 1)?.v).toBe('#N/A')
  })

  it('writes boolean and empty cell values without a formula', async () => {
    const source = xlsxPackage({
      dimension: 'A1:C1',
      sheetData: '<row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>1</v></c><c r="C1"><v>2</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    const first = cellAt(sheet, 0, 0)
    const second = cellAt(sheet, 0, 1)
    const third = cellAt(sheet, 0, 2)
    if (first === undefined || second === undefined || third === undefined) throw new Error('fixture has no cells')
    first.v = false
    second.v = 1
    third.v = null
    const result = await writeWorkbookFromSource(data, source)
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe(false)
    expect(cellAt(sheetOf(reread), 0, 1)?.v).toBe(true)
  })

  it('reports malformed and unreadable shared-string parts', async () => {
    for (const shared of ['<sst><si ', '<sst><si>text']) {
      const source = xlsxPackage({
        dimension: 'A1',
        sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
        sharedStringsUnreferenced: shared,
      })
      const data = await readWorkbookWithSource(source, 'xlsx')
      const cell = cellAt(sheetOf(data), 0, 0)
      if (cell === undefined) throw new Error('fixture has no A1')
      cell.v = 2
      const result = await writeWorkbookFromSource(data, source)
      expect(result.bytes.byteLength).toBeGreaterThan(0)
    }

    const source = xlsxPackage({
      dimension: 'A1:B1',
      sheetData: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>2</v></c></row>',
      sharedStrings: ['plain'],
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const reference = cellAt(sheetOf(data), 0, 0)
    if (reference === undefined) throw new Error('fixture has no A1')
    const corrupted = corruptEntryCrc(source, 'xl/sharedStrings.xml')
    const archive = open(corrupted, { maxEntries: 100 })
    withSourceState(data, state => ({
      ...state,
      packageDigest: archive.source.digest,
      byteLength: archive.source.bytes.byteLength,
    }))
    reference.v = 'changed'
    const error = await refusalOf(writeWorkbookFromSource(data, corrupted))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('Cannot inspect xl/sharedStrings.xml')
  })

  it('refuses a source workbook whose sheets no longer match the snapshot', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const renamed = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      workbookSheets: '<sheet name="Other" sheetId="1" r:id="rId1"/>',
    })
    const error = await refusalOf(writeWorkbookFromSource(data, renamed))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('does not declare the sheets')
  })

  it('refuses a source whose sheet count changed after the snapshot', async () => {
    const secondSheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<worksheet xmlns="${SPREADSHEET}"><dimension ref="A1"/><sheetData>`
      + '<row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'
    const workbookSheets = '<sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Copy" sheetId="2" r:id="rId2"/>'
    const contentTypesExtra = '<Override PartName="/xl/worksheets/sheet2.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    const valid = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      workbookSheets,
      workbookRelsExtra: `<Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/>`,
      entriesExtra: { 'xl/worksheets/sheet2.xml': strToU8(secondSheet) },
      contentTypesExtra,
    })
    const duplicate = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      workbookSheets,
      workbookRelsExtra: `<Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`,
      entriesExtra: { 'xl/worksheets/sheet2.xml': strToU8(secondSheet) },
      contentTypesExtra,
    })
    const data = await readWorkbookWithSource(valid, 'xlsx')
    const archive = open(duplicate, { maxEntries: 100 })
    withSourceState(data, state => ({
      ...state,
      packageDigest: archive.source.digest,
      byteLength: archive.source.bytes.byteLength,
    }))
    const error = await refusalOf(writeWorkbookFromSource(data, duplicate))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('sheet count changed')
  })

  it('aligns rows after a middle row is deleted', async () => {
    const source = xlsxPackage({
      dimension: 'A1:A3',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>'
        + '<row r="3"><c r="A3"><v>3</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    sheetOf(data).cellData = { 0: { 0: { v: 1 } }, 1: { 0: { v: 3 } } }
    const result = await writeWorkbookFromSource(data, source)
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 1, 0)?.v).toBe(3)
    expect(cellAt(sheetOf(reread), 2, 0)).toBeUndefined()
  })

  it('aligns rows after a middle row is inserted', async () => {
    const source = xlsxPackage({
      dimension: 'A1:A2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    sheetOf(data).cellData = { 0: { 0: { v: 1 } }, 2: { 0: { v: 4 } } }
    const result = await writeWorkbookFromSource(data, source)
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 2, 0)?.v).toBe(4)
  })

  it('keeps a dimension absent and an unchanged one in place', async () => {
    const withoutDimension = xlsxPackage({
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const withoutData = await readWorkbookWithSource(withoutDimension, 'xlsx')
    const first = cellAt(sheetOf(withoutData), 0, 0)
    if (first === undefined) throw new Error('fixture has no A1')
    first.v = 2
    const withoutResult = await writeWorkbookFromSource(withoutData, withoutDimension)
    expect(strFromU8(entryOf(unzipEntries(withoutResult.bytes), SHEET))).not.toContain('<dimension')

    const equalDimension = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const equalData = await readWorkbookWithSource(equalDimension, 'xlsx')
    const equalCell = cellAt(sheetOf(equalData), 0, 0)
    if (equalCell === undefined) throw new Error('fixture has no A1')
    equalCell.v = 2
    const equalResult = await writeWorkbookFromSource(equalData, equalDimension)
    expect(strFromU8(entryOf(unzipEntries(equalResult.bytes), SHEET))).toContain('<dimension ref="A1"/>')
  })

  it('deletes the only cell and refuses to insert into an empty sheet', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    sheetOf(data).cellData = {}
    const deleted = await writeWorkbookFromSource(data, source)
    expect(deleted.bytes.byteLength).toBeGreaterThan(0)

    const empty = xlsxPackage({ sheetData: '' })
    const emptyData = await readWorkbookWithSource(empty, 'xlsx')
    sheetOf(emptyData).cellData = { 0: { 0: { v: 1 } } }
    const error = await refusalOf(writeWorkbookFromSource(emptyData, empty))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('neighbouring source anchor')
  })

  it('refuses an unmodeled cell metadata attribute', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" cm="1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 2
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('cm')
  })

  it('keeps a hyperlink cell whose document is unchanged', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>link</v></c></row>',
      sheetRels: `<Relationship Id="rId1" Type="${REL}/hyperlink" Target="https://example.invalid/a" TargetMode="External"/>`,
      extra: '<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 'changed'
    const result = await writeWorkbookFromSource(data, source)
    expect(result.bytes.byteLength).toBeGreaterThan(0)
  })

  it('writes a formula without a leading equals sign', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.f = '1+2'
    cell.v = 3
    const result = await writeWorkbookFromSource(data, source)
    const sheetXml = strFromU8(entryOf(unzipEntries(result.bytes), SHEET))
    expect(sheetXml).toContain('<f>1+2</f>')
  })

  it('refuses a saved package whose graph has a dangling relationship', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      workbookRelsExtra: '<Relationship Id="rId9" Type="http://example.invalid/x" Target="missing.xml"/>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 2
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('valid OPC graph')
  })

  it('aligns a shifted row by its cell signature', async () => {
    const source = xlsxPackage({
      dimension: 'A1:A2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    sheetOf(data).cellData = { 0: { 0: { v: 1 } }, 2: { 0: { v: 2 } } }
    const result = await writeWorkbookFromSource(data, source)
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 2, 0)?.v).toBe(2)
    expect(cellAt(sheetOf(reread), 1, 0)).toBeUndefined()
  })

  it('keeps paired rows with unchanged references while appending an inserted row', async () => {
    const source = xlsxPackage({
      dimension: 'A1:A2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    sheetOf(data).cellData = {
      0: { 0: { v: 1 } },
      1: { 0: { v: 3 } },
      3: { 0: { v: 4 } },
    }
    const result = await writeWorkbookFromSource(data, source)
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 1, 0)?.v).toBe(3)
    expect(cellAt(sheetOf(reread), 3, 0)?.v).toBe(4)
  })

  it('splices two worksheets in one save', async () => {
    const secondSheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<worksheet xmlns="${SPREADSHEET}"><dimension ref="A1"/><sheetData>`
      + '<row r="1"><c r="A1"><v>10</v></c></row></sheetData></worksheet>'
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      workbookSheets: '<sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Copy" sheetId="2" r:id="rId2"/>',
      workbookRelsExtra: `<Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/>`,
      entriesExtra: { 'xl/worksheets/sheet2.xml': strToU8(secondSheet) },
      contentTypesExtra: '<Override PartName="/xl/worksheets/sheet2.xml"'
        + ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const first = cellAt(sheetOf(data), 0, 0)
    if (first === undefined) throw new Error('fixture has no A1')
    first.v = 2
    const secondId = data.sheetOrder[1]
    const second = secondId === undefined ? undefined : data.sheets[secondId]?.cellData?.[0]?.[0]
    if (second === undefined) throw new Error('fixture has no second sheet cell')
    second.v = 20
    const result = await writeWorkbookFromSource(data, source)
    expect(result.disposition.spliced.names).toContain('xl/worksheets/sheet2.xml')
    const reread = await readWorkbook(result.bytes)
    expect(reread.sheets[secondId ?? '']?.cellData?.[0]?.[0]?.v).toBe(20)
  })

  it('refuses a rich inline-text cell whose text changes', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><r><t>a</t></r><t>b</t></is></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 'changed'
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('rich inline text')
  })

  it('reuses a source style index and a zero index when unmodeled values do not change facts', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B1',
      sheetData: '<row r="1"><c r="A1" s="2"><v>1</v></c><c r="B1"><v>2</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const first = cellAt(sheetOf(data), 0, 0)
    const second = cellAt(sheetOf(data), 0, 1)
    if (first === undefined || second === undefined) throw new Error('fixture has no cells')
    first.v = 3
    first.s = { ...styleOf(first), eastAsiaFontFamily: '' }
    second.v = 4
    second.s = { eastAsiaFontFamily: '' }
    const result = await writeWorkbookFromSource(data, source)
    const sheetXml = strFromU8(entryOf(unzipEntries(result.bytes), SHEET))
    expect(sheetXml).toContain('<c r="A1" s="2"><v>3</v></c>')
    expect(sheetXml).toContain('<c r="B1"><v>4</v></c>')
  })

  it('writes a none fill, a single underline, and applyProtection from a source xf', async () => {    const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<styleSheet xmlns="${SPREADSHEET}">`
      + '<numFmts count="0"/>'
      + '<fonts count="1"><font><u val="single"/></font></fonts>'
      + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill/></fills>'
      + '<borders count="1"><border/></borders>'
      + '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
      + '<cellXfs count="2"><xf/><xf fontId="1" applyProtection="1"><protection/></xf></cellXfs>'
      + '</styleSheet>'
  const source = xlsxPackage({
    dimension: 'A1:B1',
    sheetData: '<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
    styles,
  })
  const data = await readWorkbookWithSource(source, 'xlsx')
  const first = cellAt(sheetOf(data), 0, 0)
  const second = cellAt(sheetOf(data), 0, 1)
  if (first === undefined || second === undefined) throw new Error('fixture has no cells')
  first.s = { ...styleOf(first), n: { pattern: '0.0000' } }
  second.s = {
    ul: { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE },
    bl: BooleanNumber.TRUE,
    bg: {},
  }
  const result = await writeWorkbookFromSource(data, source)
  const written = strFromU8(entryOf(unzipEntries(result.bytes), STYLES))
  expect(written).toContain('applyProtection="1"')
  expect(written).toContain('<protection locked="1"/>')
  expect(written).toContain('patternType="none"')
  expect(written).toContain('<u/>')
  })

  it('falls back for malformed style ids and missing colour/format attributes', async () => {
    const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<styleSheet xmlns="${SPREADSHEET}">`
      + '<numFmts count="1"><numFmt numFmtId="1.5" formatCode="0.0"/></numFmts>'
      + '<fonts count="2"><font/><font><color/></font></fonts>'
      + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill/></fills>'
      + '<borders count="2"><border/><border><left style="none"/></border></borders>'
      + '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
      + '<cellXfs count="2"><xf/><xf fontId="1"><alignment/></xf></cellXfs>'
      + '</styleSheet>'
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" s="1"><v>1</v></c></row>',
      styles,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { ...styleOf(cell), bl: BooleanNumber.TRUE, n: { pattern: 'General' } }
    const result = await writeWorkbookFromSource(data, source)
    expect(result.bytes.byteLength).toBeGreaterThan(0)
  })

  it('writes vertical text, an empty border, and a general number format', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = {
      tr: { a: 0, v: BooleanNumber.TRUE },
      va: 1 as BaselineOffset,
      bd: { t: { s: 0 as BorderStyleTypes, cl: { rgb: '#000000' } } },
      n: { pattern: 'General' },
    }
    const result = await writeWorkbookFromSource(data, source)
    expect(result.bytes.byteLength).toBeGreaterThan(0)
  })

  it('refuses a changed hyperlink document', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>link</v></c></row>',
      sheetRels: `<Relationship Id="rId1" Type="${REL}/hyperlink" Target="https://example.invalid/a" TargetMode="External"/>`,
      extra: '<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.p = {
      id: 'cell-doc',
      documentStyle: {},
      body: {
        dataStream: 'link\r\n',
        customRanges: [{
          startIndex: 0,
          endIndex: 3,
          rangeId: 'link',
          rangeType: 0,
          properties: { url: 'https://example.invalid/other' },
        }],
      },
    }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('changes its hyperlink')
  })

  it('renders two inserted rows in index order', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    sheetOf(data).cellData = {
      0: { 0: { v: 1 } },
      1: { 0: { v: 2 } },
      2: { 0: { v: 3 } },
    }
    const result = await writeWorkbookFromSource(data, source)
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 2, 0)?.v).toBe(3)
  })

  it('refuses an unreadable styles part only when an edit needs it', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const corrupted = corruptEntryCrc(source, STYLES)
    const data = await readWorkbookWithSource(source, 'xlsx')
    const archive = open(corrupted, { maxEntries: 100 })
    withSourceState(data, state => ({
      ...state,
      packageDigest: archive.source.digest,
      byteLength: archive.source.bytes.byteLength,
    }))
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { bl: BooleanNumber.TRUE }
    const error = await refusalOf(writeWorkbookFromSource(data, corrupted))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('Cannot read xl/styles.xml')
  })

  it('refuses a style-only edit on a rich-text shared string instead of flattening it', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" t="s"><v>0</v></c></row>',
      richSharedStrings: ['<si><r><rPr><b/></rPr><t>Rich</t></r></si>'],
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.s = { bl: BooleanNumber.TRUE }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('rich-text')
    expect(error.message).toContain('A1')
  })

  it('refuses a row insert on a sheet whose merged range would keep stale coordinates', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>',
      extra: '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    sheet.cellData = {
      0: sheet.cellData?.[0] ?? {},
      1: { 0: { v: 'inserted', t: 1 } },
      2: sheet.cellData?.[1] ?? {},
    }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('merged range A1:B1')
    expect(error.message).toContain(SHEET)
  })

  it('refuses a row delete on a sheet whose merged range would keep stale coordinates', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B3',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>'
        + '<row r="3"><c r="A3"><v>3</v></c></row>',
      extra: '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    sheet.cellData = {
      0: sheet.cellData?.[1] ?? {},
      1: sheet.cellData?.[2] ?? {},
    }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('merged range A1:B1')
  })

  it('refuses a column insert on a sheet whose conditional formatting range would keep stale coordinates', async () => {
    const source = xlsxPackage({
      dimension: 'A1:C1',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c>'
        + '<c r="B1" t="inlineStr"><is><t>b</t></is></c>'
        + '<c r="C1" t="inlineStr"><is><t>c</t></is></c></row>',
      extra: '<conditionalFormatting sqref="A1:C1">'
        + '<cfRule type="expression" priority="1"><formula>1</formula></cfRule>'
        + '</conditionalFormatting>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    const row = sheet.cellData?.[0]
    if (row === undefined) throw new Error('fixture has no row 1')
    sheet.cellData = {
      0: {
        0: row[0] ?? {},
        1: { v: 'x', t: 1 },
        2: row[1] ?? {},
        3: row[2] ?? {},
      },
    }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('conditional formatting range A1:C1')
  })

  it('refuses a row insert on a sheet whose data validation range would keep stale coordinates', async () => {
    const source = xlsxPackage({
      dimension: 'A1:A3',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>'
        + '<row r="3"><c r="A3"><v>3</v></c></row>',
      extra: '<dataValidations count="1">'
        + '<dataValidation type="list" allowBlank="1" sqref="A1:A3"><formula1>"a,b"</formula1></dataValidation>'
        + '</dataValidations>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    sheet.cellData = {
      0: sheet.cellData?.[0] ?? {},
      1: { 0: { v: 'inserted', t: 1 } },
      2: sheet.cellData?.[1] ?? {},
      3: sheet.cellData?.[2] ?? {},
    }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('data validation range A1:A3')
  })

  it('refuses a row insert on a sheet whose auto filter range would keep stale coordinates', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>'
        + '<row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>',
      extra: '<autoFilter ref="A1:B2"/>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    sheet.cellData = {
      0: sheet.cellData?.[0] ?? {},
      1: { 0: { v: 5, t: 2 }, 1: { v: 6, t: 2 } },
      2: sheet.cellData?.[1] ?? {},
    }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('auto filter range A1:B2')
    expect(error.message).toContain(SHEET)
  })

  it('refuses a row insert on a sheet whose table range would keep stale coordinates', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B3',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>Label</t></is></c>'
        + '<c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row>'
        + '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>10</v></c></row>'
        + '<row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>20</v></c></row>',
      extra: '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
      sheetRels: `<Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/>`,
      contentTypesExtra: '<Override PartName="/xl/tables/table1.xml"'
        + ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
      entriesExtra: {
        'xl/tables/table1.xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          + `<table xmlns="${SPREADSHEET}" id="1" name="DataTable" displayName="DataTable" ref="A1:B3" totalsRowShown="0">`
          + '<autoFilter ref="A1:B3"/>'
          + '<tableColumns count="2"><tableColumn id="1" name="Label"/><tableColumn id="2" name="Amount"/></tableColumns>'
          + '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0"'
          + ' showRowStripes="1" showColumnStripes="0"/>'
          + '</table>'),
      },
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    sheet.cellData = {
      0: sheet.cellData?.[0] ?? {},
      1: { 0: { v: 'inserted', t: 1 }, 1: { v: 30, t: 2 } },
      2: sheet.cellData?.[1] ?? {},
      3: sheet.cellData?.[2] ?? {},
    }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('table range A1:B3')
    expect(error.message).toContain(SHEET)
  })

  it('refuses a row insert when a declared table part is absent from the package', async () => {
    const source = xlsxPackage({
      dimension: 'A1:A2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>'
        + '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
      sheetRels: `<Relationship Id="rId1" Type="${REL}/customXml" Target="../tables/table1.xml"/>`,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    sheet.cellData = {
      0: sheet.cellData?.[0] ?? {},
      1: { 0: { v: 'inserted', t: 1 } },
      2: sheet.cellData?.[1] ?? {},
    }
    const original = source.slice()
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('table relationship rId1 (xl/tables/table1.xml)')
    expect(error.message).toContain('has no resolvable table range')
    expect(error.message).toContain(SHEET)
    expect(source).toEqual(original)
  })

  it('refuses a row insert when a declared table target names no package part', async () => {
    const source = xlsxPackage({
      dimension: 'A1:A2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>'
        + '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
      sheetRels: `<Relationship Id="rId1" Type="${REL}/customXml" Target="https://example.invalid/table1.xml"/>`,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    sheet.cellData = {
      0: sheet.cellData?.[0] ?? {},
      1: { 0: { v: 'inserted', t: 1 } },
      2: sheet.cellData?.[1] ?? {},
    }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain(`table relationship rId1 in ${SHEET}`)
    expect(error.message).toContain('has no resolvable table range')
  })

  it('refuses a column insert when a declared table part carries no ref', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
      extra: '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
      sheetRels: `<Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/>`,
      contentTypesExtra: '<Override PartName="/xl/tables/table1.xml"'
        + ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
      entriesExtra: {
        'xl/tables/table1.xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          + `<table xmlns="${SPREADSHEET}" id="1" name="DataTable" displayName="DataTable" totalsRowShown="0">`
          + '<tableColumns count="2"><tableColumn id="1" name="Label"/><tableColumn id="2" name="Amount"/></tableColumns>'
          + '</table>'),
      },
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    const row = sheet.cellData?.[0]
    if (row === undefined) throw new Error('fixture has no row 1')
    sheet.cellData = {
      0: {
        0: row[0] ?? {},
        1: { v: 'inserted', t: 1 },
        2: row[1] ?? {},
      },
    }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('table relationship rId1 (xl/tables/table1.xml)')
    expect(error.message).toContain('has no resolvable table range')
  })

  it('saves a value edit when a declared table part carries no ref', async () => {
    const source = xlsxPackage({
      dimension: 'A1:A2',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c></row>',
      extra: '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
      sheetRels: `<Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/>`,
      contentTypesExtra: '<Override PartName="/xl/tables/table1.xml"'
        + ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
      entriesExtra: {
        'xl/tables/table1.xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          + `<table xmlns="${SPREADSHEET}" id="1" name="DataTable" displayName="DataTable" totalsRowShown="0">`
          + '<tableColumns count="2"><tableColumn id="1" name="Label"/><tableColumn id="2" name="Amount"/></tableColumns>'
          + '</table>'),
      },
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 3
    const result = await writeWorkbookFromSource(data, source)
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe(3)
  })

  it('refuses a merged range when the snapshot omits its merge metadata', async () => {
    const source = xlsxPackage({
      dimension: 'A1:B1',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>Merged</t></is></c></row>',
      extra: '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    delete sheet.mergeData
    const cell = cellAt(sheet, 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 'Updated'
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('merged range A1:B1')
  })

  it('ignores merge slaves when checking a vertical merge for a structural shift', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>Merged</t></is></c></row>',
      extra: '<mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 'Updated'
    const result = await writeWorkbookFromSource(data, source)
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe('Updated')
    expect(sheetOf(reread).mergeData).toEqual([{ startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 }])
  })

  it('rewrites a row-shifted cell on a vertical-merge sheet that is not refused', async () => {
    const source = xlsxPackage({
      dimension: 'A1:A2',
      sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>Merged</t></is></c></row>',
      extra: '<mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const sheet = sheetOf(data)
    const original = sheet.cellData
    sheet.cellData = {
      0: {},
      1: original?.[0] ?? {},
      2: original?.[1] ?? {},
    }
    const result = await writeWorkbookFromSource(data, source)
    const sheetXml = strFromU8(entryOf(unzipEntries(result.bytes), SHEET))
    expect(rowCellReferences(sheetXml)).toEqual(['row 2', 'A2', 'row 3', 'A3'])
    expect(sheetXml).not.toContain('<c r="A1"')

    const rereadSheet = sheetOf(await readWorkbook(result.bytes))
    expect(cellAt(rereadSheet, 1, 0)?.v).toBeUndefined()
    expect(cellAt(rereadSheet, 2, 0)?.v).toBe('Merged')
    expect(rereadSheet.mergeData).toEqual([{ startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 }])
  })

  it('escapes a string value whose model claims a numeric cell type', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = '</v><c r="Z9">'
    cell.t = 2
    const result = await writeWorkbookFromSource(data, source)
    const sheetXml = strFromU8(entryOf(unzipEntries(result.bytes), SHEET))
    expect(sheetXml.match(/<c\b/gu)).toHaveLength(1)
    expect(sheetXml).not.toMatch(/<c\b[^>]*\br="Z9"/u)
    expect(sheetXml).toContain('&lt;/v&gt;&lt;c r="Z9"&gt;')
    const reread = await readWorkbook(result.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe('</v><c r="Z9">')
  })

  it('aligns a few thousand rows with one map-backed offset scan', { timeout: 30_000 }, async () => {
    const rowCount = 3000
    const rows = Array.from(
      { length: rowCount },
      (_, index) => `<row r="${String(index + 1)}"><c r="A${String(index + 1)}"><v>${String(index + 1)}</v></c></row>`,
    ).join('')
    const source = xlsxPackage({
      dimension: `A1:A${String(rowCount)}`,
      sheetData: rows,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cellData: Record<number, Record<number, ICellData>> = { 0: { 0: { v: 'inserted', t: 1 } } }
    for (let index = 1; index <= rowCount; index += 1) {
      cellData[index] = { 0: { v: 100 + index, t: 2 } }
    }
    sheetOf(data).cellData = cellData
    const result = await writeWorkbookFromSource(data, source)
    const reread = await readWorkbook(result.bytes)
    const readSheet = sheetOf(reread)
    expect(cellAt(readSheet, 0, 0)?.v).toBe('inserted')
    expect(cellAt(readSheet, 1, 0)?.v).toBe(101)
    expect(cellAt(readSheet, rowCount, 0)?.v).toBe(100 + rowCount)
    expect(cellAt(readSheet, rowCount + 1, 0)).toBeUndefined()
  })

  it('re-anchors a stale snapshot whose spans exceed a shortened second save', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: `<row r="1"><c r="A1" t="inlineStr"><is><t>${'A'.repeat(400)}</t></is></c></row>`,
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 'short'
    const first = await writeWorkbookFromSource(data, source)
    const second = await writeWorkbookFromSource(data, first.bytes)
    expect(second.bytes).toEqual(first.bytes)
    const reread = await readWorkbook(second.bytes)
    expect(cellAt(sheetOf(reread), 0, 0)?.v).toBe('short')
  })

  it('re-anchors a sidecar whose recorded byte length alone is stale', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    withSourceState(data, state => ({ ...state, byteLength: state.byteLength + 1 }))
    const cell = cellAt(sheetOf(data), 0, 0)
    if (cell === undefined) throw new Error('fixture has no A1')
    cell.v = 2
    const result = await writeWorkbookFromSource(data, source)
    expect(strFromU8(entryOf(unzipEntries(result.bytes), SHEET))).toContain('<v>2</v>')
  })

  it('refuses a sidecar source value that is not an object', async () => {
    const source = xlsxPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
    })
    const data = await readWorkbookWithSource(source, 'xlsx')
    ;(data as SidecarCarrier).officeExchange = { source: 'malformed' }
    const error = await refusalOf(writeWorkbookFromSource(data, source))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('Workbook snapshot source state is malformed')
  })
})
