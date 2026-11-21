/**
 * In-place OOXML preservation across the fixture corpus.
 *
 * Every OOXML fixture except the documented whole-file refusal cases is read
 * and written through the real provider: an unedited snapshot must reproduce
 * the source bytes exactly, a targeted edit must leave every untouched entry
 * identical and every byte outside the edited unit span exact, and the save
 * disposition must account for every entry the independent diff reports.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { IDocumentData, IParagraph, IWorkbookData } from '@univerjs/core'
import { Context } from '@forwork/cordis'
import ExcelJS from 'exceljs'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import LocalOfficeExchange from '../src/provider/index.ts'
import { open } from '../src/package/ooxml-package.ts'
import type { SourcePackage } from '../src/package/types.ts'
import type { OfficeSourceState, UnitAnchor } from '../src/source/source-state.ts'
import type { OfficeWriteResult, UnitSnapshot } from '../src/snapshots.ts'
import type { EditableOfficeFormat, SaveDisposition } from '../src/types.ts'
import {
  comparePackages,
  formatDiffReport,
  unzipEntries,
  validateOpc,
  type EntryStatus,
  type PackageDiff,
} from './ooxml-package.ts'

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url))
const DOCUMENT_PART = 'word/document.xml'
const CONTENT_TYPES_PART = '[Content_Types].xml'
const OOXML_EXTENSIONS: readonly string[] = ['.docx', '.docm', '.xlsx', '.xlsm']
const KNOWN_REFUSAL_FIXTURES: ReadonlySet<string> = new Set([
  'unsupported-macro.docx',
  'unsupported-tracked.docx',
])
const MACRO_CONTENT_TYPES: Readonly<Record<string, string>> = {
  'macro.docm': 'application/vnd.ms-word.document.macroEnabled.main+xml',
  'macro.xlsm': 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
}
const PERF_CELLS = 10_000
const PERF_COLUMNS = 50
const PERF_OPAQUE_BYTES = 4_500_000
const PERF_SAVE_BUDGET_MS = 30_000
const PERF_TIMEOUT_MS = 120_000

/** Every OOXML package in the fixture directory, sorted by name. */
const ALL_OOXML_FIXTURES: readonly string[] = readdirSync(FIXTURES)
  .filter(name => OOXML_EXTENSIONS.some(extension => name.endsWith(extension)))
  .sort()

/** Fixtures the save corpus drives; the documented refusal cases stay excluded. */
const CORPUS: readonly string[] = ALL_OOXML_FIXTURES.filter(name => !KNOWN_REFUSAL_FIXTURES.has(name))

/** Representative corpus fixtures whose disposition must match the independent diff. */
const DISPOSITION_FIXTURES: readonly string[] = [
  'roundtrip.docx',
  'preserve-image.docx',
  'preserve-chart.docx',
  'roundtrip.xlsx',
  'preserve-pivot-table.xlsx',
  'macro.docm',
]

/**
 * FR-004 construct classes mapped to the fixtures that document them.
 *
 * Tracked changes are documented by `unsupported-tracked.docx`: the class is
 * part of the corpus evidence while the save matrix excludes that file because
 * reading it reports a whole-file refusal case.
 */
const FR004_CLASSES: Readonly<Record<string, readonly string[]>> = {
  macros: ['macro.docm', 'macro.xlsm'],
  charts: ['preserve-chart.docx'],
  'drawings and shapes': ['preserve-image.docx', 'preserve-shape.docx'],
  comments: ['preserve-comment.docx'],
  'threaded comments': ['preserve-comment.docx'],
  'embedded objects': ['preserve-embedded-object.docx'],
  'pivot tables and caches': ['preserve-pivot-table.xlsx'],
  'tracked changes': ['unsupported-tracked.docx'],
  'Excel tables': ['preserve-excel-table.xlsx'],
  'defined names': ['preserve-defined-names.xlsx'],
  'calc chains': ['preserve-calc-chain.xlsx'],
  'custom XML': ['preserve-unmodeled-parts.docx'],
  'external links': ['preserve-external-link.docx'],
  themes: ['preserve-unmodeled-parts.docx'],
  'font tables': ['preserve-unmodeled-parts.docx'],
  'web settings': ['preserve-unmodeled-parts.docx'],
  'document properties': ['preserve-unmodeled-parts.docx'],
  fields: ['preserve-field.docx'],
  'content controls': ['preserve-content-control.docx'],
  bookmarks: ['preserve-bookmark.docx'],
  'other package entries': ['container-stored.docx', 'container-descriptor.docx', 'container-comment.docx'],
}

/** The edited unit's source location, used for the byte-level span assertion. */
interface EditEvidence {
  /** Part the edit changes. */
  readonly part: string
  /** Edited unit's source byte span, for a replacement or in-paragraph insertion. */
  readonly span?: { readonly start: number; readonly end: number }
  /** Source byte offset receiving a new block, for a paragraph insertion. */
  readonly insertionAt?: number
}

/** One fixture edit the corpus matrix applies to the snapshot. */
type EditPlan =
  | { readonly kind: 'docx-text'; readonly from: string; readonly to: string }
  | { readonly kind: 'docx-text-at-unit'; readonly anchorKey: string; readonly text: string }
  | { readonly kind: 'docx-insert-paragraph'; readonly text: string }
  | {
    readonly kind: 'xlsx-cell'
    readonly sheetId: string
    readonly row: number
    readonly column: number
    readonly value: string | number
  }

const EDITS: Readonly<Record<string, EditPlan>> = {
  'roundtrip.docx': { kind: 'docx-text', from: 'Quarterly', to: 'Financial' },
  'container-comment.docx': { kind: 'docx-text', from: 'Quarterly', to: 'Financial' },
  'container-stored.docx': { kind: 'docx-text', from: 'Quarterly', to: 'Financial' },
  'container-descriptor.docx': { kind: 'docx-text', from: 'Quarterly', to: 'Financial' },
  'unsupported-header.docx': { kind: 'docx-text', from: 'Quarterly', to: 'Financial' },
  'offset-shift.docx': { kind: 'docx-text', from: 'First paragraph', to: 'Primary paragraph' },
  'preserve-image.docx': { kind: 'docx-text', from: 'after image', to: 'later image' },
  'preserve-external-image.docx': { kind: 'docx-text', from: 'after image', to: 'later image' },
  'preserve-notes.docx': { kind: 'docx-text', from: 'Body text', to: 'Document text' },
  'preserve-note-image.docx': { kind: 'docx-text', from: 'Body text', to: 'Document text' },
  'preserve-header-image.docx': { kind: 'docx-text', from: 'Header image fixture', to: 'Masthead image fixture' },
  'preserve-footer-image.docx': { kind: 'docx-text', from: 'Footer image fixture', to: 'Page footer fixture' },
  'preserve-multi-media.docx': { kind: 'docx-text', from: 'Before media', to: 'Ahead of media' },
  'preserve-id-collision.docx': { kind: 'docx-text', from: 'Collision ', to: 'Collapse ' },
  'preserve-chart.docx': { kind: 'docx-text', from: 'Chart ', to: 'Graph ' },
  'preserve-embedded-object.docx': { kind: 'docx-text', from: 'embedded object', to: 'embedded package' },
  'preserve-external-link.docx': { kind: 'docx-text', from: 'the external reference', to: 'the linked reference' },
  'preserve-field.docx': { kind: 'docx-text', from: '2026-01-01', to: '2027-02-02' },
  'preserve-bookmark.docx': { kind: 'docx-text', from: 'Bookmarked text', to: 'Anchored text' },
  'preserve-unmodeled-parts.docx': { kind: 'docx-text', from: 'Unmodeled parts fixture', to: 'Carried parts fixture' },
  'macro.docm': { kind: 'docx-text', from: 'Quarterly', to: 'Financial' },
  'preserve-shape.docx': { kind: 'docx-text-at-unit', anchorKey: 'para_1', text: 'Shape caption' },
  'preserve-table-cell-image.docx': { kind: 'docx-text-at-unit', anchorKey: 'table-1/0/0', text: 'Cell caption' },
  'preserve-content-control.docx': { kind: 'docx-insert-paragraph', text: 'Inserted paragraph' },
  'preserve-comment.docx': { kind: 'docx-insert-paragraph', text: 'Inserted paragraph' },
  'roundtrip.xlsx': { kind: 'xlsx-cell', sheetId: 'sheet-1', row: 1, column: 0, value: 'Sprocket' },
  'macro.xlsm': { kind: 'xlsx-cell', sheetId: 'sheet-1', row: 1, column: 0, value: 'Sprocket' },
  'preserve-pivot-table.xlsx': { kind: 'xlsx-cell', sheetId: 'sheet-1', row: 1, column: 0, value: 'Gamma' },
  'preserve-excel-table.xlsx': { kind: 'xlsx-cell', sheetId: 'sheet-1', row: 1, column: 0, value: 'Gamma' },
  'preserve-defined-names.xlsx': { kind: 'xlsx-cell', sheetId: 'sheet-1', row: 1, column: 1, value: 11 },
  'preserve-calc-chain.xlsx': { kind: 'xlsx-cell', sheetId: 'sheet-1', row: 0, column: 0, value: 5 },
}

/** One preserved construct's survival evidence. */
interface ConstructCase {
  /** Construct class reported in the test name. */
  readonly label: string
  /** Fixture carrying the construct. */
  readonly fixture: string
  /** Source entries the edit must reproduce byte-for-byte. */
  readonly identicalEntries: readonly string[]
  /** Source elements whose exact XML must appear in the written part. */
  readonly fragments?: readonly {
    readonly part: string
    readonly element: string
    /** `element` compares the whole element; `openTag` compares only its start tag. */
    readonly mode?: 'element' | 'openTag'
  }[]
}

const CONSTRUCT_CASES: readonly ConstructCase[] = [
  {
    label: 'macro project',
    fixture: 'macro.docm',
    identicalEntries: ['word/vbaProject.bin', 'word/styles.xml'],
  },
  {
    label: 'macro project',
    fixture: 'macro.xlsm',
    identicalEntries: ['xl/vbaProject.bin', 'xl/workbook.xml', 'xl/styles.xml'],
  },
  {
    label: 'chart',
    fixture: 'preserve-chart.docx',
    identicalEntries: ['word/charts/chart1.xml', 'word/_rels/document.xml.rels'],
    fragments: [{ part: DOCUMENT_PART, element: 'c:chart' }],
  },
  {
    label: 'pivot table and cache',
    fixture: 'preserve-pivot-table.xlsx',
    identicalEntries: [
      'xl/pivotTables/pivotTable1.xml',
      'xl/pivotCache/pivotCacheDefinition1.xml',
      'xl/pivotCache/pivotCacheRecords1.xml',
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
      'xl/worksheets/_rels/sheet1.xml.rels',
    ],
  },
  {
    label: 'Excel table',
    fixture: 'preserve-excel-table.xlsx',
    identicalEntries: ['xl/tables/table1.xml', 'xl/worksheets/_rels/sheet1.xml.rels'],
  },
  {
    label: 'defined names',
    fixture: 'preserve-defined-names.xlsx',
    identicalEntries: ['xl/workbook.xml'],
  },
  {
    label: 'calc chain',
    fixture: 'preserve-calc-chain.xlsx',
    identicalEntries: ['xl/calcChain.xml', 'xl/_rels/workbook.xml.rels'],
  },
  {
    label: 'threaded comment parts and range',
    fixture: 'preserve-comment.docx',
    identicalEntries: [
      'word/comments.xml',
      'word/commentsExtended.xml',
      'word/commentsIds.xml',
      'word/people.xml',
      'word/_rels/document.xml.rels',
    ],
    fragments: [
      { part: DOCUMENT_PART, element: 'w:commentRangeStart' },
      { part: DOCUMENT_PART, element: 'w:commentRangeEnd' },
      { part: DOCUMENT_PART, element: 'w:commentReference' },
    ],
  },
  {
    label: 'embedded object',
    fixture: 'preserve-embedded-object.docx',
    identicalEntries: [
      'word/embeddings/oleObject1.bin',
      'word/media/ole-preview.png',
      'word/_rels/document.xml.rels',
    ],
    fragments: [{ part: DOCUMENT_PART, element: 'w:object' }],
  },
  {
    label: 'shape',
    fixture: 'preserve-shape.docx',
    identicalEntries: ['word/drawings/vmlDrawing1.vml', 'word/_rels/document.xml.rels'],
    fragments: [{ part: DOCUMENT_PART, element: 'w:pict' }],
  },
  {
    label: 'theme, font table, web settings, custom XML, and document properties',
    fixture: 'preserve-unmodeled-parts.docx',
    identicalEntries: [
      'word/theme/theme1.xml',
      'word/fontTable.xml',
      'word/webSettings.xml',
      'customXml/item1.xml',
      'customXml/itemProps1.xml',
      'customXml/_rels/item1.xml.rels',
      'docProps/core.xml',
      'docProps/app.xml',
      '_rels/.rels',
    ],
  },
  {
    label: 'external link',
    fixture: 'preserve-external-link.docx',
    identicalEntries: ['word/_rels/document.xml.rels'],
    fragments: [{ part: DOCUMENT_PART, element: 'w:hyperlink', mode: 'openTag' }],
  },
  {
    label: 'field',
    fixture: 'preserve-field.docx',
    identicalEntries: ['word/styles.xml'],
    fragments: [
      { part: DOCUMENT_PART, element: 'w:fldSimple' },
      { part: DOCUMENT_PART, element: 'w:fldChar' },
      { part: DOCUMENT_PART, element: 'w:instrText' },
    ],
  },
  {
    label: 'bookmark',
    fixture: 'preserve-bookmark.docx',
    identicalEntries: ['word/styles.xml'],
    fragments: [
      { part: DOCUMENT_PART, element: 'w:bookmarkStart' },
      { part: DOCUMENT_PART, element: 'w:bookmarkEnd' },
    ],
  },
  {
    label: 'content control',
    fixture: 'preserve-content-control.docx',
    identicalEntries: ['word/styles.xml'],
    fragments: [{ part: DOCUMENT_PART, element: 'w:sdt' }],
  },
  {
    label: 'inline drawing',
    fixture: 'preserve-image.docx',
    identicalEntries: ['word/media/image1.jpeg', 'word/_rels/document.xml.rels'],
    fragments: [{ part: DOCUMENT_PART, element: 'w:drawing' }],
  },
  {
    label: 'linked drawing',
    fixture: 'preserve-external-image.docx',
    identicalEntries: ['word/_rels/document.xml.rels'],
    fragments: [{ part: DOCUMENT_PART, element: 'w:drawing' }],
  },
  {
    label: 'header image',
    fixture: 'preserve-header-image.docx',
    identicalEntries: ['word/header1.xml', 'word/media/header1.png', 'word/_rels/header1.xml.rels'],
  },
  {
    label: 'footer image',
    fixture: 'preserve-footer-image.docx',
    identicalEntries: ['word/footer1.xml', 'word/media/footer1.png', 'word/_rels/footer1.xml.rels'],
  },
  {
    label: 'note image',
    fixture: 'preserve-note-image.docx',
    identicalEntries: [
      'word/media/note1.png',
      'word/footnotes.xml',
      'word/endnotes.xml',
      'word/_rels/footnotes.xml.rels',
    ],
  },
  {
    label: 'multi-media drawings',
    fixture: 'preserve-multi-media.docx',
    identicalEntries: [
      'word/media/image1.png',
      'word/media/image2.jpeg',
      'word/media/image3.emf',
      'word/_rels/document.xml.rels',
    ],
  },
  {
    label: 'table-cell drawing',
    fixture: 'preserve-table-cell-image.docx',
    identicalEntries: ['word/media/cell1.png', 'word/_rels/document.xml.rels'],
    fragments: [{ part: DOCUMENT_PART, element: 'w:drawing' }],
  },
  {
    label: 'offset drawing',
    fixture: 'offset-shift.docx',
    identicalEntries: ['word/media/offset1.png', 'word/_rels/document.xml.rels'],
    fragments: [{ part: DOCUMENT_PART, element: 'w:drawing' }],
  },
  {
    label: 'note parts and references',
    fixture: 'preserve-notes.docx',
    identicalEntries: ['word/footnotes.xml', 'word/endnotes.xml', 'word/_rels/document.xml.rels'],
    fragments: [
      { part: DOCUMENT_PART, element: 'w:footnoteReference' },
      { part: DOCUMENT_PART, element: 'w:endnoteReference' },
    ],
  },
  {
    label: 'relationship id collision',
    fixture: 'preserve-id-collision.docx',
    identicalEntries: ['word/header1.xml', 'word/media/collision1.png', 'word/_rels/document.xml.rels'],
  },
]

interface SidecarCarrier {
  officeExchange?: {
    readonly source?: unknown
  }
}

interface FixtureRead {
  /** Complete source package bytes. */
  readonly bytes: Uint8Array
  /** Indexed source package supplied to the writer. */
  readonly source: SourcePackage
  /** Provider read result. */
  readonly snapshot: UnitSnapshot
  /** Detected source format. */
  readonly format: EditableOfficeFormat
}

let ctx: Context

beforeAll(async () => {
  ctx = new Context()
  await ctx.plugin(LocalOfficeExchange)
})

afterAll(async () => {
  await ctx.fiber.dispose()
})

function formatOf(name: string): EditableOfficeFormat {
  if (name.endsWith('.docx')) return 'docx'
  if (name.endsWith('.docm')) return 'docm'
  if (name.endsWith('.xlsx')) return 'xlsx'
  return 'xlsm'
}

async function readFixture(name: string): Promise<FixtureRead> {
  const path = join(FIXTURES, name)
  const bytes = new Uint8Array(readFileSync(path))
  const format = formatOf(name)
  const read = await ctx.officeExchange.read(path)
  expect(read.format).toBe(format)
  return {
    bytes,
    source: open(bytes, { maxEntries: 10_000 }).source,
    snapshot: read.snapshot,
    format,
  }
}

function sourceStateOf(data: object): OfficeSourceState | undefined {
  return (data as SidecarCarrier).officeExchange?.source as OfficeSourceState | undefined
}

function anchorsOf(data: object): readonly UnitAnchor[] {
  const state = sourceStateOf(data)
  if (state?.anchorMode !== 'per-unit') throw new Error('snapshot carries no per-unit source anchors')
  return state.anchors
}

function anchorOf(data: object, part: string, key: string): UnitAnchor {
  const anchor = anchorsOf(data).find(candidate => candidate.part === part && candidate.key === key)
  if (anchor === undefined) throw new Error(`snapshot has no anchor ${key} in ${part}`)
  return anchor
}

function anchorByKey(data: object, key: string): UnitAnchor {
  const anchor = anchorsOf(data).find(candidate => candidate.key === key)
  if (anchor === undefined) throw new Error(`snapshot has no anchor ${key}`)
  return anchor
}

function bodyOf(data: IDocumentData): NonNullable<IDocumentData['body']> {
  if (data.body === undefined) throw new Error('document snapshot has no body')
  return data.body
}

function contentStartOf(stream: string, breakIndex: number): number {
  let index = breakIndex
  while (index > 0) {
    const previous = stream[index - 1]
    if (previous === '\r' || previous === '\n' || previous === '\x0F' || previous === '\x1C') break
    index -= 1
  }
  return index
}

function spliceStory(body: NonNullable<IDocumentData['body']>, start: number, end: number, text: string): void {
  const stream = body.dataStream ?? ''
  const delta = text.length - (end - start)
  const deltaStart = start + text.length
  body.dataStream = stream.slice(0, start) + text + stream.slice(end)
  for (const paragraph of body.paragraphs ?? []) {
    if (paragraph.startIndex >= end) paragraph.startIndex += delta
  }
  for (const section of body.sectionBreaks ?? []) {
    if (section.startIndex >= end) section.startIndex += delta
  }
  for (const range of body.tables ?? []) {
    if (range.startIndex >= end) range.startIndex += delta
    if (range.endIndex > end) range.endIndex += delta
  }
  for (const block of body.docxRawBlocks ?? []) {
    if (block.startIndex >= end) block.startIndex += delta
  }
  body.textRuns = (body.textRuns ?? []).flatMap((run) => {
    if (run.ed <= start) return [run]
    if (run.st >= end) return [{ ...run, st: run.st + delta, ed: run.ed + delta }]
    const clippedStart = run.st < start ? run.st : start
    const clippedEnd = run.ed > end ? run.ed + delta : deltaStart
    return clippedEnd > clippedStart ? [{ ...run, st: clippedStart, ed: clippedEnd }] : []
  })
}

function paragraphContaining(data: IDocumentData, offset: number): IParagraph {
  const paragraphs = (bodyOf(data).paragraphs ?? [])
    .filter(paragraph => paragraph.startIndex >= offset)
    .sort((left, right) => left.startIndex - right.startIndex)
  const paragraph = paragraphs[0]
  if (paragraph === undefined) throw new Error(`no paragraph ends after stream offset ${String(offset)}`)
  return paragraph
}

function replaceDocumentText(data: IDocumentData, from: string, to: string): EditEvidence {
  const body = bodyOf(data)
  const start = (body.dataStream ?? '').indexOf(from)
  if (start < 0) throw new Error(`fixture text "${from}" is absent`)
  const paragraph = paragraphContaining(data, start)
  const anchor = anchorOf(data, DOCUMENT_PART, paragraph.paragraphId)
  spliceStory(body, start, start + from.length, to)
  return { part: anchor.part, span: { start: anchor.start, end: anchor.end } }
}

function insertTextAtUnit(data: IDocumentData, anchorKey: string, text: string): EditEvidence {
  const body = bodyOf(data)
  const anchor = anchorOf(data, DOCUMENT_PART, anchorKey)
  if (anchor.kind === 'paragraph') {
    const paragraph = (body.paragraphs ?? []).find(candidate => candidate.paragraphId === anchorKey)
    if (paragraph === undefined) throw new Error(`fixture has no paragraph ${anchorKey}`)
    const at = contentStartOf(body.dataStream ?? '', paragraph.startIndex)
    spliceStory(body, at, at, text)
    return { part: anchor.part, span: { start: anchor.start, end: anchor.end } }
  }
  const table = (body.tables ?? []).find(candidate => anchorKey.startsWith(`${candidate.tableId}/`))
  if (table === undefined) throw new Error(`fixture has no table for ${anchorKey}`)
  const inner = (body.paragraphs ?? []).filter(
    paragraph => paragraph.startIndex > table.startIndex && paragraph.startIndex < table.endIndex,
  )
  const paragraph = inner[0]
  if (paragraph === undefined || inner.length !== 1) {
    throw new Error(`fixture cell ${anchorKey} does not hold exactly one paragraph`)
  }
  const at = contentStartOf(body.dataStream ?? '', paragraph.startIndex)
  spliceStory(body, at, at, text)
  return { part: anchor.part, span: { start: anchor.start, end: anchor.end } }
}

function insertParagraphBeforeFirstUnit(data: IDocumentData, text: string): EditEvidence {
  const body = bodyOf(data)
  const first = anchorsOf(data)
    .filter(anchor => anchor.part === DOCUMENT_PART)
    .sort((left, right) => left.start - right.start)[0]
  if (first === undefined) throw new Error('fixture has no body anchor')
  spliceStory(body, 0, 0, `${text}\r`)
  body.paragraphs = [
    ...(body.paragraphs ?? []),
    { paragraphId: 'para_inserted', startIndex: text.length },
  ].sort((left, right) => left.startIndex - right.startIndex)
  return { part: DOCUMENT_PART, insertionAt: first.start }
}

function editWorkbookCell(
  data: IWorkbookData,
  sheetId: string,
  row: number,
  column: number,
  value: string | number,
): EditEvidence {
  const sheet = data.sheets?.[sheetId]
  if (sheet === undefined) throw new Error(`workbook has no sheet ${sheetId}`)
  const cell = sheet.cellData?.[row]?.[column]
  if (cell === undefined) throw new Error(`sheet ${sheetId} has no cell ${cellReference(row, column)}`)
  cell.v = value
  const anchor = anchorByKey(data, `${sheetId}!${String(row)}:${String(column)}`)
  return { part: anchor.part, span: { start: anchor.start, end: anchor.end } }
}

function applyEdit(snapshot: UnitSnapshot, plan: EditPlan): EditEvidence {
  switch (plan.kind) {
    case 'docx-text':
      if (snapshot.kind !== 'document') throw new Error('a document edit needs a document snapshot')
      return replaceDocumentText(snapshot.data, plan.from, plan.to)
    case 'docx-text-at-unit':
      if (snapshot.kind !== 'document') throw new Error('a document edit needs a document snapshot')
      return insertTextAtUnit(snapshot.data, plan.anchorKey, plan.text)
    case 'docx-insert-paragraph':
      if (snapshot.kind !== 'document') throw new Error('a document edit needs a document snapshot')
      return insertParagraphBeforeFirstUnit(snapshot.data, plan.text)
    case 'xlsx-cell':
      if (snapshot.kind !== 'workbook') throw new Error('a workbook edit needs a workbook snapshot')
      return editWorkbookCell(snapshot.data, plan.sheetId, plan.row, plan.column, plan.value)
  }
}

function cellReference(row: number, column: number): string {
  let name = ''
  let index = column
  for (;;) {
    name = String.fromCharCode(65 + (index % 26)) + name
    index = Math.floor(index / 26) - 1
    if (index < 0) return `${name}${String(row + 1)}`
  }
}

function entryOf(entries: ReadonlyMap<string, Uint8Array>, name: string): Uint8Array {
  const bytes = entries.get(name)
  if (bytes === undefined) throw new Error(`package has no ${name}`)
  return bytes
}

function partCount(entries: ReadonlyMap<string, Uint8Array>): number {
  return [...entries.keys()].filter(name => !name.endsWith('/')).length
}

/** File-entry names the independent diff reports with one status, sorted. */
function diffStatusNames(diff: PackageDiff, status: EntryStatus): readonly string[] {
  return diff.entries
    .filter(entry => !entry.name.endsWith('/') && entry.status === status)
    .map(entry => entry.name)
    .sort()
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

function expectInsertionAt(sourcePart: Uint8Array, writtenPart: Uint8Array, at: number): void {
  expect(writtenPart.subarray(0, at)).toEqual(sourcePart.subarray(0, at))
  const tail = sourcePart.subarray(at)
  expect(writtenPart.subarray(writtenPart.byteLength - tail.byteLength)).toEqual(tail)
  expect(writtenPart.byteLength).toBeGreaterThan(sourcePart.byteLength)
}

function expectEvidenceHolds(source: Uint8Array, written: Uint8Array, evidence: EditEvidence): void {
  const sourcePart = entryOf(unzipEntries(source), evidence.part)
  const writtenPart = entryOf(unzipEntries(written), evidence.part)
  if (evidence.span !== undefined) expectOnlySpan(sourcePart, writtenPart, evidence.span)
  else if (evidence.insertionAt !== undefined) expectInsertionAt(sourcePart, writtenPart, evidence.insertionAt)
  else throw new Error('edit evidence names neither a span nor an insertion offset')
}

function expectDispositionMatchesDiff(
  source: Uint8Array,
  written: Uint8Array,
  disposition: SaveDisposition,
): PackageDiff {
  const diff = comparePackages(source, written, {
    changed: disposition.spliced.names,
    added: disposition.added.names,
    removed: disposition.removed.names,
  })
  expect(diff.ok, formatDiffReport(diff)).toBe(true)
  const changed = diff.entries.filter(entry => entry.status === 'changed').map(entry => entry.name).sort()
  const added = diff.entries.filter(entry => entry.status === 'added').map(entry => entry.name).sort()
  const removed = diff.entries.filter(entry => entry.status === 'removed').map(entry => entry.name).sort()
  expect(changed).toEqual([...disposition.spliced.names].sort())
  expect(added).toEqual([...disposition.added.names].sort())
  expect(removed).toEqual([...disposition.removed.names].sort())
  for (const entry of diff.entries) {
    if (entry.status !== 'identical') continue
    expect(entry.sourceSha256).toBe(entry.writtenSha256)
    expect(entry.firstDifferenceOffset).toBeNull()
  }
  const spliced = new Set(disposition.spliced.names)
  if (diff.relationships.added.length + diff.relationships.removed.length > 0) {
    expect([...spliced].some(name => name.endsWith('.rels'))).toBe(true)
  }
  if (diff.contentTypes.added.length + diff.contentTypes.removed.length > 0) {
    expect(spliced.has(CONTENT_TYPES_PART)).toBe(true)
  }
  return diff
}

function elementXml(xml: string, name: string): string {
  const pattern = new RegExp(`<${name}(?=[\\s/>])[^>]*(?:/>|>[\\s\\S]*?</${name}>)`, 'u')
  const match = pattern.exec(xml)
  if (match === null) throw new Error(`part has no ${name}`)
  return match[0]
}

function startTagXml(xml: string, name: string): string {
  const pattern = new RegExp(`<${name}(?=[\\s/>])[^>]*>`, 'u')
  const match = pattern.exec(xml)
  if (match === null) throw new Error(`part has no ${name} start tag`)
  return match[0]
}

async function writeSnapshot(
  fixture: FixtureRead,
  snapshot: UnitSnapshot = fixture.snapshot,
): Promise<OfficeWriteResult> {
  return ctx.officeExchange.write(snapshot, fixture.format, fixture.source)
}

/** Incompressible bytes for the opaque package entry, deterministic across runs. */
function opaqueBytes(length: number): Uint8Array<ArrayBuffer> {
  const bytes = new Uint8Array(length)
  let state = 0x2545f491
  for (let index = 0; index < length; index += 1) {
    state ^= state << 13
    state ^= state >>> 17
    state ^= state << 5
    bytes[index] = state & 0xff
  }
  return bytes
}

/**
 * Build the T058 workbook: 10,000 cells plus an opaque ~4.5 MB part.
 * @param path - Destination `.xlsx` path.
 * @returns Complete package byte length.
 */
async function writePerformanceWorkbook(path: string): Promise<number> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Data')
  for (let index = 0; index < PERF_CELLS; index += 1) {
    const row = Math.floor(index / PERF_COLUMNS) + 1
    const column = (index % PERF_COLUMNS) + 1
    sheet.getCell(row, column).value = index
  }
  const encoded = new Uint8Array(await workbook.xlsx.writeBuffer())
  const entries = unzipSync(encoded)
  const contentTypes = entries[CONTENT_TYPES_PART]
  if (contentTypes === undefined) throw new Error('performance workbook has no content-types part')
  const original = strFromU8(contentTypes)
  const declared = original.replace(
    '</Types>',
    '<Default Extension="bin" ContentType="application/octet-stream"/></Types>',
  )
  if (declared === original) throw new Error('performance workbook content types did not accept the opaque part')
  entries[CONTENT_TYPES_PART] = strToU8(declared)
  entries['xl/media/perf.bin'] = opaqueBytes(PERF_OPAQUE_BYTES)
  const packaged = zipSync(entries)
  await writeFile(path, packaged)
  return packaged.byteLength
}

describe('OOXML preservation corpus', () => {
  it('enumerates at least thirty OOXML fixtures excluding the known refusal cases', () => {
    expect(ALL_OOXML_FIXTURES.length).toBeGreaterThanOrEqual(30)
    expect(CORPUS.length).toBeGreaterThanOrEqual(30)
    for (const name of KNOWN_REFUSAL_FIXTURES) {
      expect(ALL_OOXML_FIXTURES).toContain(name)
      expect(CORPUS).not.toContain(name)
    }
  })

  it('documents at least one fixture per FR-004 construct class', () => {
    for (const [label, fixtures] of Object.entries(FR004_CLASSES)) {
      const present = fixtures.filter(fixture => ALL_OOXML_FIXTURES.includes(fixture))
      expect(present, `FR-004 class ${label} has no fixture`).not.toHaveLength(0)
    }
  })
})

describe('no-op save identity', () => {
  for (const fixtureName of CORPUS) {
    it(`reproduces ${fixtureName} byte-for-byte`, async () => {
      const fixture = await readFixture(fixtureName)
      const result = await writeSnapshot(fixture)
      expect(result.bytes).toEqual(fixture.bytes)
      expect(result.disposition.spliced).toEqual({ count: 0, names: [] })
      expect(result.disposition.added).toEqual({ count: 0, names: [] })
      expect(result.disposition.removed).toEqual({ count: 0, names: [] })
      expect(result.disposition.verbatimCount).toBe(partCount(unzipEntries(fixture.bytes)))
      expect(validateOpc(unzipEntries(result.bytes))).toEqual([])
      const repeated = await writeSnapshot(fixture)
      expect(repeated.bytes).toEqual(result.bytes)
      expect(repeated.disposition).toEqual(result.disposition)
    })
  }
})

describe('targeted edit attribution', () => {
  for (const fixtureName of CORPUS) {
    it(`splices only the edited unit in ${fixtureName}`, async () => {
      const plan = EDITS[fixtureName]
      if (plan === undefined) throw new Error(`corpus fixture ${fixtureName} has no edit plan`)
      const fixture = await readFixture(fixtureName)
      const evidence = applyEdit(fixture.snapshot, plan)
      const result = await writeSnapshot(fixture)
      expect(result.bytes).not.toEqual(fixture.bytes)
      expectDispositionMatchesDiff(fixture.bytes, result.bytes, result.disposition)
      expectEvidenceHolds(fixture.bytes, result.bytes, evidence)
      expect(validateOpc(unzipEntries(result.bytes))).toEqual([])
      const repeated = await writeSnapshot(fixture)
      expect(repeated.bytes).toEqual(result.bytes)
      expect(repeated.disposition).toEqual(result.disposition)
    })
  }
})

describe('save disposition agrees with the independent diff', () => {
  for (const fixtureName of DISPOSITION_FIXTURES) {
    it(`accounts for every entry in ${fixtureName}`, async () => {
      const plan = EDITS[fixtureName]
      if (plan === undefined) throw new Error(`fixture ${fixtureName} has no edit plan`)
      const fixture = await readFixture(fixtureName)
      applyEdit(fixture.snapshot, plan)
      const result = await writeSnapshot(fixture)
      expect(result.bytes).not.toEqual(fixture.bytes)
      const sourceEntries = unzipEntries(fixture.bytes)
      const writtenEntries = unzipEntries(result.bytes)
      const diff = comparePackages(fixture.bytes, result.bytes, {
        changed: result.disposition.spliced.names,
        added: result.disposition.added.names,
        removed: result.disposition.removed.names,
      })
      expect(diff.ok, formatDiffReport(diff)).toBe(true)

      const identical = diffStatusNames(diff, 'identical')
      const changed = diffStatusNames(diff, 'changed')
      const added = diffStatusNames(diff, 'added')
      const removed = diffStatusNames(diff, 'removed')

      expect(changed).toEqual([...result.disposition.spliced.names].sort())
      expect(added).toEqual([...result.disposition.added.names].sort())
      expect(removed).toEqual([...result.disposition.removed.names].sort())
      expect(result.disposition.verbatimCount).toBe(identical.length)
      expect(result.disposition.spliced.count).toBe(changed.length)
      expect(result.disposition.added.count).toBe(added.length)
      expect(result.disposition.removed.count).toBe(removed.length)

      expect(result.disposition.spliced.names).toHaveLength(result.disposition.spliced.count)
      expect(result.disposition.added.names).toHaveLength(result.disposition.added.count)
      expect(result.disposition.removed.names).toHaveLength(result.disposition.removed.count)

      expect(identical.length + changed.length + added.length + removed.length)
        .toBe(diff.entries.filter(entry => !entry.name.endsWith('/')).length)
      expect(identical.length + changed.length + added.length).toBe(partCount(writtenEntries))
      expect(identical.length + changed.length + removed.length).toBe(partCount(sourceEntries))
      expect(validateOpc(writtenEntries)).toEqual([])
    })
  }
})

describe('preserved constructs survive a non-intersecting edit', () => {
  for (const testCase of CONSTRUCT_CASES) {
    it(`preserves ${testCase.label} in ${testCase.fixture}`, async () => {
      const plan = EDITS[testCase.fixture]
      if (plan === undefined) throw new Error(`fixture ${testCase.fixture} has no edit plan`)
      const fixture = await readFixture(testCase.fixture)
      applyEdit(fixture.snapshot, plan)
      const result = await writeSnapshot(fixture)
      expect(result.bytes).not.toEqual(fixture.bytes)
      expect(result.disposition.added).toEqual({ count: 0, names: [] })
      expect(result.disposition.removed).toEqual({ count: 0, names: [] })
      expect(validateOpc(unzipEntries(result.bytes))).toEqual([])

      const sourceEntries = unzipEntries(fixture.bytes)
      const writtenEntries = unzipEntries(result.bytes)
      for (const name of testCase.identicalEntries) {
        expect(entryOf(writtenEntries, name), `${name} changed`).toEqual(entryOf(sourceEntries, name))
      }
      for (const fragment of testCase.fragments ?? []) {
        const sourceXml = strFromU8(entryOf(sourceEntries, fragment.part))
        const writtenXml = strFromU8(entryOf(writtenEntries, fragment.part))
        const xml = fragment.mode === 'openTag'
          ? startTagXml(sourceXml, fragment.element)
          : elementXml(sourceXml, fragment.element)
        expect(writtenXml, `${fragment.element} changed`).toContain(xml)
      }
    })
  }
})

describe('macro-enabled packages stay macro-enabled', () => {
  for (const fixtureName of ['macro.docm', 'macro.xlsm']) {
    it(`keeps the VBA project and macro-enabled content type in ${fixtureName}`, async () => {
      const plan = EDITS[fixtureName]
      if (plan === undefined) throw new Error(`fixture ${fixtureName} has no edit plan`)
      const fixture = await readFixture(fixtureName)
      applyEdit(fixture.snapshot, plan)
      const result = await writeSnapshot(fixture)
      const sourceEntries = unzipEntries(fixture.bytes)
      const writtenEntries = unzipEntries(result.bytes)
      const macroEntry = fixtureName.endsWith('.docm') ? 'word/vbaProject.bin' : 'xl/vbaProject.bin'
      expect(entryOf(writtenEntries, macroEntry)).toEqual(entryOf(sourceEntries, macroEntry))

      const sourceTypes = strFromU8(entryOf(sourceEntries, CONTENT_TYPES_PART))
      const writtenTypes = strFromU8(entryOf(writtenEntries, CONTENT_TYPES_PART))
      expect(writtenTypes).toEqual(sourceTypes)
      expect(writtenTypes).toContain(MACRO_CONTENT_TYPES[fixtureName] ?? '')
      const macroFree = fixtureName.endsWith('.docm')
        ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
      expect(writtenTypes).not.toContain(macroFree)
    })
  }
})

describe('large-workbook save performance', () => {
  it('saves a 5 MB, 10,000-cell workbook in under 30 seconds', { timeout: PERF_TIMEOUT_MS }, async () => {
    const directory = await mkdtemp(join(tmpdir(), 'dsh-office-exchange-perf-'))
    try {
      const path = join(directory, 'perf.xlsx')
      const packageBytes = await writePerformanceWorkbook(path)
      expect(packageBytes).toBeGreaterThanOrEqual(PERF_OPAQUE_BYTES)
      const read = await ctx.officeExchange.read(path)
      expect(read.format).toBe('xlsx')
      if (read.snapshot.kind !== 'workbook') throw new Error('performance fixture did not read as a workbook')
      editWorkbookCell(read.snapshot.data, 'sheet-1', 1, 1, 424_242)
      const source = open(new Uint8Array(readFileSync(path)), { maxEntries: 10_000 }).source

      const startedAt = performance.now()
      const result = await ctx.officeExchange.write(read.snapshot, 'xlsx', source)
      const elapsedMs = performance.now() - startedAt

      console.info(
        `large-workbook save: ${elapsedMs.toFixed(0)} ms, ${String(packageBytes)} bytes `
        + `(${(packageBytes / 1024 / 1024).toFixed(2)} MiB), ${String(PERF_CELLS)} cells`,
      )
      expect(validateOpc(unzipEntries(result.bytes))).toEqual([])
      expect(elapsedMs).toBeLessThan(PERF_SAVE_BUDGET_MS)
    } finally {
      await rm(directory, { recursive: true, force: true })
    }
  })
})
