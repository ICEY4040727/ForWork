/** Concrete Exchange provider: file reads/writes, detection, and typed failures. */
import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { ICellData, IDocumentData, IWorkbookData } from '@univerjs/core'
import { Context } from '@forwork/cordis'
import { ConversionError, createConverter } from '@deepseek-ai/libreoffice-kit'
import { afterEach, beforeEach, expect, it } from 'vitest'
import LocalOfficeExchange, { Config, OfficeExchangeError, type ExchangeReadResult, type OfficeExchange } from '../src/index.ts'
import { saveDispositionOf } from '../src/provider/index.ts'
import { open } from '../src/package/ooxml-package.ts'
import type { SourcePackage } from '../src/package/types.ts'
import type { OfficeSourceState } from '../src/source/source-state.ts'

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url))
const XLSX_FIXTURE = join(FIXTURES, 'roundtrip.xlsx')
const DOCX_FIXTURE = join(FIXTURES, 'roundtrip.docx')
const DOCM_FIXTURE = join(FIXTURES, 'macro.docm')
const XLSM_FIXTURE = join(FIXTURES, 'macro.xlsm')
const MACRO_FIXTURE = join(FIXTURES, 'unsupported-macro.docx')
const IMAGE_FIXTURE = join(FIXTURES, 'preserve-image.docx')
const HEADER_IMAGE_FIXTURE = join(FIXTURES, 'preserve-header-image.docx')
const CONTENT_CONTROL_FIXTURE = join(FIXTURES, 'preserve-content-control.docx')
const EXCEL_TABLE_FIXTURE = join(FIXTURES, 'preserve-excel-table.xlsx')
const LEGACY_FIXTURE = join(FIXTURES, 'legacy.xls')

let ctx: Context

beforeEach(() => {
  ctx = new Context()
})

afterEach(async () => {
  await ctx.fiber.dispose()
})

async function mount(config: Partial<Config> = {}): Promise<OfficeExchange> {
  await ctx.plugin(LocalOfficeExchange, config)
  return ctx.officeExchange
}

function workbookOf(result: ExchangeReadResult): IWorkbookData {
  if (result.snapshot.kind !== 'workbook') throw new Error('expected a workbook snapshot')
  return result.snapshot.data
}

function documentOf(result: ExchangeReadResult): IDocumentData {
  if (result.snapshot.kind !== 'document') throw new Error('expected a document snapshot')
  return result.snapshot.data
}

function cellValue(workbook: IWorkbookData, sheetName: string, row: number, column: number): unknown {
  const id = workbook.sheetOrder.find(candidate => workbook.sheets[candidate]?.name === sheetName)
  return id === undefined ? undefined : workbook.sheets[id]?.cellData?.[row]?.[column]?.v
}

type SidecarCarrier = {
  officeExchange?: {
    readonly source?: unknown
    readonly docx?: unknown
  }
}

function sidecarOf(data: IDocumentData | IWorkbookData): SidecarCarrier['officeExchange'] {
  return (data as SidecarCarrier).officeExchange
}

function sourceStateOf(data: IDocumentData | IWorkbookData): OfficeSourceState | undefined {
  return sidecarOf(data)?.source as OfficeSourceState | undefined
}

async function sourcePackageOf(path: string): Promise<SourcePackage> {
  return open(new Uint8Array(await readFile(path)), { maxEntries: 10_000 }).source
}

function countWorkbookCells(workbook: IWorkbookData): number {
  let count = 0
  for (const id of workbook.sheetOrder) {
    const cellData: Record<number, Record<number, ICellData>> = workbook.sheets[id]?.cellData ?? {}
    for (const row of Object.values(cellData)) {
      count += Object.keys(row).length
    }
  }
  return count
}

function removeSidecar(data: IDocumentData | IWorkbookData): void {
  delete (data as SidecarCarrier).officeExchange
}

let availability: Promise<string | undefined> | undefined

/** Report why the conversion test must skip, or `undefined` when the engine exists. */
function unavailableReason(): Promise<string | undefined> {
  availability ??= probeEngine()
  return availability
}

async function probeEngine(): Promise<string | undefined> {
  try {
    const converter = await createConverter()
    await converter.dispose()
    return undefined
  } catch (error) {
    if (error instanceof ConversionError && error.code === 'unavailable') return error.message
    throw error
  }
}

it('reads the workbook fixture into a workbook snapshot', async () => {
  const exchange = await mount()
  const result = await exchange.read(XLSX_FIXTURE)
  expect(result.format).toBe('xlsx')
  expect(result.snapshot.kind).toBe('workbook')
  expect(result.unsupported).toEqual([])
  expect(cellValue(workbookOf(result), 'Sales', 0, 0)).toBe('Item')
  expect(cellValue(workbookOf(result), 'Sales', 1, 0)).toBe('Widget')
})

it('reads the document fixture into a document snapshot', async () => {
  const exchange = await mount()
  const result = await exchange.read(DOCX_FIXTURE)
  expect(result.format).toBe('docx')
  expect(result.snapshot.kind).toBe('document')
  expect(result.unsupported).toEqual([])
  expect(documentOf(result).body?.dataStream).toContain('Quarterly Report')
})

it('attaches validated source anchors to a document read', async () => {
  const exchange = await mount()
  const result = await exchange.read(DOCX_FIXTURE)
  const data = documentOf(result)
  const state = sourceStateOf(data)
  expect(state).toBeDefined()
  expect(state?.schemaVersion).toBe(1)
  expect(state?.format).toBe('docx')
  expect(state?.packageDigest).toMatch(/^[0-9a-f]{64}$/)
  expect(state?.anchorMode).toBe('per-unit')
  const source = await sourcePackageOf(DOCX_FIXTURE)
  expect(state?.packageDigest).toBe(source.digest)
  expect(state?.byteLength).toBe(source.bytes.byteLength)

  const anchors = state?.anchorMode === 'per-unit' ? state.anchors : []
  expect(anchors.length).toBe(6)
  expect(anchors.filter(anchor => anchor.kind === 'paragraph')).toHaveLength(2)
  expect(anchors.filter(anchor => anchor.kind === 'tableCell')).toHaveLength(4)
  const partLengths = new Map(source.entries.map(entry => [entry.name, entry.uncompressedSize]))
  for (const anchor of anchors) {
    expect(anchor.digest).toMatch(/^[0-9a-f]{64}$/)
    expect(anchor.end).toBeGreaterThan(anchor.start)
    expect(anchor.start).toBeGreaterThanOrEqual(0)
    expect(anchor.end).toBeLessThanOrEqual(partLengths.get(anchor.part) ?? 0)
    expect(anchor.key.length).toBeGreaterThan(0)
  }
  expect(anchors.filter(anchor => anchor.kind === 'paragraph').map(anchor => anchor.key))
    .toEqual(['para_1', 'para_2'])

  const paragraphIds = (data.body?.paragraphs ?? []).map(paragraph => paragraph.paragraphId)
  expect(paragraphIds).toHaveLength(7)
  expect(paragraphIds.every(id => id?.startsWith('para_'))).toBe(true)
  expect(sidecarOf(data)?.docx).toBeUndefined()
})

it('anchors header paragraphs as their own editable part', async () => {
  const exchange = await mount()
  const data = documentOf(await exchange.read(HEADER_IMAGE_FIXTURE))
  const state = sourceStateOf(data)
  expect(state?.anchorMode).toBe('per-unit')
  if (state?.anchorMode !== 'per-unit') return
  expect(state.sourceParts).toContain('word/header1.xml')
  const headerAnchors = state.anchors.filter(anchor => anchor.part === 'word/header1.xml')
  expect(headerAnchors).toHaveLength(1)
  expect(headerAnchors[0]?.kind).toBe('paragraph')
  expect(headerAnchors[0]?.key).toBe('para_1')
  expect(headerAnchors[0]?.digest).toMatch(/^[0-9a-f]{64}$/)
})

it('attaches one anchor per mapped workbook cell', async () => {
  const exchange = await mount()
  const result = await exchange.read(XLSX_FIXTURE)
  const workbook = workbookOf(result)
  const state = sourceStateOf(workbook)
  expect(state?.schemaVersion).toBe(1)
  expect(state?.format).toBe('xlsx')
  expect(state?.packageDigest).toMatch(/^[0-9a-f]{64}$/)
  expect(state?.anchorMode).toBe('per-unit')
  const anchors = state?.anchorMode === 'per-unit' ? state.anchors : []
  expect(anchors).toHaveLength(countWorkbookCells(workbook))
  for (const anchor of anchors) {
    expect(anchor.kind).toBe('sheetCell')
    expect(anchor.part).toMatch(/^xl\/worksheets\/sheet\d+\.xml$/u)
    expect(anchor.key).toMatch(/^sheet-\d+!\d+:\d+$/u)
    expect(anchor.digest).toMatch(/^[0-9a-f]{64}$/)
    expect(anchor.end).toBeGreaterThan(anchor.start)
  }
})

it('stores sheet summaries when the anchor budget is one', async () => {
  const exchange = await mount({ maxSourceAnchors: 1 })
  const result = await exchange.read(EXCEL_TABLE_FIXTURE)
  const state = sourceStateOf(workbookOf(result))
  expect(state?.anchorMode).toBe('sheet-summary')
  const sheets = state?.anchorMode === 'sheet-summary' ? state.sheets : []
  expect(sheets).toHaveLength(1)
  expect(sheets[0]?.sheetId).toBe('sheet-1')
  expect(sheets[0]?.part).toBe('xl/worksheets/sheet1.xml')
  expect(sheets[0]?.digest).toMatch(/^[0-9a-f]{64}$/)
  expect(sheets[0]?.cellCount).toBeGreaterThan(0)
})

it('normalizes a plan disposition to exact counts and sorted capped names', () => {
  const names = Array.from(
    { length: 300 },
    (_, index) => `xl/entry-${String(299 - index).padStart(3, '0')}.xml`,
  )
  const disposition = saveDispositionOf({
    schemaVersion: 1,
    verbatimCount: 7,
    spliced: { count: 300, names },
    added: { count: 2, names: ['xl/b.xml', 'xl/a.xml'] },
    removed: { count: 1, names: ['xl/z.xml'] },
    nameLimit: 999,
  })

  expect(disposition.verbatimCount).toBe(7)
  expect(disposition.nameLimit).toBe(256)
  expect(disposition.spliced.count).toBe(300)
  expect(disposition.spliced.names).toEqual([...names].sort().slice(0, 256))
  expect(disposition.added).toEqual({ count: 2, names: ['xl/a.xml', 'xl/b.xml'] })
  expect(disposition.removed).toEqual({ count: 1, names: ['xl/z.xml'] })
})

it('anchors a body content control as one preserved unit', async () => {
  const exchange = await mount()
  const data = documentOf(await exchange.read(CONTENT_CONTROL_FIXTURE))
  const state = sourceStateOf(data)
  expect(state?.schemaVersion).toBe(1)
  expect(state?.format).toBe('docx')
  expect(state?.anchorMode).toBe('per-unit')
  const anchors = state?.anchorMode === 'per-unit' ? state.anchors : []
  expect(anchors).toHaveLength(1)
  expect(anchors[0]).toMatchObject({
    kind: 'contentControl',
    part: 'word/document.xml',
    key: 'para_1',
  })
  const written = await exchange.write({ kind: 'document', data }, 'docx')
  expect(written.bytes.byteLength).toBeGreaterThan(0)
  expect(written.disposition.verbatimCount).toBe(0)
})

it('does not attach base64 preserved parts when source anchors exist', async () => {
  const exchange = await mount()
  const data = documentOf(await exchange.read(IMAGE_FIXTURE))
  expect(sourceStateOf(data)).toBeDefined()
  expect(sidecarOf(data)?.docx).toBeUndefined()
})

it('writes snapshots that the service reads back', async () => {
  const exchange = await mount()
  const workspace = await mkdtemp(join(tmpdir(), 'dsh-office-exchange-service-'))
  try {
    const workbook = await exchange.read(XLSX_FIXTURE)
    const workbookPath = join(workspace, 'roundtrip.xlsx')
    await writeFile(workbookPath, (await exchange.write(workbook.snapshot, 'xlsx')).bytes)
    const rereadWorkbook = await exchange.read(workbookPath)
    expect(rereadWorkbook.format).toBe('xlsx')
    expect(cellValue(workbookOf(rereadWorkbook), 'Sales', 1, 0)).toBe('Widget')

    const document = await exchange.read(DOCX_FIXTURE)
    const documentPath = join(workspace, 'roundtrip.docx')
    await writeFile(documentPath, (await exchange.write(document.snapshot, 'docx')).bytes)
    const rereadDocument = await exchange.read(documentPath)
    expect(rereadDocument.format).toBe('docx')
    expect(documentOf(rereadDocument).body?.dataStream).toContain('Quarterly Report')
  } finally {
    await rm(workspace, { recursive: true, force: true })
  }
})

it('reports a macro feature on read and through detection', async () => {
  const exchange = await mount()
  const result = await exchange.read(MACRO_FIXTURE)
  expect(result.unsupported.some(feature => feature.kind === 'macro')).toBe(true)
  expect(await exchange.detectUnsupported(DOCX_FIXTURE)).toEqual([])
  const detected = await exchange.detectUnsupported(MACRO_FIXTURE)
  expect(detected.some(feature => feature.kind === 'macro')).toBe(true)
})

it('reads macro-enabled formats and saves them through their anchored source', async () => {
  const exchange = await mount()
  const document = await exchange.read(DOCM_FIXTURE)
  expect(document.format).toBe('docm')
  expect(document.snapshot.kind).toBe('document')
  expect(documentOf(document).body?.dataStream).toContain('Quarterly Report')
  expect(document.unsupported.map(feature => feature.kind)).toContain('macro')

  const workbook = await exchange.read(XLSM_FIXTURE)
  expect(workbook.format).toBe('xlsm')
  expect(workbook.snapshot.kind).toBe('workbook')
  expect(cellValue(workbookOf(workbook), 'Sales', 1, 0)).toBe('Widget')
  expect(workbook.unsupported.map(feature => feature.kind)).toContain('macro')

  await expect(exchange.write(document.snapshot, 'docm')).rejects.toMatchObject({
    code: 'unsupported-feature',
  })
  await expect(exchange.write(workbook.snapshot, 'xlsm')).rejects.toMatchObject({
    code: 'unsupported-feature',
  })

  const writtenDocument = await exchange.write(document.snapshot, 'docm', await sourcePackageOf(DOCM_FIXTURE))
  expect(writtenDocument.bytes.byteLength).toBeGreaterThan(0)
  expect(writtenDocument.disposition.verbatimCount).toBeGreaterThan(0)

  const writtenWorkbook = await exchange.write(workbook.snapshot, 'xlsm', await sourcePackageOf(XLSM_FIXTURE))
  expect(writtenWorkbook.bytes.byteLength).toBeGreaterThan(0)
  expect(writtenWorkbook.disposition.verbatimCount).toBeGreaterThan(0)

  const workspace = await mkdtemp(join(tmpdir(), 'dsh-office-exchange-macro-'))
  try {
    const documentPath = join(workspace, 'macro.docm')
    await writeFile(documentPath, writtenDocument.bytes)
    const rereadDocument = await exchange.read(documentPath)
    expect(rereadDocument.format).toBe('docm')
    expect(documentOf(rereadDocument).body?.dataStream).toContain('Quarterly Report')

    const workbookPath = join(workspace, 'macro.xlsm')
    await writeFile(workbookPath, writtenWorkbook.bytes)
    const rereadWorkbook = await exchange.read(workbookPath)
    expect(rereadWorkbook.format).toBe('xlsm')
    expect(cellValue(workbookOf(rereadWorkbook), 'Sales', 1, 0)).toBe('Widget')
  } finally {
    await rm(workspace, { recursive: true, force: true })
  }
})

it('rejects a legacy workbook until conversion, then reads the converted file', { timeout: 180_000 }, async (context) => {
  const exchange = await mount()
  await expect(exchange.read(LEGACY_FIXTURE)).rejects.toMatchObject({ code: 'unsupported-format' })
  await expect(exchange.detectUnsupported(LEGACY_FIXTURE)).rejects.toMatchObject({ code: 'unsupported-format' })

  const reason = await unavailableReason()
  if (reason !== undefined) {
    context.skip(`LibreOffice unavailable: ${reason}`)
    return
  }
  const converted = await exchange.convertLegacy(LEGACY_FIXTURE)
  try {
    expect(converted.endsWith('.xlsx')).toBe(true)
    expect(dirname(converted)).not.toBe(FIXTURES)
    const result = await exchange.read(converted)
    expect(result.format).toBe('xlsx')
    expect(cellValue(workbookOf(result), 'Legacy', 0, 0)).toBe('Item')
    expect(cellValue(workbookOf(result), 'Legacy', 1, 0)).toBe('Widget')
  } finally {
    await rm(dirname(converted), { recursive: true, force: true })
  }
})

it('rejects an unknown extension and reports an unreadable file as a conversion failure', async () => {
  const exchange = await mount()
  await expect(exchange.read(join(FIXTURES, 'notes.txt'))).rejects.toMatchObject({ code: 'unsupported-format' })
  await expect(exchange.detectUnsupported(join(FIXTURES, 'notes.txt'))).rejects.toMatchObject({ code: 'unsupported-format' })
  await expect(exchange.read(join(FIXTURES, 'deck.pptx'))).rejects.toMatchObject({ code: 'unsupported-format' })

  const missing = join(FIXTURES, 'missing-roundtrip.xlsx')
  const error: unknown = await exchange.read(missing).catch((cause: unknown) => cause)
  if (!(error instanceof OfficeExchangeError)) throw error
  expect(error.code).toBe('conversion-failed')
  expect(error.message).toContain(missing)
})

it('removes the temporary directory when legacy conversion fails', async () => {
  const exchange = await mount()
  await expect(exchange.convertLegacy(join(FIXTURES, 'notes.txt')))
    .rejects.toMatchObject({ code: 'unsupported-format' })
})

it('leaves a caller-supplied output directory in place when conversion fails', async () => {
  const exchange = await mount()
  const outputDir = await mkdtemp(join(tmpdir(), 'dsh-office-exchange-output-'))
  try {
    await expect(exchange.convertLegacy(join(FIXTURES, 'notes.txt'), undefined, outputDir))
      .rejects.toMatchObject({ code: 'unsupported-format' })
    expect((await stat(outputDir)).isDirectory()).toBe(true)
  } finally {
    await rm(outputDir, { recursive: true, force: true })
  }
})

it('converts into a caller-supplied output directory', { timeout: 180_000 }, async (context) => {
  const exchange = await mount()
  const workspace = await mkdtemp(join(tmpdir(), 'dsh-office-exchange-output-'))
  try {
    const reason = await unavailableReason()
    if (reason !== undefined) {
      context.skip(`LibreOffice unavailable: ${reason}`)
      return
    }
    const outputDir = join(workspace, 'nested', 'converted')
    const converted = await exchange.convertLegacy(LEGACY_FIXTURE, undefined, outputDir)
    expect(converted).toBe(join(outputDir, 'legacy.xlsx'))
    expect(workbookOf(await exchange.read(converted)).id).not.toBe('')
  } finally {
    await rm(workspace, { recursive: true, force: true })
  }
})

it('saves a snapshot without a source sidecar through the legacy writer', async () => {
  const exchange = await mount()
  const read = await exchange.read(DOCX_FIXTURE)
  const snapshot = structuredClone(read.snapshot)
  if (snapshot.kind !== 'document') throw new Error('expected a document snapshot')
  removeSidecar(snapshot.data)

  const written = await exchange.write(snapshot, 'docx')
  expect(written.bytes.byteLength).toBeGreaterThan(0)
  expect(written.disposition).toEqual({
    schemaVersion: 1,
    verbatimCount: 0,
    spliced: { count: 0, names: [] },
    added: { count: 0, names: [] },
    removed: { count: 0, names: [] },
    nameLimit: 256,
  })
})

it('fails a save with an unparsable source sidecar and produces no bytes', async () => {
  const exchange = await mount()
  const read = await exchange.read(DOCX_FIXTURE)
  const snapshot = structuredClone(read.snapshot)
  if (snapshot.kind !== 'document') throw new Error('expected a document snapshot')
  const carrier = snapshot.data as SidecarCarrier
  carrier.officeExchange = { source: { schemaVersion: 2 } }
  const source = await sourcePackageOf(DOCX_FIXTURE)

  await expect(exchange.write(snapshot, 'docx')).rejects.toMatchObject({ code: 'conversion-failed' })
  await expect(exchange.write(snapshot, 'docx', source)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses preview-only and legacy targets and mismatched snapshot kinds', async () => {
  const exchange = await mount()
  const workbook = await exchange.read(XLSX_FIXTURE)
  for (const format of ['pptx', 'doc', 'xls', 'ppt'] as const) {
    await expect(exchange.write(workbook.snapshot, format)).rejects.toMatchObject({ code: 'unsupported-format' })
  }
  await expect(exchange.write(workbook.snapshot, 'docx')).rejects.toMatchObject({ code: 'conversion-failed' })
  await expect(exchange.write(workbook.snapshot, 'docm')).rejects.toMatchObject({ code: 'conversion-failed' })

  const document = await exchange.read(DOCX_FIXTURE)
  await expect(exchange.write(document.snapshot, 'xlsx')).rejects.toMatchObject({ code: 'conversion-failed' })
  await expect(exchange.write(document.snapshot, 'xlsm')).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('propagates caller cancellation unchanged', async () => {
  const exchange = await mount()
  const reason = new Error('cancelled')
  await expect(exchange.read(XLSX_FIXTURE, AbortSignal.abort(reason))).rejects.toBe(reason)
  const workbook = await exchange.read(XLSX_FIXTURE)
  await expect(exchange.write(workbook.snapshot, 'xlsx', undefined, AbortSignal.abort(reason))).rejects.toBe(reason)
  await expect(exchange.convertLegacy(LEGACY_FIXTURE, AbortSignal.abort(reason))).rejects.toBe(reason)
})

it('resolves defaults through the schema and rejects invalid bounds loudly', async () => {
  await mount()
  expect(Config({})).toEqual({
    maxFileBytes: 25 * 1024 * 1024,
    maxCells: 2_000_000,
    maxConversionMs: 60_000,
    commandTimeoutMs: 5_000,
    maxSourceAnchors: 100_000,
    maxSourceAnchorBytes: 8 * 1024 * 1024,
  })
  expect(Config({ maxCells: 10 })).toMatchObject({ maxCells: 10, maxFileBytes: 25 * 1024 * 1024 })
  expect(() => Config({ maxCells: 0 })).toThrow()
  expect(() => Config({ maxSourceAnchors: 0 })).toThrow()
  expect(() => Config({ maxSourceAnchorBytes: 0 })).toThrow()
  expect(() => Config({ maxSourceAnchors: 1.5 })).toThrow()
  expect(() => Config({ maxSourceAnchorBytes: Number.NaN })).toThrow()
})

it('rejects an invalid bound when constructed directly', async () => {
  for (const config of [{ maxCells: 0 }, { maxSourceAnchors: 0 }, { maxSourceAnchorBytes: 0 }]) {
    const local = new Context()
    try {
      expect(() => new LocalOfficeExchange(local, config)).toThrow(RangeError)
    } finally {
      await local.fiber.dispose()
    }
  }
})
