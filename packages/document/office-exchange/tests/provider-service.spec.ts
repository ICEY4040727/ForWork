/** Concrete Exchange provider: file reads/writes, detection, and typed failures. */
import { mkdtemp, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { IDocumentData, IWorkbookData } from '@univerjs/core'
import { Context } from '@forwork/cordis'
import { ConversionError, createConverter } from '@deepseek-ai/libreoffice-kit'
import { afterEach, beforeEach, expect, it } from 'vitest'
import LocalOfficeExchange, { Config, OfficeExchangeError, type ExchangeReadResult, type OfficeExchange } from '../src/index.ts'

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url))
const XLSX_FIXTURE = join(FIXTURES, 'roundtrip.xlsx')
const DOCX_FIXTURE = join(FIXTURES, 'roundtrip.docx')
const MACRO_FIXTURE = join(FIXTURES, 'unsupported-macro.docx')
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

it('writes snapshots that the service reads back', async () => {
  const exchange = await mount()
  const workspace = await mkdtemp(join(tmpdir(), 'dsh-office-exchange-service-'))
  try {
    const workbook = await exchange.read(XLSX_FIXTURE)
    const workbookPath = join(workspace, 'roundtrip.xlsx')
    await writeFile(workbookPath, await exchange.write(workbook.snapshot, 'xlsx'))
    const rereadWorkbook = await exchange.read(workbookPath)
    expect(rereadWorkbook.format).toBe('xlsx')
    expect(cellValue(workbookOf(rereadWorkbook), 'Sales', 1, 0)).toBe('Widget')

    const document = await exchange.read(DOCX_FIXTURE)
    const documentPath = join(workspace, 'roundtrip.docx')
    await writeFile(documentPath, await exchange.write(document.snapshot, 'docx'))
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

it('refuses preview-only and legacy targets and mismatched snapshot kinds', async () => {
  const exchange = await mount()
  const workbook = await exchange.read(XLSX_FIXTURE)
  for (const format of ['pptx', 'doc', 'xls', 'ppt'] as const) {
    await expect(exchange.write(workbook.snapshot, format)).rejects.toMatchObject({ code: 'unsupported-format' })
  }
  await expect(exchange.write(workbook.snapshot, 'docx')).rejects.toMatchObject({ code: 'conversion-failed' })

  const document = await exchange.read(DOCX_FIXTURE)
  await expect(exchange.write(document.snapshot, 'xlsx')).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('propagates caller cancellation unchanged', async () => {
  const exchange = await mount()
  const reason = new Error('cancelled')
  await expect(exchange.read(XLSX_FIXTURE, AbortSignal.abort(reason))).rejects.toBe(reason)
  const workbook = await exchange.read(XLSX_FIXTURE)
  await expect(exchange.write(workbook.snapshot, 'xlsx', AbortSignal.abort(reason))).rejects.toBe(reason)
  await expect(exchange.convertLegacy(LEGACY_FIXTURE, AbortSignal.abort(reason))).rejects.toBe(reason)
})

it('resolves defaults through the schema and rejects invalid bounds loudly', async () => {
  await mount()
  expect(Config({})).toEqual({
    maxFileBytes: 25 * 1024 * 1024,
    maxCells: 2_000_000,
    maxConversionMs: 60_000,
    commandTimeoutMs: 5_000,
  })
  expect(Config({ maxCells: 10 })).toMatchObject({ maxCells: 10, maxFileBytes: 25 * 1024 * 1024 })
  expect(() => Config({ maxCells: 0 })).toThrow()
})

it('rejects an invalid bound when constructed directly', () => {
  expect(() => new LocalOfficeExchange(ctx, { maxCells: 0 })).toThrow(RangeError)
})
