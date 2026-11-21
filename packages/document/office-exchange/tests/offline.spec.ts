/**
 * FR-010: editable Office files are read, edited, and saved with no document
 * server, network transport, or cloud account. Each round-trip runs through the
 * local provider with empty configuration while every outbound transport is
 * replaced by a refusal, so a hidden remote dependency fails the operation.
 */
import childProcess from 'node:child_process'
import dns from 'node:dns'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import http from 'node:http'
import https from 'node:https'
import net from 'node:net'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { ICellData, IDocumentData, IWorksheetData, IWorkbookData } from '@univerjs/core'
import { BooleanNumber } from '@univerjs/core'
import { Context } from '@forwork/cordis'
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import LocalOfficeExchange, {
  Config,
  type ExchangeReadResult,
  type OfficeExchange,
  type SourcePackage,
} from '../src/index.ts'
import { open } from '../src/package/ooxml-package.ts'

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url))
const XLSX_FIXTURE = join(FIXTURES, 'roundtrip.xlsx')
const DOCX_FIXTURE = join(FIXTURES, 'roundtrip.docx')

let ctx: Context

beforeEach(() => {
  ctx = new Context()
})

afterEach(async () => {
  vi.restoreAllMocks()
  await ctx.fiber.dispose()
})

async function mount(): Promise<OfficeExchange> {
  await ctx.plugin(LocalOfficeExchange, {})
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

function sheetOf(workbook: IWorkbookData, name: string): Partial<IWorksheetData> {
  const id = workbook.sheetOrder.find(candidate => workbook.sheets[candidate]?.name === name)
  if (id === undefined) throw new Error(`missing sheet ${name}`)
  return workbook.sheets[id] ?? {}
}

function cellAt(sheet: Partial<IWorksheetData>, row: number, column: number): ICellData | undefined {
  return sheet.cellData?.[row]?.[column]
}

/** Compare editable content without the source sidecar, whose digest tracks the file. */
function withoutSidecar<T extends IDocumentData | IWorkbookData>(data: T): T {
  const clone = structuredClone(data)
  delete (clone as { officeExchange?: unknown }).officeExchange
  return clone
}

interface OfflineGuard {
  /** Fail unless nothing refused by the guard was ever consulted. */
  assertUntouched(): void
  /** Restore every patched transport surface. */
  restore(): void
}

/**
 * Replace outbound network and external-process entry points with a refusal for
 * one bounded operation. HTTP, HTTPS, `fetch`, DNS, TCP sockets, and child
 * processes are the surfaces a document server or cloud client would need.
 * @returns the guard's assertion and restoration handles.
 */
function blockExternalTransports(): OfflineGuard {
  const consulted: string[] = []
  const refuse = (surface: string) => (): never => {
    consulted.push(surface)
    throw new Error(`FR-010: ${surface} must not be consulted for a local Office edit`)
  }
  const spies = [
    vi.spyOn(globalThis, 'fetch').mockImplementation(refuse('fetch')),
    vi.spyOn(http, 'request').mockImplementation(refuse('http.request')),
    vi.spyOn(https, 'request').mockImplementation(refuse('https.request')),
    vi.spyOn(dns, 'lookup').mockImplementation(refuse('dns.lookup')),
    vi.spyOn(dns.promises, 'lookup').mockImplementation(refuse('dns.promises.lookup')),
    vi.spyOn(net.Socket.prototype, 'connect').mockImplementation(refuse('net.Socket.connect')),
    vi.spyOn(childProcess, 'spawn').mockImplementation(refuse('child_process.spawn')),
    vi.spyOn(childProcess, 'exec').mockImplementation(refuse('child_process.exec')),
    vi.spyOn(childProcess, 'execFile').mockImplementation(refuse('child_process.execFile')),
    vi.spyOn(childProcess, 'fork').mockImplementation(refuse('child_process.fork')),
  ]
  return {
    assertUntouched(): void {
      expect(consulted).toEqual([])
      for (const spy of spies) expect(spy).not.toHaveBeenCalled()
    },
    restore(): void {
      for (const spy of spies) spy.mockRestore()
    },
  }
}

/** Run one operation with every external transport blocked, restoring afterwards. */
async function withoutExternalTransports<T>(
  run: (assertUntouched: () => void) => Promise<T>,
): Promise<T> {
  const guard = blockExternalTransports()
  try {
    return await run(() => { guard.assertUntouched() })
  } finally {
    guard.restore()
  }
}

/** Save provider output to disk, check its OOXML signature, and reopen it through the provider. */
async function saveAndReopen(
  exchange: OfficeExchange,
  directory: string,
  name: string,
  bytes: Uint8Array,
): Promise<ExchangeReadResult> {
  expect(Array.from(bytes.slice(0, 2))).toEqual([0x50, 0x4b])
  expect(bytes.byteLength).toBeGreaterThan(0)
  const path = join(directory, name)
  await writeFile(path, bytes)
  return await exchange.read(path)
}

/**
 * Index one fixture as the source package a save splices into.
 * @param path - Fixture path.
 * @returns The indexed package carrying the original bytes.
 */
async function sourcePackageOf(path: string): Promise<SourcePackage> {
  return open(new Uint8Array(await readFile(path)), { maxEntries: 10_000 }).source
}

it('configures no document server endpoint', () => {
  expect(Object.keys(Config({})).sort()).toEqual([
    'commandTimeoutMs',
    'maxCells',
    'maxConversionMs',
    'maxFileBytes',
    'maxSourceAnchorBytes',
    'maxSourceAnchors',
  ])
})

it('edits and saves a workbook with no external transport', async () => {
  const exchange = await mount()
  const workspace = await mkdtemp(join(tmpdir(), 'dsh-office-offline-xlsx-'))
  try {
    await withoutExternalTransports(async (assertUntouched) => {
      const source = await sourcePackageOf(XLSX_FIXTURE)
      const read = await exchange.read(XLSX_FIXTURE)
      expect(read.unsupported).toEqual([])
      // The splice path carries every entry when the model matches the source.
      const baseline = await exchange.write(read.snapshot, 'xlsx', source)
      expect(baseline.bytes).toEqual(source.bytes)
      expect(baseline.disposition.spliced.count).toBe(0)
      const baselineBytes = baseline.bytes

      const editedSnapshot = structuredClone(read.snapshot)
      if (editedSnapshot.kind !== 'workbook') throw new Error('expected a workbook snapshot')
      const editedCell = cellAt(sheetOf(editedSnapshot.data, 'Sales'), 1, 0)
      if (editedCell === undefined) throw new Error('missing fixture cell B2')
      expect(editedCell.v).toBe('Widget')
      editedCell.v = 'Gadget'
      const edited = await exchange.write(editedSnapshot, 'xlsx', source)
      expect(edited.disposition.spliced.names).toEqual(['xl/worksheets/sheet1.xml'])
      expect(edited.bytes).not.toEqual(baselineBytes)
      const editedBytes = edited.bytes

      const baselineRead = workbookOf(await saveAndReopen(exchange, workspace, 'baseline.xlsx', baselineBytes))
      const saved = workbookOf(await saveAndReopen(exchange, workspace, 'edited.xlsx', editedBytes))

      const expected = structuredClone(baselineRead)
      const expectedCell = cellAt(sheetOf(expected, 'Sales'), 1, 0)
      if (expectedCell === undefined) throw new Error('missing saved fixture cell B2')
      expectedCell.v = 'Gadget'
      expect(withoutSidecar(saved)).toEqual(withoutSidecar(expected))

      const sales = sheetOf(saved, 'Sales')
      expect(saved.sheetOrder.map(id => saved.sheets[id]?.name)).toEqual(['Sales', 'Notes'])
      expect(cellAt(sales, 0, 0)?.v).toBe('Item')
      expect(cellAt(sales, 1, 0)?.v).toBe('Gadget')
      expect(cellAt(sales, 1, 2)?.v).toBe(1200.5)
      expect(cellAt(sales, 4, 1)?.f).toBe('=SUM(B2:B4)')
      expect(sales.mergeData).toContainEqual({
        startRow: 6,
        startColumn: 0,
        endRow: 7,
        endColumn: 3,
      })
      expect(cellAt(sheetOf(saved, 'Notes'), 0, 0)?.v).toBe('Note')

      assertUntouched()
    })
  } finally {
    await rm(workspace, { recursive: true, force: true })
  }
})

it('edits and saves a document with no external transport', async () => {
  const exchange = await mount()
  const workspace = await mkdtemp(join(tmpdir(), 'dsh-office-offline-docx-'))
  try {
    await withoutExternalTransports(async (assertUntouched) => {
      const source = await sourcePackageOf(DOCX_FIXTURE)
      const read = await exchange.read(DOCX_FIXTURE)
      expect(read.unsupported).toEqual([])
      // The splice path carries every entry when the model matches the source.
      const baseline = await exchange.write(read.snapshot, 'docx', source)
      expect(baseline.bytes).toEqual(source.bytes)
      expect(baseline.disposition.spliced.count).toBe(0)
      const baselineBytes = baseline.bytes

      const editedSnapshot = structuredClone(read.snapshot)
      if (editedSnapshot.kind !== 'document') throw new Error('expected a document snapshot')
      const editedBody = editedSnapshot.data.body
      if (editedBody === undefined) throw new Error('missing document body')
      expect(editedBody.dataStream).toContain('bold text')
      editedBody.dataStream = editedBody.dataStream.replace('bold text', 'cold text')
      const edited = await exchange.write(editedSnapshot, 'docx', source)
      expect(edited.disposition.spliced.names).toEqual(['word/document.xml'])
      expect(edited.bytes).not.toEqual(baselineBytes)
      const editedBytes = edited.bytes

      const baselineDocument = documentOf(await saveAndReopen(exchange, workspace, 'baseline.docx', baselineBytes))
      const saved = documentOf(await saveAndReopen(exchange, workspace, 'edited.docx', editedBytes))

      const expected = structuredClone(baselineDocument)
      const expectedBody = expected.body
      if (expectedBody === undefined) throw new Error('missing saved document body')
      expectedBody.dataStream = expectedBody.dataStream.replace('bold text', 'cold text')
      expect(withoutSidecar(saved)).toEqual(withoutSidecar(expected))

      const savedBody = saved.body
      if (savedBody === undefined) throw new Error('missing saved document body')
      expect(savedBody.dataStream).toContain('cold text')
      expect(savedBody.dataStream).not.toContain('bold text')
      expect(savedBody.paragraphs).toEqual(baselineDocument.body?.paragraphs)
      expect(savedBody.tables).toHaveLength(1)
      expect(saved.tableSource?.['table-1']?.tableRows).toHaveLength(2)
      const editedRun = (savedBody.textRuns ?? [])
        .find(run => savedBody.dataStream.slice(run.st, run.ed) === 'cold text')
      expect(editedRun?.ts?.bl).toBe(BooleanNumber.TRUE)

      assertUntouched()
    })
  } finally {
    await rm(workspace, { recursive: true, force: true })
  }
})
