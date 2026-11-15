/** The `officeDocument` Remote surface reads, saves under a version guard, and confines paths. */
import { createHash } from 'node:crypto'
import { copyFile, mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join, relative } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Context } from '@deepseek-ai/cordis'
import { LocalFileSystem } from '@deepseek-ai/dsh-fs-local'
import { SessionId } from '@deepseek-ai/dsh-session/types'
import SessionStore from '@deepseek-ai/dsh-session'
import type { SessionEvent } from '@deepseek-ai/dsh-session'
import { documentIdFor } from '@deepseek-ai/dsh-session-document'
import type { DocumentId } from '@deepseek-ai/dsh-session'
import type { JsonValue } from '@deepseek-ai/dsh-util-values'
import { ConversionError, createConverter } from '@deepseek-ai/libreoffice-kit'
import { afterEach, beforeEach, expect, it } from 'vitest'
import LocalOfficeExchange, { type OfficeDocumentSnapshot } from '../src/index.ts'

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url))
const scope = { sessionId: SessionId('office-document-test'), workspaceRoot: '' }

let ctx: Context
let root: string
const signal = new AbortController().signal
const events: SessionEvent[] = []

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'dsh-office-document-'))
  scope.workspaceRoot = root
  events.length = 0
  ctx = new Context()
  await ctx.plugin(LocalFileSystem, { cwd: root })
  await ctx.plugin(SessionStore)
  const session = ctx.sessions.create(scope.sessionId, { meta: { cwd: root } })
  ctx.on('session/event', (subject, event) => {
    if (subject === session) events.push(event)
  })
  await ctx.plugin(LocalOfficeExchange)
})
afterEach(async () => {
  await ctx.fiber.dispose()
  await rm(root, { recursive: true, force: true })
})

async function copyFixture(name: string): Promise<string> {
  await copyFile(join(FIXTURES, name), join(root, name))
  return name
}

function jsonObject(value: JsonValue | undefined): Record<string, JsonValue> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) throw new Error('expected a JSON object')
  return value
}

function workbookCell(snapshot: OfficeDocumentSnapshot, row: number, column: number): JsonValue {
  if (snapshot.kind !== 'workbook') throw new Error('expected a workbook snapshot')
  const data = jsonObject(snapshot.data)
  const order = data['sheetOrder']
  if (!Array.isArray(order) || typeof order[0] !== 'string') throw new Error('expected a sheet order')
  const sheet = jsonObject(jsonObject(data['sheets'])[order[0]])
  return jsonObject(jsonObject(jsonObject(sheet['cellData'])[String(row)])[String(column)])['v'] ?? null
}

function setWorkbookCell(snapshot: OfficeDocumentSnapshot, row: number, column: number, value: string): void {
  if (snapshot.kind !== 'workbook') throw new Error('expected a workbook snapshot')
  const data = jsonObject(snapshot.data)
  const order = data['sheetOrder']
  if (!Array.isArray(order) || typeof order[0] !== 'string') throw new Error('expected a sheet order')
  const sheet = jsonObject(jsonObject(data['sheets'])[order[0]])
  const cells = jsonObject(sheet['cellData'])
  const rowCells = jsonObject(cells[String(row)])
  jsonObject(rowCells[String(column)])['v'] = value
}

function documentStream(snapshot: OfficeDocumentSnapshot): string {
  if (snapshot.kind !== 'document') throw new Error('expected a document snapshot')
  const body = jsonObject(jsonObject(snapshot.data)['body'])
  const stream = body['dataStream']
  if (typeof stream !== 'string') throw new Error('expected a document data stream')
  return stream
}

function rewriteDocumentStream(snapshot: OfficeDocumentSnapshot, from: string, to: string): void {
  if (snapshot.kind !== 'document') throw new Error('expected a document snapshot')
  const body = jsonObject(jsonObject(snapshot.data)['body'])
  body['dataStream'] = documentStream(snapshot).replace(from, to)
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

it('reads workbook and document fixtures with the file version', async () => {
  const workbookPath = await copyFixture('roundtrip.xlsx')
  const workbook = await ctx.officeDocument.read(scope, workbookPath, signal)
  expect(workbook.format).toBe('xlsx')
  expect(workbook.snapshot.kind).toBe('workbook')
  expect(workbook.unsupported).toEqual([])
  expect(workbook.version).not.toBe('')
  expect(workbookCell(workbook.snapshot, 1, 0)).toBe('Widget')

  const documentPath = await copyFixture('roundtrip.docx')
  const document = await ctx.officeDocument.read(scope, documentPath, signal)
  expect(document.format).toBe('docx')
  expect(document.snapshot.kind).toBe('document')
  expect(document.unsupported).toEqual([])
  expect(document.version).not.toBe('')
  expect(documentStream(document.snapshot)).toContain('Quarterly Report')
  expect(document.version).not.toBe(workbook.version)
})

it('saves a modified workbook at the read version and reads the change back', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  setWorkbookCell(read.snapshot, 1, 0, 'Sprocket')

  const saved = await ctx.officeDocument.save(scope, path, read.snapshot, 'xlsx', read.version, signal)
  expect(saved.version).not.toBe(read.version)

  const reread = await ctx.officeDocument.read(scope, path, signal)
  expect(reread.version).toBe(saved.version)
  expect(workbookCell(reread.snapshot, 1, 0)).toBe('Sprocket')
})

it('saves a modified document at the read version and reads the change back', async () => {
  const path = await copyFixture('roundtrip.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  rewriteDocumentStream(read.snapshot, 'Quarterly', 'Financial')

  const saved = await ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, signal)
  expect(saved.version).not.toBe(read.version)

  const reread = await ctx.officeDocument.read(scope, path, signal)
  expect(documentStream(reread.snapshot)).toContain('Financial Report')
})

it('refuses a stale save as a conflict and leaves the file unchanged', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  setWorkbookCell(read.snapshot, 1, 0, 'Sprocket')

  await writeFile(join(root, path), 'external change')
  await expect(ctx.officeDocument.save(scope, path, read.snapshot, 'xlsx', read.version, signal))
    .rejects.toMatchObject({ code: 'office-document/conflict' })
  expect(await readFile(join(root, path), 'utf8')).toBe('external change')
})

it('refuses a save whose file disappeared as a conflict', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  setWorkbookCell(read.snapshot, 1, 0, 'Sprocket')
  await rm(join(root, path))
  await expect(ctx.officeDocument.save(scope, path, read.snapshot, 'xlsx', read.version, signal))
    .rejects.toMatchObject({ code: 'office-document/conflict' })
})

it('refuses reads and saves outside the workspace root', async () => {
  const outside = await mkdtemp(join(tmpdir(), 'dsh-office-document-outside-'))
  try {
    const path = join(outside, 'roundtrip.xlsx')
    await copyFile(join(FIXTURES, 'roundtrip.xlsx'), path)
    await expect(ctx.officeDocument.read(scope, path, signal))
      .rejects.toMatchObject({ code: 'office-document/out-of-scope' })

    const snapshot: OfficeDocumentSnapshot = { kind: 'workbook', data: {} }
    await expect(ctx.officeDocument.save(scope, path, snapshot, 'xlsx', 'version', signal))
      .rejects.toMatchObject({ code: 'office-document/out-of-scope' })
  } finally {
    await rm(outside, { recursive: true, force: true })
  }
})

it('reports a macro feature on the macro fixture', async () => {
  const path = await copyFixture('unsupported-macro.docx')
  const features = await ctx.officeDocument.detectUnsupported(scope, path, signal)
  expect(features.some(feature => feature.kind === 'macro')).toBe(true)
})

function hashOf(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

it('converts a legacy workbook inside the workspace root and reads it back', { timeout: 180_000 }, async (context) => {
  const path = await copyFixture('legacy.xls')
  const reason = await unavailableReason()
  if (reason !== undefined) {
    context.skip(`LibreOffice unavailable: ${reason}`)
    return
  }
  const sourceBefore = await readFile(join(root, path))
  const converted = await ctx.officeDocument.convertLegacy(scope, path, signal)
  expect(converted.endsWith('.xlsx')).toBe(true)
  expect(relative(root, converted).startsWith('..')).toBe(false)
  expect((await readFile(converted)).byteLength).toBeGreaterThan(0)

  const reopened = await ctx.officeDocument.read(scope, converted, signal)
  expect(reopened.format).toBe('xlsx')
  expect(workbookCell(reopened.snapshot, 1, 0)).toBe('Widget')
  expect(hashOf(await readFile(join(root, path)))).toBe(hashOf(sourceBefore))
})

it('removes the conversion directory it created when conversion fails', async () => {
  await writeFile(join(root, 'notes.txt'), 'plain text')
  await expect(ctx.officeDocument.convertLegacy(scope, 'notes.txt', signal))
    .rejects.toMatchObject({ code: 'office-document/unsupported-format' })
  expect(await readdir(join(root, '.dsh', 'office-tmp'))).toEqual([])
})

it('keeps two conversions of the same source apart', { timeout: 180_000 }, async (context) => {
  const path = await copyFixture('legacy.xls')
  const reason = await unavailableReason()
  if (reason !== undefined) {
    context.skip(`LibreOffice unavailable: ${reason}`)
    return
  }
  const first = await ctx.officeDocument.convertLegacy(scope, path, signal)
  const second = await ctx.officeDocument.convertLegacy(scope, path, signal)
  expect(dirname(first)).not.toBe(dirname(second))
  expect((await readFile(first)).byteLength).toBeGreaterThan(0)
  expect((await readFile(second)).byteLength).toBeGreaterThan(0)
  expect((await ctx.officeDocument.read(scope, first, signal)).format).toBe('xlsx')
  expect((await ctx.officeDocument.read(scope, second, signal)).format).toBe('xlsx')
})




/** The Session-scoped identity the client and the tools use for one path. */
function documentIdOf(path: string): DocumentId {
  return documentIdFor(scope.sessionId, path)
}

it('appends document/open once for a document a Session opened', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  await ctx.officeDocument.read(scope, path, signal)
  await ctx.officeDocument.read(scope, path, signal)

  const opens = events.filter(event => event.type === 'document/open')
  expect(opens).toHaveLength(1)
  expect(opens[0]?.data).toEqual({
    documentId: documentIdOf(path),
    path,
    format: 'xlsx',
    contentVersion: 0,
  })
})

it('appends document/save with the user actor and the next content version', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  setWorkbookCell(read.snapshot, 1, 0, 'Sprocket')
  await ctx.officeDocument.save(scope, path, read.snapshot, 'xlsx', read.version, signal)

  const saves = events.filter(event => event.type === 'document/save')
  expect(saves).toHaveLength(1)
  expect(saves[0]?.data).toEqual({
    documentId: documentIdOf(path),
    actor: 'user',
    contentVersion: 1,
  })
})

it('closes a document and opens it again on the next read', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  await ctx.officeDocument.read(scope, path, signal)
  await ctx.officeDocument.closeDocument(scope, documentIdOf(path), 'closed', signal)
  await ctx.officeDocument.read(scope, path, signal)

  const closed = events.filter(event => event.type === 'document/close')
  expect(closed).toHaveLength(1)
  expect(closed[0]?.data).toEqual({ documentId: documentIdOf(path), reason: 'closed' })
  expect(events.filter(event => event.type === 'document/open')).toHaveLength(2)
})

it('rebinds a moved document without changing its identity', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  await copyFile(join(root, path), join(root, 'renamed.xlsx'))
  await ctx.officeDocument.read(scope, path, signal)
  const open = events.filter(event => event.type === 'document/open')[0]

  await ctx.officeDocument.rebindDocument(scope, documentIdOf(path), 'renamed.xlsx', signal)

  const rebound = events.filter(event => event.type === 'document/rebind')
  expect(rebound).toHaveLength(1)
  expect(rebound[0]?.data).toEqual({
    documentId: documentIdOf(path),
    fromPath: path,
    toPath: 'renamed.xlsx',
  })
  expect(events.filter(event => event.type === 'document/open')[0]).toBe(open)
})

it('refuses a rebind outside the workspace root', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  await ctx.officeDocument.read(scope, path, signal)
  await expect(ctx.officeDocument.rebindDocument(scope, documentIdOf(path), '../outside.xlsx', signal))
    .rejects.toMatchObject({ code: 'office-document/out-of-scope' })
})

it('marks a document missing once while keeping it open', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  await ctx.officeDocument.read(scope, path, signal)
  await ctx.officeDocument.markDocumentMissing(scope, documentIdOf(path), signal)
  await ctx.officeDocument.markDocumentMissing(scope, documentIdOf(path), signal)

  const missing = events.filter(event => event.type === 'document/missing')
  expect(missing).toHaveLength(1)
  expect(missing[0]?.data).toEqual({ documentId: documentIdOf(path), path })
  expect(events.filter(event => event.type === 'document/close')).toHaveLength(0)
})

it('refuses a document identity from another Session', async () => {
  await expect(ctx.officeDocument.closeDocument(scope, 'other-session:report.xlsx', 'closed', signal))
    .rejects.toMatchObject({ code: 'gateway/bad-request' })
})
