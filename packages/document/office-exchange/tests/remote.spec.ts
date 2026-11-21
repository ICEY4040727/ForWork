/** The `officeDocument` Remote surface reads, saves under a version guard, and confines paths. */
import { createHash } from 'node:crypto'
import { copyFile, mkdir, mkdtemp, readdir, readFile, rename, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join, relative } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Context } from '@forwork/cordis'
import { LocalFileSystem } from '@forwork/fs-local'
import { SessionId } from '@forwork/session/types'
import SessionStore from '@forwork/session'
import type { SessionEvent } from '@forwork/session'
import { documentIdFor } from '@forwork/session-document'
import { documentProjectionDefinition } from '@forwork/session-document'
import SessionProjectionRegistry from '@forwork/session-projection'
import type { DocumentId } from '@forwork/session'
import type { JsonValue } from '@forwork/util-values'
import { ConversionError, createConverter } from '@deepseek-ai/libreoffice-kit'
import { afterEach, beforeEach, expect, it } from 'vitest'
import LocalOfficeExchange, { type OfficeDocumentSnapshot, type PreservedConstructKind } from '../src/index.ts'
import { unzipEntries } from './ooxml-package.ts'

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

/** The save-fact disposition one real splice save reports for this file. */
function saveDisposition(
  verbatimCount: number,
  spliced: readonly string[],
): {
  readonly schemaVersion: 1
  readonly verbatimCount: number
  readonly spliced: { readonly count: number; readonly names: readonly string[] }
  readonly added: { readonly count: number; readonly names: readonly string[] }
  readonly removed: { readonly count: number; readonly names: readonly string[] }
} {
  return {
    schemaVersion: 1,
    verbatimCount,
    spliced: { count: spliced.length, names: spliced },
    added: { count: 0, names: [] },
    removed: { count: 0, names: [] },
  }
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

  const saved = await ctx.officeDocument.save(scope, path, read.snapshot, 'xlsx', read.version, { actor: 'user' }, signal)
  expect(saved.version).not.toBe(read.version)

  const reread = await ctx.officeDocument.read(scope, path, signal)
  expect(reread.version).toBe(saved.version)
  expect(workbookCell(reread.snapshot, 1, 0)).toBe('Sprocket')
})

it('saves a modified document at the read version and reads the change back', async () => {
  const path = await copyFixture('roundtrip.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  rewriteDocumentStream(read.snapshot, 'Quarterly', 'Financial')

  const saved = await ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, { actor: 'user' }, signal)
  expect(saved.version).not.toBe(read.version)

  const reread = await ctx.officeDocument.read(scope, path, signal)
  expect(documentStream(reread.snapshot)).toContain('Financial Report')
})

it('saves a workbook twice from one editor window with untouched entries byte-identical', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  const source = new Uint8Array(await readFile(join(root, path)))
  const read = await ctx.officeDocument.read(scope, path, signal)
  const untouched = ['docProps/core.xml', 'xl/sharedStrings.xml', 'xl/styles.xml', 'xl/theme/theme1.xml']

  setWorkbookCell(read.snapshot, 1, 0, 'Sprocket')
  const first = await ctx.officeDocument.save(
    scope, path, read.snapshot, 'xlsx', read.version, { actor: 'user' }, signal,
  )
  const afterFirst = new Uint8Array(await readFile(join(root, path)))

  // The same editor window keeps the snapshot read before the first save, so
  // its sidecar still names the original package digest.
  setWorkbookCell(read.snapshot, 2, 0, 'Cog')
  const second = await ctx.officeDocument.save(
    scope, path, read.snapshot, 'xlsx', first.version, { actor: 'user' }, signal,
  )
  expect(second.version).not.toBe(first.version)
  const afterSecond = new Uint8Array(await readFile(join(root, path)))

  const sourceEntries = unzipEntries(source)
  const firstEntries = unzipEntries(afterFirst)
  const secondEntries = unzipEntries(afterSecond)
  for (const name of untouched) {
    expect(firstEntries.get(name), name).toEqual(sourceEntries.get(name))
    expect(secondEntries.get(name), name).toEqual(firstEntries.get(name))
  }
  expect(firstEntries.get('xl/worksheets/sheet1.xml')).not.toEqual(sourceEntries.get('xl/worksheets/sheet1.xml'))
  expect(secondEntries.get('xl/worksheets/sheet1.xml')).not.toEqual(firstEntries.get('xl/worksheets/sheet1.xml'))

  const reread = await ctx.officeDocument.read(scope, path, signal)
  expect(reread.version).toBe(second.version)
  expect(workbookCell(reread.snapshot, 1, 0)).toBe('Sprocket')
  expect(workbookCell(reread.snapshot, 2, 0)).toBe('Cog')
})

it('saves a document twice from one editor window with untouched entries byte-identical', async () => {
  const path = await copyFixture('roundtrip.docx')
  const source = new Uint8Array(await readFile(join(root, path)))
  const read = await ctx.officeDocument.read(scope, path, signal)

  rewriteDocumentStream(read.snapshot, 'Quarterly', 'Financial')
  const first = await ctx.officeDocument.save(
    scope, path, read.snapshot, 'docx', read.version, { actor: 'user' }, signal,
  )
  const afterFirst = new Uint8Array(await readFile(join(root, path)))

  rewriteDocumentStream(read.snapshot, 'bold text', 'cold text')
  const second = await ctx.officeDocument.save(
    scope, path, read.snapshot, 'docx', first.version, { actor: 'user' }, signal,
  )
  expect(second.version).not.toBe(first.version)
  const afterSecond = new Uint8Array(await readFile(join(root, path)))

  const sourceEntries = unzipEntries(source)
  const firstEntries = unzipEntries(afterFirst)
  const secondEntries = unzipEntries(afterSecond)
  expect(firstEntries.get('word/styles.xml')).toEqual(sourceEntries.get('word/styles.xml'))
  expect(secondEntries.get('word/styles.xml')).toEqual(firstEntries.get('word/styles.xml'))
  expect(firstEntries.get('word/document.xml')).not.toEqual(sourceEntries.get('word/document.xml'))
  expect(secondEntries.get('word/document.xml')).not.toEqual(firstEntries.get('word/document.xml'))

  const reread = await ctx.officeDocument.read(scope, path, signal)
  expect(reread.version).toBe(second.version)
  expect(documentStream(reread.snapshot)).toContain('Financial Report')
  expect(documentStream(reread.snapshot)).toContain('cold text')
})

/** Deterministic incompressible text, so an added cell value cannot shrink the package. */
function noise(length: number): string {
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
  let state = 0x12345678
  let text = ''
  for (let index = 0; index < length; index += 1) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    // High bits of a power-of-two LCG are the well-distributed ones.
    text += alphabet.charAt((state >>> 8) % alphabet.length)
  }
  return text
}

it('refuses an output over the byte bound as bound-exceeded and writes nothing', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  const sourceBytes = new Uint8Array(await readFile(join(root, path)))
  await ctx.fiber.dispose()
  ctx = new Context()
  await ctx.plugin(LocalFileSystem, { cwd: root })
  await ctx.plugin(SessionStore)
  ctx.sessions.create(scope.sessionId, { meta: { cwd: root } })
  await ctx.plugin(LocalOfficeExchange, { maxFileBytes: sourceBytes.byteLength })

  const read = await ctx.officeDocument.read(scope, path, signal)
  setWorkbookCell(read.snapshot, 1, 0, noise(64 * 1024))
  await expect(ctx.officeDocument.save(
    scope, path, read.snapshot, 'xlsx', read.version, { actor: 'user' }, signal,
  )).rejects.toMatchObject({ code: 'office-document/bound-exceeded' })
  expect(new Uint8Array(await readFile(join(root, path)))).toEqual(sourceBytes)
})

it('refuses a sidecar over the anchor budget as bound-exceeded', async () => {
  const path = await copyFixture('roundtrip.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  const sourceBytes = new Uint8Array(await readFile(join(root, path)))
  await ctx.fiber.dispose()
  ctx = new Context()
  await ctx.plugin(LocalFileSystem, { cwd: root })
  await ctx.plugin(SessionStore)
  ctx.sessions.create(scope.sessionId, { meta: { cwd: root } })
  await ctx.plugin(LocalOfficeExchange, { maxSourceAnchors: 1 })

  await expect(ctx.officeDocument.save(
    scope, path, read.snapshot, 'docx', read.version, { actor: 'user' }, signal,
  )).rejects.toMatchObject({ code: 'office-document/bound-exceeded' })
  expect(new Uint8Array(await readFile(join(root, path)))).toEqual(sourceBytes)
})

it('maps a sidecar format mismatch to unsupported-format', async () => {
  const path = await copyFixture('roundtrip.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  const data = jsonObject(read.snapshot.data)
  const source = jsonObject(jsonObject(data['officeExchange'])['source'])
  source['format'] = 'xlsx'

  await expect(ctx.officeDocument.save(
    scope, path, read.snapshot, 'docx', read.version, { actor: 'user' }, signal,
  )).rejects.toMatchObject({ code: 'office-document/unsupported-format' })
})

it('refuses a stale save as a conflict and leaves the file unchanged', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  setWorkbookCell(read.snapshot, 1, 0, 'Sprocket')

  await writeFile(join(root, path), 'external change')
  await expect(ctx.officeDocument.save(scope, path, read.snapshot, 'xlsx', read.version, { actor: 'user' }, signal))
    .rejects.toMatchObject({ code: 'office-document/conflict' })
  expect(await readFile(join(root, path), 'utf8')).toBe('external change')
})

it('refuses a save whose file disappeared as a conflict', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  setWorkbookCell(read.snapshot, 1, 0, 'Sprocket')
  await rm(join(root, path))
  await expect(ctx.officeDocument.save(scope, path, read.snapshot, 'xlsx', read.version, { actor: 'user' }, signal))
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
    await expect(ctx.officeDocument.save(scope, path, snapshot, 'xlsx', 'version', { actor: 'user' }, signal))
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

/** Fixtures that carry preserved constructs, with the classes each one carries. */
const PRESERVED_CORPUS: readonly (readonly [string, readonly PreservedConstructKind[]])[] = [
  ['macro.docm', ['macro']],
  ['macro.xlsm', ['document-properties', 'macro', 'theme']],
  ['preserve-bookmark.docx', ['bookmark']],
  ['preserve-calc-chain.xlsx', ['calc-chain']],
  ['preserve-chart.docx', ['chart']],
  ['preserve-comment.docx', ['comment']],
  ['preserve-content-control.docx', ['content-control']],
  ['preserve-defined-names.xlsx', ['defined-names']],
  ['preserve-embedded-object.docx', ['embedded-object']],
  ['preserve-excel-table.xlsx', ['excel-table']],
  ['preserve-field.docx', ['field']],
  ['preserve-pivot-table.xlsx', ['pivot-table']],
  ['preserve-calc-chain.xlsx', ['calc-chain']],
  ['preserve-shape.docx', ['drawing', 'shape']],
  ['preserve-unmodeled-parts.docx', ['custom-xml', 'document-properties', 'font-table', 'theme', 'web-settings']],
  ['roundtrip.xlsx', ['document-properties', 'theme']],
]

/**
 * Every preserved class the fixture corpus is built to carry. `external-link`
 * has no corpus fixture: the docx hyperlink fixture targets an external URL,
 * which is a relationship, not a workbook external-link part.
 */
const CORPUS_CLASSES: readonly PreservedConstructKind[] = [
  'bookmark', 'calc-chain', 'chart', 'comment', 'content-control', 'custom-xml', 'defined-names',
  'document-properties', 'drawing', 'embedded-object', 'excel-table', 'field', 'font-table',
  'macro', 'pivot-table', 'shape', 'theme', 'web-settings',
]

it('reports every preserved class the fixture corpus carries', async () => {
  const observed = new Set<PreservedConstructKind>()
  for (const [name, expected] of PRESERVED_CORPUS) {
    const path = await copyFixture(name)
    const read = await ctx.officeDocument.read(scope, path, signal)
    expect(read.preserved, name).toEqual(expected)
    for (const kind of expected) observed.add(kind)
  }
  expect([...observed].sort()).toEqual([...CORPUS_CLASSES].sort())
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
  await ctx.officeDocument.save(scope, path, read.snapshot, 'xlsx', read.version, { actor: 'user' }, signal)

  const saves = events.filter(event => event.type === 'document/save')
  expect(saves).toHaveLength(1)
  expect(saves[0]?.data).toEqual({
    documentId: documentIdOf(path),
    actor: 'user',
    contentVersion: 1,
    disposition: saveDisposition(10, ['xl/worksheets/sheet1.xml']),
  })
})


it('returns the writer disposition with the saved version', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  setWorkbookCell(read.snapshot, 1, 0, 'Sprocket')

  const saved = await ctx.officeDocument.save(
    scope, path, read.snapshot, 'xlsx', read.version, { actor: 'user' }, signal,
  )

  expect(saved.disposition).toEqual({
    ...saveDisposition(10, ['xl/worksheets/sheet1.xml']),
    nameLimit: 256,
  })
})

it('keeps one document identity when the same file is read and saved by different paths', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  await ctx.officeDocument.read(scope, join(root, path), signal)
  setWorkbookCell(read.snapshot, 1, 0, 'Sprocket')
  await ctx.officeDocument.save(
    scope, join(root, path), read.snapshot, 'xlsx', read.version, { actor: 'agent', turn: 3 }, signal,
  )

  const opens = events.filter(event => event.type === 'document/open')
  expect(opens).toHaveLength(1)
  expect(opens[0]?.data).toMatchObject({ documentId: documentIdOf(path), path })
  const saves = events.filter(event => event.type === 'document/save')
  expect(saves).toHaveLength(1)
  expect(saves[0]?.data).toEqual({
    documentId: documentIdOf(path),
    actor: 'agent',
    contentVersion: 1,
    turn: 3,
    disposition: saveDisposition(10, ['xl/worksheets/sheet1.xml']),
  })
})

it('keeps same-named files in different directories as distinct documents', async () => {
  await mkdir(join(root, 'a'))
  await mkdir(join(root, 'b'))
  await copyFile(join(FIXTURES, 'roundtrip.xlsx'), join(root, 'a', 'report.xlsx'))
  await copyFile(join(FIXTURES, 'roundtrip.xlsx'), join(root, 'b', 'report.xlsx'))
  const first = await ctx.officeDocument.read(scope, 'a/report.xlsx', signal)
  await ctx.officeDocument.read(scope, 'b/report.xlsx', signal)
  await ctx.officeDocument.save(scope, 'a/report.xlsx', first.snapshot, 'xlsx', first.version, { actor: 'user' }, signal)

  const opens = events.filter(event => event.type === 'document/open')
  expect(opens.map(event => event.data.documentId)).toEqual([
    documentIdOf('a/report.xlsx'),
    documentIdOf('b/report.xlsx'),
  ])
  const saves = events.filter(event => event.type === 'document/save')
  expect(saves).toHaveLength(1)
  expect(saves[0]?.data).toEqual({
    documentId: documentIdOf('a/report.xlsx'),
    actor: 'user',
    contentVersion: 1,
    disposition: saveDisposition(11, []),
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

it('saves a rebound document under the identity the rebind preserved', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  await copyFile(join(root, path), join(root, 'renamed.xlsx'))
  await ctx.officeDocument.read(scope, path, signal)
  await ctx.officeDocument.rebindDocument(scope, documentIdOf(path), 'renamed.xlsx', signal)

  const moved = await ctx.officeDocument.read(scope, 'renamed.xlsx', signal)
  setWorkbookCell(moved.snapshot, 1, 0, 'Sprocket')
  await ctx.officeDocument.save(
    scope, 'renamed.xlsx', moved.snapshot, 'xlsx', moved.version, { actor: 'user' }, signal,
  )

  expect(events.filter(event => event.type === 'document/open')).toHaveLength(1)
  const saves = events.filter(event => event.type === 'document/save')
  expect(saves).toHaveLength(1)
  expect(saves[0]?.data).toEqual({
    documentId: documentIdOf(path),
    actor: 'user',
    contentVersion: 1,
    disposition: saveDisposition(10, ['xl/worksheets/sheet1.xml']),
  })
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

it('keeps a path-derived identity when no document projection is registered', async () => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-document-no-projection-'))
  const localCtx = new Context()
  const localEvents: SessionEvent[] = []
  try {
    await localCtx.plugin(LocalFileSystem, { cwd: root })
    await localCtx.plugin(SessionStore)
    await localCtx.plugin(SessionProjectionRegistry)
    await localCtx.plugin(LocalOfficeExchange)
    const localScope = { sessionId: SessionId('office-document-no-projection'), workspaceRoot: root }
    const session = localCtx.sessions.create(localScope.sessionId, { meta: { cwd: root } })
    localCtx.on('session/event', (subject, event) => {
      if (subject === session) localEvents.push(event)
    })
    await copyFile(join(FIXTURES, 'roundtrip.xlsx'), join(root, 'roundtrip.xlsx'))

    await localCtx.officeDocument.read(localScope, 'roundtrip.xlsx', signal)

    expect(localEvents.filter(event => event.type === 'document/open').map(event => event.data.documentId))
      .toEqual([documentIdFor(localScope.sessionId, 'roundtrip.xlsx')])
  } finally {
    await localCtx.fiber.dispose()
    await rm(root, { recursive: true, force: true })
  }
})

it('reopens a moved document under its projected identity after an Office restart', async () => {
  const restartRoot = await mkdtemp(join(tmpdir(), 'dsh-office-document-restart-'))
  const restartCtx = new Context()
  const restartEvents: SessionEvent[] = []
  try {
    await restartCtx.plugin(LocalFileSystem, { cwd: restartRoot })
    await restartCtx.plugin(SessionStore)
    await restartCtx.plugin(SessionProjectionRegistry)
    const projection = await restartCtx.sessionProjections.register(documentProjectionDefinition)
    const restartScope = { sessionId: SessionId('office-document-restart'), workspaceRoot: restartRoot }
    const session = restartCtx.sessions.create(restartScope.sessionId, { meta: { cwd: restartRoot } })
    restartCtx.on('session/event', (subject, event) => {
      if (subject === session) restartEvents.push(event)
    })
    const exchange = await restartCtx.plugin(LocalOfficeExchange)
    await copyFile(join(FIXTURES, 'roundtrip.xlsx'), join(restartRoot, 'roundtrip.xlsx'))
    await restartCtx.officeDocument.read(restartScope, 'roundtrip.xlsx', signal)
    const documentId = documentIdFor(restartScope.sessionId, 'roundtrip.xlsx')
    await rename(join(restartRoot, 'roundtrip.xlsx'), join(restartRoot, 'renamed.xlsx'))
    await restartCtx.officeDocument.rebindDocument(restartScope, documentId, 'renamed.xlsx', signal)

    await exchange.dispose()
    await restartCtx.plugin(LocalOfficeExchange)

    const reopened = await restartCtx.officeDocument.read(restartScope, 'renamed.xlsx', signal)
    setWorkbookCell(reopened.snapshot, 1, 0, 'Sprocket')
    await restartCtx.officeDocument.save(
      restartScope, 'renamed.xlsx', reopened.snapshot, 'xlsx', reopened.version, { actor: 'user' }, signal,
    )

    expect(projection.stateOf(session).order).toEqual([documentId])
    expect(projection.stateOf(session).documents[documentId]).toMatchObject({
      documentId,
      path: 'renamed.xlsx',
      status: 'open',
    })
    expect(restartEvents.filter(event => event.type === 'document/open')).toHaveLength(1)
    expect(restartEvents.filter(event => event.type === 'document/save').map(event => event.data)).toEqual([{
      documentId,
      actor: 'user',
      contentVersion: 1,
      disposition: saveDisposition(10, ['xl/worksheets/sheet1.xml']),
    }])
  } finally {
    await restartCtx.fiber.dispose()
    await rm(restartRoot, { recursive: true, force: true })
  }
})

it('keeps the disposition readable after reopening the session log', async () => {
  const path = await copyFixture('roundtrip.xlsx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  setWorkbookCell(read.snapshot, 1, 0, 'Sprocket')
  await ctx.officeDocument.save(scope, path, read.snapshot, 'xlsx', read.version, { actor: 'user' }, signal)

  const persisted = JSON.parse(JSON.stringify(events)) as SessionEvent[]
  const saved = persisted.find(event => event.type === 'document/save')
  expect(saved?.type === 'document/save' ? saved.data.disposition : undefined)
    .toEqual(saveDisposition(10, ['xl/worksheets/sheet1.xml']))

  const restartCtx = new Context()
  try {
    await restartCtx.plugin(SessionStore)
    await restartCtx.plugin(SessionProjectionRegistry)
    const projection = await restartCtx.sessionProjections.register(documentProjectionDefinition)
    const restored = restartCtx.sessions.create(scope.sessionId, { seed: persisted, meta: { cwd: root } })
    const state = projection.stateOf(restored)
    expect(state.documents[documentIdOf(path)]).toMatchObject({
      documentId: documentIdOf(path),
      path,
      format: 'xlsx',
      contentVersion: 1,
      status: 'open',
    })
  } finally {
    await restartCtx.fiber.dispose()
  }
})
