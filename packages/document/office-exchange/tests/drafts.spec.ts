/** The `officeDocument` Remote persists, reads, and clears recovery drafts inside the workspace. */
import { createHash } from 'node:crypto'
import { copyFile, mkdir, mkdtemp, readdir, readFile, rm, symlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Context } from '@forwork/cordis'
import { LocalFileSystem } from '@forwork/fs-local'
import { SessionId } from '@forwork/session/types'
import SessionStore from '@forwork/session'
import { afterEach, beforeEach, expect, it } from 'vitest'
import LocalOfficeExchange, { type OfficeDocumentDraft, type OfficeDocumentSnapshot } from '../src/index.ts'

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url))
const scope = { sessionId: SessionId('office-draft-test'), workspaceRoot: '' }

const DOCUMENT_ID = 'draft-document'
const VERSION = 'version-1'
const WORKBOOK: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'workbook-1' } }
const DOCUMENT: OfficeDocumentSnapshot = { kind: 'document', data: { id: 'document-1' } }

let ctx: Context
let root: string
const signal = new AbortController().signal

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'dsh-office-draft-'))
  scope.workspaceRoot = root
  ctx = new Context()
  await ctx.plugin(LocalFileSystem, { cwd: root })
  await ctx.plugin(SessionStore)
  ctx.sessions.create(scope.sessionId, { meta: { cwd: root } })
  await ctx.plugin(LocalOfficeExchange)
})
afterEach(async () => {
  await ctx.fiber.dispose()
  await rm(root, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 })
})

function draftsDirectory(): string {
  return join(root, '.dsh', 'office-tmp', 'drafts')
}

async function copyFixture(name: string): Promise<void> {
  await copyFile(join(FIXTURES, name), join(root, name))
}

async function writeDefaultDraft(): Promise<void> {
  await ctx.officeDocument.writeDraft(scope, DOCUMENT_ID, 'roundtrip.xlsx', VERSION, WORKBOOK, signal)
}

async function storedDraft(): Promise<OfficeDocumentDraft> {
  const draft = await ctx.officeDocument.readDraft(scope, DOCUMENT_ID, signal)
  if (draft === null) throw new Error('expected a stored draft')
  return draft
}

async function draftFileNames(): Promise<string[]> {
  // The filesystem stages a `.tmpdir` sibling during an atomic write; only the published draft counts.
  return (await readdir(draftsDirectory())).filter(name => name.endsWith('.json'))
}

async function draftFileName(): Promise<string> {
  const [name] = await draftFileNames()
  if (name === undefined) throw new Error('expected a stored draft file')
  return name
}

async function overwriteDraftFile(content: string): Promise<void> {
  await writeFile(join(draftsDirectory(), await draftFileName()), content, 'utf8')
}

it('writes and reads a draft round trip', async () => {
  const { savedAt } = await ctx.officeDocument.writeDraft(
    scope, DOCUMENT_ID, 'roundtrip.xlsx', VERSION, WORKBOOK, signal,
  )
  expect(Number.isNaN(Date.parse(savedAt))).toBe(false)

  const draft = await storedDraft()
  expect(draft).toEqual({
    documentId: DOCUMENT_ID,
    path: 'roundtrip.xlsx',
    contentVersion: VERSION,
    snapshot: WORKBOOK,
    savedAt,
  })
  const digest = createHash('sha256').update(DOCUMENT_ID).digest('hex')
  expect(await draftFileName()).toBe(`${digest}.json`)
})

it('replaces the stored draft when it is written again', async () => {
  await writeDefaultDraft()
  const second = await ctx.officeDocument.writeDraft(
    scope, DOCUMENT_ID, 'roundtrip.docx', 'version-2', DOCUMENT, signal,
  )

  const draft = await storedDraft()
  expect(draft.path).toBe('roundtrip.docx')
  expect(draft.contentVersion).toBe('version-2')
  expect(draft.snapshot).toEqual(DOCUMENT)
  expect(draft.savedAt).toBe(second.savedAt)
  expect(await draftFileNames()).toHaveLength(1)
})

it('clears a draft and treats a repeated clear as a no-op', async () => {
  await writeDefaultDraft()
  await ctx.officeDocument.clearDraft(scope, DOCUMENT_ID, signal)
  expect(await ctx.officeDocument.readDraft(scope, DOCUMENT_ID, signal)).toBeNull()

  await expect(ctx.officeDocument.clearDraft(scope, DOCUMENT_ID, signal)).resolves.toBeUndefined()
  expect(await draftFileNames()).toEqual([])
})

it('returns null when no draft exists', async () => {
  expect(await ctx.officeDocument.readDraft(scope, DOCUMENT_ID, signal)).toBeNull()
})

it('refuses draft storage that resolves outside the workspace root', async () => {
  const outside = await mkdtemp(join(tmpdir(), 'dsh-office-draft-outside-'))
  try {
    await mkdir(join(root, '.dsh', 'office-tmp'), { recursive: true })
    await symlink(outside, draftsDirectory(), process.platform === 'win32' ? 'junction' : 'dir')

    await expect(ctx.officeDocument.writeDraft(scope, DOCUMENT_ID, 'roundtrip.xlsx', VERSION, WORKBOOK, signal))
      .rejects.toMatchObject({ code: 'office-document/out-of-scope' })
    await expect(ctx.officeDocument.readDraft(scope, DOCUMENT_ID, signal))
      .rejects.toMatchObject({ code: 'office-document/out-of-scope' })
    await expect(ctx.officeDocument.clearDraft(scope, DOCUMENT_ID, signal))
      .rejects.toMatchObject({ code: 'office-document/out-of-scope' })
    expect(await readdir(outside)).toEqual([])
  } finally {
    await rm(outside, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 })
  }
})

it('reports malformed draft JSON as conversion-failed', async () => {
  await writeDefaultDraft()
  await overwriteDraftFile('not json')
  await expect(ctx.officeDocument.readDraft(scope, DOCUMENT_ID, signal))
    .rejects.toMatchObject({ code: 'office-document/conversion-failed' })
})

const validRecord = (): Record<string, unknown> => ({
  documentId: DOCUMENT_ID,
  path: 'roundtrip.xlsx',
  contentVersion: VERSION,
  snapshot: WORKBOOK,
  savedAt: '2026-11-14T00:00:00.000Z',
})

const MALFORMED_DRAFTS: readonly (readonly [string, string])[] = [
  ['a non-object root', '[]'],
  ['a missing document id', JSON.stringify({ ...validRecord(), documentId: undefined })],
  ['a missing path', JSON.stringify({ ...validRecord(), path: undefined })],
  ['a missing content version', JSON.stringify({ ...validRecord(), contentVersion: undefined })],
  ['a missing timestamp', JSON.stringify({ ...validRecord(), savedAt: undefined })],
  ['a non-object snapshot', JSON.stringify({ ...validRecord(), snapshot: 7 })],
  ['an unknown snapshot kind', JSON.stringify({ ...validRecord(), snapshot: { kind: 'slides', data: {} } })],
  ['a missing snapshot payload', JSON.stringify({ ...validRecord(), snapshot: { kind: 'document' } })],
]

it.each(MALFORMED_DRAFTS)('reports %s as conversion-failed', async (_label, content) => {
  await writeDefaultDraft()
  await overwriteDraftFile(content)
  await expect(ctx.officeDocument.readDraft(scope, DOCUMENT_ID, signal))
    .rejects.toMatchObject({ code: 'office-document/conversion-failed' })
})

it('leaves the source document unchanged', async () => {
  await copyFixture('roundtrip.xlsx')
  const before = await readFile(join(root, 'roundtrip.xlsx'))
  await writeDefaultDraft()
  expect(await readFile(join(root, 'roundtrip.xlsx'))).toEqual(before)
})

it('reports a non-file draft entry as conversion-failed', async () => {
  await writeDefaultDraft()
  const entry = join(draftsDirectory(), await draftFileName())
  await rm(entry)
  await mkdir(entry)

  await expect(ctx.officeDocument.readDraft(scope, DOCUMENT_ID, signal))
    .rejects.toMatchObject({ code: 'office-document/conversion-failed' })
  await expect(ctx.officeDocument.clearDraft(scope, DOCUMENT_ID, signal))
    .rejects.toMatchObject({ code: 'office-document/conversion-failed' })
})

it('reports an unusable draft store as conversion-failed', async () => {
  await writeFile(join(root, '.dsh'), 'not a directory')
  await expect(ctx.officeDocument.writeDraft(scope, DOCUMENT_ID, 'roundtrip.xlsx', VERSION, WORKBOOK, signal))
    .rejects.toMatchObject({ code: 'office-document/conversion-failed' })
})
