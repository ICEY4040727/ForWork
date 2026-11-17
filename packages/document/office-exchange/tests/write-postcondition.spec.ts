/**
 * Contract: the Host save boundary refuses a source whose unsupported set is
 * non-empty. `office-exchange.contract.md` requires `write` never to drop
 * content it does not explicitly change, so a save that skips the check is
 * non-conforming even though the client and tools already refuse earlier.
 */
import { createHash } from 'node:crypto'
import { copyFile, mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Context } from '@forwork/cordis'
import { LocalFileSystem } from '@forwork/fs-local'
import SessionStore, { SessionId } from '@forwork/session'
import type { SessionEvent } from '@forwork/session'
import { documentIdFor, foldDocumentProjection } from '@forwork/session-document'
import type { DocumentProjectionState } from '@forwork/session-document'
import type { JsonValue } from '@forwork/util-values'
import { afterEach, beforeEach, expect, it } from 'vitest'
import LocalOfficeExchange, { type OfficeDocumentSnapshot } from '../src/index.ts'

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url))
const scope = { sessionId: SessionId('office-write-postcondition'), workspaceRoot: '' }
const signal = new AbortController().signal

let ctx: Context
let root: string
const events: SessionEvent[] = []

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'dsh-office-write-postcondition-'))
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

function hashOf(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

function jsonObject(value: JsonValue | undefined): Record<string, JsonValue> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) throw new Error('expected a JSON object')
  return value
}

function documentStream(snapshot: OfficeDocumentSnapshot): string {
  if (snapshot.kind !== 'document') throw new Error('expected a document snapshot')
  const body = jsonObject(jsonObject(snapshot.data)['body'])
  const stream = body['dataStream']
  if (typeof stream !== 'string') throw new Error('expected a document data stream')
  return stream
}

function openedDocuments(): DocumentProjectionState {
  let state: DocumentProjectionState = { order: [], documents: {} }
  for (const event of events) state = foldDocumentProjection(state, event)
  return state
}

it('refuses a save whose source still reports unsupported content', async () => {
  const path = await copyFixture('unsupported-macro.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  expect(read.unsupported.map(feature => feature.kind)).toContain('macro')
  const before = await readFile(join(root, path))

  await expect(ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, signal))
    .rejects.toMatchObject({
      code: 'office-document/unsupported-feature',
      message: expect.stringContaining('macro'),
    })

  expect(hashOf(await readFile(join(root, path)))).toBe(hashOf(before))
  expect(openedDocuments().documents[documentIdFor(scope.sessionId, path)]?.contentVersion).toBe(0)
})

for (const [fixture, kind] of [
  ['unsupported-tracked.docx', 'tracked-changes'],
  ['unsupported-header.docx', 'header-footer'],
] as const) {
  it(`refuses a save that would drop ${kind}`, async () => {
    const path = await copyFixture(fixture)
    const read = await ctx.officeDocument.read(scope, path, signal)
    expect(read.unsupported.map(feature => feature.kind)).toContain(kind)
    const before = await readFile(join(root, path))

    await expect(ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, signal))
      .rejects.toMatchObject({
        code: 'office-document/unsupported-feature',
        message: expect.stringContaining(kind),
      })

    expect(hashOf(await readFile(join(root, path)))).toBe(hashOf(before))
    expect(openedDocuments().documents[documentIdFor(scope.sessionId, path)]?.contentVersion).toBe(0)
  })
}

it('appends no document/save for a refused save', async () => {
  const path = await copyFixture('unsupported-macro.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)

  await expect(ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, signal))
    .rejects.toMatchObject({ code: 'office-document/unsupported-feature' })

  expect(events.filter(event => event.type === 'document/open')).toHaveLength(1)
  expect(events.filter(event => event.type === 'document/save')).toHaveLength(0)
})

it('saves a supported source and advances the file and content versions', async () => {
  const path = await copyFixture('roundtrip.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  expect(read.unsupported).toEqual([])
  const before = await readFile(join(root, path))
  if (read.snapshot.kind !== 'document') throw new Error('expected a document snapshot')
  const body = jsonObject(jsonObject(read.snapshot.data)['body'])
  body['dataStream'] = documentStream(read.snapshot).replace('Quarterly', 'Financial')

  const saved = await ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, signal)

  expect(saved.version).not.toBe(read.version)
  expect(hashOf(await readFile(join(root, path)))).not.toBe(hashOf(before))
  const reread = await ctx.officeDocument.read(scope, path, signal)
  expect(documentStream(reread.snapshot)).toContain('Financial Report')
  expect(openedDocuments().documents[documentIdFor(scope.sessionId, path)]?.contentVersion).toBe(1)
  expect(events.filter(event => event.type === 'document/save')).toHaveLength(1)
})
