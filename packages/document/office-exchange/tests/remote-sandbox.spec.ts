/**
 * Per-call sandbox policy on the `officeDocument` Remote: every mutating
 * filesystem call carries the Session's policy, a composition without a policy
 * service still works, and a real sandboxed filesystem saves inside a workspace
 * outside the Host cwd.
 */
import { mkdtemp, realpath, rm, writeFile } from 'node:fs/promises'
import { homedir, tmpdir } from 'node:os'
import { join } from 'node:path'
import { LocaleType } from '@univerjs/core'
import type { IWorkbookData } from '@univerjs/core'
import type { WorkspaceFileScope } from '@forwork/api-workspace-files'
import { Context } from '@forwork/cordis'
import { LocalFileSystem } from '@forwork/fs-local'
import { SandboxedFileSystem } from '@forwork/fs-sandbox'
import SandboxPolicyService, { setSandboxMode } from '@forwork/sandbox-policy'
import type { Session } from '@forwork/session'
import SessionStore, { SessionId } from '@forwork/session'
import SessionProjectionRegistry from '@forwork/session-projection'
import { strToU8, zipSync } from 'fflate'
import { afterEach, expect, it, vi } from 'vitest'
import { OfficeDocumentRemote, OfficeExchange } from '../src/index.ts'
import type {
  ExchangeReadResult,
  OfficeDocumentSnapshot,
  OfficeWriteResult,
  UnsupportedFeature,
} from '../src/index.ts'

const SESSION_ID = SessionId('remote-sandbox')
const signal = new AbortController().signal

interface Harness {
  readonly ctx: Context
  readonly root: string
  readonly scope: WorkspaceFileScope
  readonly session: Session | undefined
}

const active: Harness[] = []

afterEach(async () => {
  vi.restoreAllMocks()
  for (const harness of active.splice(0)) {
    await harness.ctx.fiber.dispose()
    await rm(harness.root, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 })
  }
})

/** Exchange provider accepting every write so only the filesystem path is under test. */
class StubExchange extends OfficeExchange {
  read(): Promise<ExchangeReadResult> {
    return Promise.resolve({ format: 'xlsx', snapshot: { kind: 'workbook', data: workbookData() }, unsupported: [] })
  }

  write(): Promise<OfficeWriteResult> {
    return Promise.resolve({
      bytes: new Uint8Array([0x50, 0x4b]),
      disposition: {
        schemaVersion: 1,
        verbatimCount: 0,
        spliced: { count: 0, names: [] },
        added: { count: 0, names: [] },
        removed: { count: 0, names: [] },
        nameLimit: 256,
      },
    })
  }

  detectUnsupported(): Promise<readonly UnsupportedFeature[]> {
    return Promise.resolve([])
  }

  convertLegacy(): Promise<string> {
    return Promise.resolve('')
  }
}

function workbookData(): IWorkbookData {
  return { id: 'book', name: 'Book', appVersion: '1.0.0', locale: LocaleType.EN_US, styles: {}, sheets: {}, sheetOrder: [] }
}

function workbookSnapshot(): OfficeDocumentSnapshot {
  return {
    kind: 'workbook',
    data: { id: 'book', name: 'Book', appVersion: '1.0.0', locale: 'enUS', styles: {}, sheets: {}, sheetOrder: [] },
  }
}

/** A minimal source archive the guarded save can index before replacing it. */
function validSource(): Uint8Array {
  return zipSync({ '[Content_Types].xml': strToU8('<Types/>') })
}

function track(ctx: Context, root: string, session: Session | undefined): Harness {
  const scope: WorkspaceFileScope = { sessionId: SESSION_ID, workspaceRoot: root }
  const harness: Harness = { ctx, root, scope, session }
  active.push(harness)
  return harness
}

/**
 * Local filesystem composition whose policy mode comes from the live Session
 * override rather than the deployment default.
 * @param policy - `live` mounts the real policy service with a live Session,
 *   `persisted` mounts it without one, `absent` mounts no policy service.
 */
async function localHarness(policy: 'live' | 'persisted' | 'absent' = 'live'): Promise<Harness> {
  const root = await realpath(await mkdtemp(join(tmpdir(), 'dsh-office-remote-sandbox-')))
  const ctx = new Context()
  await ctx.plugin(LocalFileSystem, { cwd: root })
  await ctx.plugin(SessionStore)
  if (policy !== 'absent') {
    await ctx.plugin(SessionProjectionRegistry)
    await ctx.plugin(SandboxPolicyService, { mode: 'read-only', workspaceRoot: process.cwd() })
  }
  const session = policy === 'persisted' ? undefined : ctx.sessions.create(SESSION_ID, { meta: { cwd: root } })
  if (session !== undefined) setSandboxMode(session, 'workspace-write')
  await ctx.plugin(OfficeDocumentRemote, { maxBytes: 100_000_000 })
  await ctx.plugin(StubExchange)
  return track(ctx, root, session)
}

/** Real sandbox enforcement over a workspace outside both the Host cwd and the temp areas. */
async function sandboxHarness(): Promise<Harness> {
  const root = await realpath(await mkdtemp(join(homedir(), '.dsh-office-remote-sandbox-')))
  const ctx = new Context()
  await ctx.plugin(SessionProjectionRegistry)
  await ctx.plugin(SandboxPolicyService, { mode: 'read-only', workspaceRoot: process.cwd() })
  await ctx.plugin(SandboxedFileSystem, { cwd: root })
  await ctx.plugin(SessionStore)
  const session = ctx.sessions.create(SESSION_ID, { meta: { cwd: root } })
  setSandboxMode(session, 'workspace-write')
  await ctx.plugin(OfficeDocumentRemote, { maxBytes: 100_000_000 })
  await ctx.plugin(StubExchange)
  return track(ctx, root, session)
}

/** The version of a workspace file the Remote must read at. */
async function versionOf(harness: Harness, path: string): Promise<string> {
  const resolved = await harness.ctx.fs.resolve(path, { cwd: harness.root })
  const info = await harness.ctx.fs.stat(resolved)
  if (info === undefined) throw new Error(`expected "${path}" to exist`)
  return info.version
}

async function saveSource(harness: Harness, path: string): Promise<string> {
  await writeFile(join(harness.root, path), validSource())
  return await versionOf(harness, path)
}

it('passes the live Session policy to the save write', async () => {
  const harness = await localHarness()
  const version = await saveSource(harness, 'budget.xlsx')
  const write = vi.spyOn(harness.ctx.fs, 'writeBytes')

  await harness.ctx.officeDocument.save(
    harness.scope, 'budget.xlsx', workbookSnapshot(), 'xlsx', version, { actor: 'user' }, signal,
  )

  const policy = write.mock.calls[0]?.[4]
  const session = harness.session
  if (session === undefined) throw new Error('expected a live session')
  expect(policy).toEqual(harness.ctx.sandboxPolicy.resolve({ session }))
  expect(policy?.workspaceRoot).toBe(harness.scope.workspaceRoot)
  expect(policy?.mode).toBe('workspace-write')
})

it('passes the live Session policy to the draft write and clear', async () => {
  const harness = await localHarness()
  const write = vi.spyOn(harness.ctx.fs, 'writeText')
  const remove = vi.spyOn(harness.ctx.fs, 'deleteFile')

  await harness.ctx.officeDocument.writeDraft(
    harness.scope, 'draft-1', 'budget.xlsx', 'version-1', workbookSnapshot(), signal,
  )
  await harness.ctx.officeDocument.clearDraft(harness.scope, 'draft-1', signal)

  const session = harness.session
  if (session === undefined) throw new Error('expected a live session')
  const expected = harness.ctx.sandboxPolicy.resolve({ session })
  const writePolicy = write.mock.calls[0]?.[4]
  const deletePolicy = remove.mock.calls[0]?.[3]
  expect(writePolicy).toEqual(expected)
  expect(deletePolicy).toEqual(expected)
  expect(writePolicy?.workspaceRoot).toBe(harness.scope.workspaceRoot)
  expect(deletePolicy?.workspaceRoot).toBe(harness.scope.workspaceRoot)
  expect(writePolicy?.mode).toBe('workspace-write')
  expect(deletePolicy?.mode).toBe('workspace-write')
})

it('leaves the policy undefined in a composition without sandboxPolicy', async () => {
  const harness = await localHarness('absent')
  const write = vi.spyOn(harness.ctx.fs, 'writeBytes')
  const writeText = vi.spyOn(harness.ctx.fs, 'writeText')
  const remove = vi.spyOn(harness.ctx.fs, 'deleteFile')
  const version = await saveSource(harness, 'budget.xlsx')

  await harness.ctx.officeDocument.save(
    harness.scope, 'budget.xlsx', workbookSnapshot(), 'xlsx', version, { actor: 'user' }, signal,
  )
  await harness.ctx.officeDocument.writeDraft(
    harness.scope, 'draft-1', 'budget.xlsx', version, workbookSnapshot(), signal,
  )
  await harness.ctx.officeDocument.clearDraft(harness.scope, 'draft-1', signal)

  expect(write.mock.calls[0]?.[4]).toBeUndefined()
  expect(writeText.mock.calls[0]?.[4]).toBeUndefined()
  expect(remove.mock.calls[0]?.[3]).toBeUndefined()
})

it('fences a persisted Session draft to the wire root under the default mode', async () => {
  const harness = await localHarness('persisted')
  const write = vi.spyOn(harness.ctx.fs, 'writeText')

  await harness.ctx.officeDocument.writeDraft(
    harness.scope, 'draft-1', 'budget.xlsx', 'version-1', workbookSnapshot(), signal,
  )

  expect(write.mock.calls[0]?.[4]).toEqual({ mode: 'read-only', workspaceRoot: harness.scope.workspaceRoot })
})

it('saves and stores drafts outside the Host cwd through fs-sandbox', async () => {
  const harness = await sandboxHarness()
  const version = await saveSource(harness, 'budget.xlsx')
  const write = vi.spyOn(harness.ctx.fs, 'writeBytes')

  const saved = await harness.ctx.officeDocument.save(
    harness.scope, 'budget.xlsx', workbookSnapshot(), 'xlsx', version, { actor: 'user' }, signal,
  )
  expect(saved.version).not.toBe(version)
  const savePolicy = write.mock.calls[0]?.[4]
  expect(savePolicy?.workspaceRoot).toBe(harness.scope.workspaceRoot)
  expect(savePolicy?.mode).toBe('workspace-write')

  await harness.ctx.officeDocument.writeDraft(
    harness.scope, 'draft-1', 'budget.xlsx', saved.version, workbookSnapshot(), signal,
  )
  expect(await harness.ctx.officeDocument.readDraft(harness.scope, 'draft-1', signal)).not.toBeNull()
  await harness.ctx.officeDocument.clearDraft(harness.scope, 'draft-1', signal)
  expect(await harness.ctx.officeDocument.readDraft(harness.scope, 'draft-1', signal)).toBeNull()
})
