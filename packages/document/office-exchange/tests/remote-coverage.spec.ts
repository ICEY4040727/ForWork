/** Remote branch coverage: bounds, ZIP failure mapping, stale-write mapping, and preserved classification. */
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Context } from '@forwork/cordis'
import { LocalFileSystem } from '@forwork/fs-local'
import { SessionId } from '@forwork/session/types'
import SessionStore from '@forwork/session'
import type { SessionEvent } from '@forwork/session'
import type { WorkspaceFileScope } from '@forwork/api-workspace-files'
import { LocaleType } from '@univerjs/core'
import type { IWorkbookData } from '@univerjs/core'
import { strToU8, zipSync } from 'fflate'
import { afterEach, expect, it, vi } from 'vitest'
import { LocalOfficeExchange, OfficeDocumentRemote, OfficeExchange, OfficeExchangeError } from '../src/index.ts'
import {
  readUint32,
  writeUint32,
  ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH,
} from '../src/package/zip-format.ts'
import type {
  ExchangeReadResult,
  OfficeDocumentSnapshot,
  OfficeWriteResult,
  UnsupportedFeature,
} from '../src/index.ts'

const signal = new AbortController().signal
const active: { readonly ctx: Context; readonly root: string; readonly scope: WorkspaceFileScope }[] = []

afterEach(async () => {
  vi.restoreAllMocks()
  for (const harness of active.splice(0)) {
    await harness.ctx.fiber.dispose()
    await rm(harness.root, { recursive: true, force: true })
  }
})

type ReadFactory = () => Promise<ExchangeReadResult>

class ConfigurableExchange extends OfficeExchange {
  constructor(
    ctx: Context,
    private readonly options: {
      readonly produce: ReadFactory
      readonly writeBytes?: Uint8Array
      readonly failWrite?: OfficeExchangeError
    },
  ) {
    super(ctx)
  }

  read(): Promise<ExchangeReadResult> {
    return this.options.produce()
  }

  write(): Promise<OfficeWriteResult> {
    if (this.options.failWrite !== undefined) return Promise.reject(this.options.failWrite)
    return Promise.resolve({
      bytes: this.options.writeBytes ?? new Uint8Array([0x50, 0x4b]),
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

interface Harness {
  readonly ctx: Context
  readonly root: string
  readonly scope: WorkspaceFileScope
  readonly events: SessionEvent[]
}

async function harness(options: {
  readonly maxBytes?: number
  readonly produce?: ReadFactory
  readonly writeBytes?: Uint8Array
  readonly failWrite?: OfficeExchangeError
  readonly localExchange?: boolean
} = {}): Promise<Harness> {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-remote-coverage-'))
  const ctx = new Context()
  await ctx.plugin(LocalFileSystem, { cwd: root })
  await ctx.plugin(SessionStore)
  await ctx.plugin(OfficeDocumentRemote, { maxBytes: options.maxBytes ?? 100_000 })
  if (options.localExchange === true) {
    await ctx.plugin(LocalOfficeExchange)
  } else {
    await ctx.plugin(ConfigurableExchange, {
      produce: options.produce ?? (() => Promise.resolve(workbookRead())),
      ...(options.writeBytes === undefined ? {} : { writeBytes: options.writeBytes }),
      ...(options.failWrite === undefined ? {} : { failWrite: options.failWrite }),
    })
  }
  const scope: WorkspaceFileScope = { sessionId: SessionId('remote-coverage'), workspaceRoot: root }
  const session = ctx.sessions.create(scope.sessionId, { meta: { cwd: root } })
  const events: SessionEvent[] = []
  ctx.on('session/event', (subject, event) => {
    if (subject === session) events.push(event)
  })
  const created: Harness = { ctx, root, scope, events }
  active.push(created)
  return created
}

function workbookData(): IWorkbookData {
  return {
    id: 'book',
    name: 'Book',
    appVersion: '1.0.0',
    locale: LocaleType.EN_US,
    styles: {},
    sheets: {},
    sheetOrder: [],
  }
}

function workbookRead(): ExchangeReadResult {
  return {
    format: 'xlsx',
    snapshot: { kind: 'workbook', data: workbookData() },
    unsupported: [],
  }
}

async function versionOf(target: Harness, path: string): Promise<string> {
  const resolved = await target.ctx.fs.resolve(path)
  const info = await target.ctx.fs.stat(resolved)
  if (info === undefined) throw new Error('expected a file version')
  return info.version
}

async function sourceFile(target: Harness, path: string, content: Uint8Array | string): Promise<string> {
  const bytes = typeof content === 'string' ? strToU8(content) : content
  await writeFile(join(target.root, path), bytes)
  return versionOf(target, path)
}

function workbookSnapshot(): OfficeDocumentSnapshot {
  return {
    kind: 'workbook',
    data: {
      id: 'book',
      name: 'Book',
      appVersion: '1.0.0',
      locale: 'enUS',
      styles: {},
      sheets: {},
      sheetOrder: [],
    },
  }
}

function validSource(): Uint8Array {
  return zipSync({ '[Content_Types].xml': strToU8('<Types/>') })
}

function zipWithEntries(count: number): Uint8Array {
  const entries: Record<string, Uint8Array> = {}
  for (let index = 0; index < count; index += 1) {
    entries[`word/media/part${String(index)}.xml`] = new Uint8Array(0)
  }
  return zipSync(entries)
}

/** A tiny archive whose only entry declares a far larger expansion than it carries. */
function zipWithDeclaredExpansion(declared: number): Uint8Array {
  const bytes = zipSync({ 'word/document.xml': strToU8('x') })
  const eocd = bytes.byteLength - ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH
  const central = readUint32(bytes, eocd + 16, 'central offset')
  const local = readUint32(bytes, central + 42, 'local offset')
  writeUint32(bytes, local + 22, declared)
  writeUint32(bytes, central + 24, declared)
  return bytes
}

it('refuses a serialized file above the Remote byte bound', async () => {
  const activeHarness = await harness({ maxBytes: 1_000, writeBytes: new Uint8Array(2_000) })
  const version = await sourceFile(activeHarness, 'budget.xlsx', validSource())

  await expect(activeHarness.ctx.officeDocument.save(
    activeHarness.scope, 'budget.xlsx', workbookSnapshot(), 'xlsx', version, { actor: 'user' }, signal,
  )).rejects.toMatchObject({ code: 'office-document/bound-exceeded' })
})

it('maps a malformed source package to a conversion failure', async () => {
  const activeHarness = await harness()
  const version = await sourceFile(activeHarness, 'budget.xlsx', 'not a zip archive')

  await expect(activeHarness.ctx.officeDocument.save(
    activeHarness.scope, 'budget.xlsx', workbookSnapshot(), 'xlsx', version, { actor: 'user' }, signal,
  )).rejects.toMatchObject({ code: 'office-document/conversion-failed' })
})

it('maps a source package over the entry bound to a bound failure', async () => {
  const activeHarness = await harness({ maxBytes: 5_000_000 })
  const version = await sourceFile(activeHarness, 'budget.xlsx', zipWithEntries(10_001))

  await expect(activeHarness.ctx.officeDocument.save(
    activeHarness.scope, 'budget.xlsx', workbookSnapshot(), 'xlsx', version, { actor: 'user' }, signal,
  )).rejects.toMatchObject({ code: 'office-document/bound-exceeded' })
})

it('maps a source package over the declared-expansion bound to a bound failure', async () => {
  const activeHarness = await harness()
  const version = await sourceFile(activeHarness, 'budget.xlsx', zipWithDeclaredExpansion(300 * 1024 * 1024))

  await expect(activeHarness.ctx.officeDocument.save(
    activeHarness.scope, 'budget.xlsx', workbookSnapshot(), 'xlsx', version, { actor: 'user' }, signal,
  )).rejects.toMatchObject({ code: 'office-document/bound-exceeded' })
})

it('maps a stale write rejection to a conflict and records no save fact', async () => {
  const activeHarness = await harness()
  const version = await sourceFile(activeHarness, 'budget.xlsx', validSource())
  vi.spyOn(activeHarness.ctx.fs, 'writeBytes').mockRejectedValue(
    Object.assign(new Error('stale'), { code: 'FS_STALE_VERSION' }),
  )

  await expect(activeHarness.ctx.officeDocument.save(
    activeHarness.scope, 'budget.xlsx', workbookSnapshot(), 'xlsx', version, { actor: 'user' }, signal,
  )).rejects.toMatchObject({ code: 'office-document/conflict' })
  expect(activeHarness.events.filter(event => event.type === 'document/save')).toHaveLength(0)
})

it('records a save fact after a successful write', async () => {
  const activeHarness = await harness()
  const version = await sourceFile(activeHarness, 'budget.xlsx', validSource())

  await activeHarness.ctx.officeDocument.save(
    activeHarness.scope, 'budget.xlsx', workbookSnapshot(), 'xlsx', version, { actor: 'user' }, signal,
  )
  expect(activeHarness.events.filter(event => event.type === 'document/save')).toHaveLength(1)
})

it('refuses bytes a read retained past the configured bound', async () => {
  const activeHarness = await harness({ maxBytes: 10 })
  await sourceFile(activeHarness, 'budget.xlsx', validSource())
  vi.spyOn(activeHarness.ctx.fs, 'readBytes').mockImplementation(
    async () => new Uint8Array(20),
  )

  await expect(activeHarness.ctx.officeDocument.read(activeHarness.scope, 'budget.xlsx', signal))
    .rejects.toMatchObject({ code: 'office-document/bound-exceeded' })
})

it('maps every Exchange refusal code onto the Remote vocabulary', async () => {
  const cases = [
    ['unsupported-feature', 'office-document/unsupported-feature'],
    ['out-of-scope', 'office-document/out-of-scope'],
    ['too-large', 'office-document/bound-exceeded'],
    ['not-implemented', 'office-document/conversion-failed'],
  ] as const
  for (const [exchangeCode, remoteCode] of cases) {
    const activeHarness = await harness({
      failWrite: new OfficeExchangeError(exchangeCode, 'refused'),
    })
    const version = await sourceFile(activeHarness, `refusal-${exchangeCode}.xlsx`, validSource())
    await expect(activeHarness.ctx.officeDocument.save(
      activeHarness.scope, `refusal-${exchangeCode}.xlsx`, workbookSnapshot(), 'xlsx', version,
      { actor: 'user' }, signal,
    )).rejects.toMatchObject({ code: remoteCode })
  }
})

it('reports the source file format through a local exchange read', async () => {
  const activeHarness = await harness({ localExchange: true })
  const source = new URL('./fixtures/roundtrip.xlsx', import.meta.url)
  const bytes = new Uint8Array(await readFile(source))
  await writeFile(join(activeHarness.root, 'roundtrip.xlsx'), bytes)
  const read = await activeHarness.ctx.officeDocument.read(activeHarness.scope, 'roundtrip.xlsx', signal)
  expect(read.format).toBe('xlsx')

})
