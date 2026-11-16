/**
 * Edge paths of the `officeDocument` Remote: read races, rebind and missing
 * identity fallbacks, composition gaps, and the error vocabulary boundary.
 */
import { copyFile, mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { LocaleType } from '@univerjs/core'
import type { IWorkbookData } from '@univerjs/core'
import type { WorkspaceFileScope } from '@forwork/api-workspace-files'
import { Context } from '@forwork/cordis'
import { LocalFileSystem } from '@forwork/fs-local'
import SessionStore, { SessionId } from '@forwork/session'
import type { SessionEvent } from '@forwork/session'
import { documentIdFor } from '@forwork/session-document'
import type { JsonValue } from '@forwork/util-values'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { LocalOfficeExchange, OfficeDocumentRemote, OfficeExchange, OfficeExchangeError } from '../src/index.ts'
import type {
  ExchangeReadResult,
  OfficeFormat,
  UnitSnapshot,
  UnsupportedFeature,
} from '../src/index.ts'

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url))
const signal = new AbortController().signal

interface Harness {
  readonly ctx: Context
  readonly root: string
  readonly scope: WorkspaceFileScope
}

const active: Harness[] = []

afterEach(async () => {
  vi.restoreAllMocks()
  for (const harness of active.splice(0)) {
    await harness.ctx.fiber.dispose()
    await rm(harness.root, { recursive: true, force: true })
  }
})

async function rootDirectory(): Promise<string> {
  return await mkdtemp(join(tmpdir(), 'dsh-office-remote-edges-'))
}

function track(ctx: Context, root: string): Harness {
  const scope: WorkspaceFileScope = { sessionId: SessionId('remote-edges'), workspaceRoot: root }
  ctx.sessions.create(scope.sessionId, { meta: { cwd: root } })
  const harness: Harness = { ctx, root, scope }
  active.push(harness)
  return harness
}

/** Local filesystem, Session store, and the local Exchange provider with its Remote. */
async function localHarness(options: { maxFileBytes?: number } = {}): Promise<Harness> {
  const root = await rootDirectory()
  const ctx = new Context()
  await ctx.plugin(LocalFileSystem, { cwd: root })
  await ctx.plugin(SessionStore)
  await ctx.plugin(LocalOfficeExchange, options)
  return track(ctx, root)
}

type ReadFactory = () => Promise<ExchangeReadResult>

/** Exchange provider returning one fixed read result and accepting any write. */
class StubExchange extends OfficeExchange {
  constructor(ctx: Context, private readonly config: { produce: ReadFactory }) {
    super(ctx)
  }

  read(): Promise<ExchangeReadResult> {
    return this.config.produce()
  }

  write(): Promise<Uint8Array> {
    return Promise.resolve(new Uint8Array([0x50, 0x4b]))
  }

  detectUnsupported(): Promise<readonly UnsupportedFeature[]> {
    return Promise.resolve([])
  }

  convertLegacy(): Promise<string> {
    return Promise.resolve('')
  }
}

/** Composition with optional filesystem, Exchange, and Remote services for absent-seam paths. */
async function stubHarness(options: {
  fs?: boolean
  exchange?: boolean
  remote?: boolean
  produce?: ReadFactory
  maxBytes?: number
} = {}): Promise<Harness> {
  const root = await rootDirectory()
  const ctx = new Context()
  if (options.fs !== false) await ctx.plugin(LocalFileSystem, { cwd: root })
  await ctx.plugin(SessionStore)
  if (options.remote !== false) {
    await ctx.plugin(OfficeDocumentRemote, { maxBytes: options.maxBytes ?? 100_000_000 })
  }
  if (options.exchange !== false && options.produce !== undefined) {
    await ctx.plugin(StubExchange, { produce: options.produce })
  }
  return track(ctx, root)
}

async function copyFixture(harness: Harness, name: string): Promise<string> {
  await copyFile(join(FIXTURES, name), join(harness.root, name))
  return name
}

function workbookData(): IWorkbookData {
  return { id: 'book', name: 'Book', appVersion: '1.0.0', locale: LocaleType.EN_US, styles: {}, sheets: {}, sheetOrder: [] }
}

function readResult(snapshot: UnitSnapshot = { kind: 'workbook', data: workbookData() }, format: OfficeFormat = 'xlsx'): ExchangeReadResult {
  return { format, snapshot, unsupported: [] }
}

function sessionEvents(harness: Harness): SessionEvent[] {
  const events: SessionEvent[] = []
  const session = harness.ctx.sessions.get(harness.scope.sessionId)
  if (session === undefined) throw new Error('expected a live Session')
  harness.ctx.on('session/event', (subject, event) => {
    if (subject === session) events.push(event)
  })
  return events
}

describe('read races', () => {
  it('refuses a read whose file changed during conversion', async () => {
    const harness = await localHarness()
    const path = await copyFixture(harness, 'roundtrip.xlsx')
    const exchange = harness.ctx.officeExchange
    const convert = exchange.read.bind(exchange)
    vi.spyOn(exchange, 'read').mockImplementation(async (input, cancellation) => {
      const result = await convert(input, cancellation)
      await writeFile(join(harness.root, path), 'changed after the read')
      return result
    })

    await expect(harness.ctx.officeDocument.read(harness.scope, path, signal))
      .rejects.toMatchObject({ code: 'office-document/conversion-failed' })
  })

  it('refuses a read whose file disappeared during conversion', async () => {
    const harness = await localHarness()
    const path = await copyFixture(harness, 'roundtrip.xlsx')
    const exchange = harness.ctx.officeExchange
    const convert = exchange.read.bind(exchange)
    vi.spyOn(exchange, 'read').mockImplementation(async (input, cancellation) => {
      const result = await convert(input, cancellation)
      await rm(join(harness.root, path))
      return result
    })

    await expect(harness.ctx.officeDocument.read(harness.scope, path, signal))
      .rejects.toMatchObject({ code: 'office-document/conversion-failed' })
  })

  it('reports a non-file target as not found', async () => {
    const harness = await localHarness()
    await mkdir(join(harness.root, 'folder'))

    await expect(harness.ctx.officeDocument.read(harness.scope, 'folder', signal))
      .rejects.toMatchObject({ code: 'office-document/not-found' })
  })

  it('reports an aborted read as cancelled', async () => {
    const harness = await localHarness()
    const path = await copyFixture(harness, 'roundtrip.xlsx')
    const controller = new AbortController()
    const reason = new Error('cancelled during conversion')
    vi.spyOn(harness.ctx.officeExchange, 'read').mockImplementation(async () => {
      controller.abort(reason)
      throw reason
    })

    await expect(harness.ctx.officeDocument.read(harness.scope, path, controller.signal))
      .rejects.toMatchObject({ code: 'gateway/cancelled' })
  })
})

describe('error vocabulary', () => {
  it('maps a filesystem size bound', async () => {
    const harness = await localHarness({ maxFileBytes: 1 })
    const path = await copyFixture(harness, 'roundtrip.xlsx')

    await expect(harness.ctx.officeDocument.read(harness.scope, path, signal))
      .rejects.toMatchObject({ code: 'office-document/bound-exceeded' })
  })

  it('maps an unsupported feature refusal', async () => {
    const harness = await localHarness()
    const path = await copyFixture(harness, 'roundtrip.xlsx')
    vi.spyOn(harness.ctx.officeExchange, 'read')
      .mockRejectedValue(new OfficeExchangeError('unsupported-feature', 'macro'))

    await expect(harness.ctx.officeDocument.read(harness.scope, path, signal))
      .rejects.toMatchObject({ code: 'office-document/unsupported-feature' })
  })

  it('maps a reported size refusal', async () => {
    const harness = await localHarness()
    const path = await copyFixture(harness, 'roundtrip.xlsx')
    vi.spyOn(harness.ctx.officeExchange, 'read')
      .mockRejectedValue(new OfficeExchangeError('too-large', 'too big'))

    await expect(harness.ctx.officeDocument.read(harness.scope, path, signal))
      .rejects.toMatchObject({ code: 'office-document/bound-exceeded' })
  })

  for (const code of ['FS_NOT_FOUND', 'FS_NOT_REGULAR_FILE', 'FS_PERMISSION_DENIED', 'FS_SANDBOX_DENIED', 'FS_IO_ERROR']) {
    it(`maps ${code} to a conversion failure`, async () => {
      const harness = await localHarness()
      const path = await copyFixture(harness, 'roundtrip.xlsx')
      vi.spyOn(harness.ctx.fs, 'stat').mockRejectedValue(Object.assign(new Error('fs'), { code }))

      await expect(harness.ctx.officeDocument.read(harness.scope, path, signal))
        .rejects.toMatchObject({ code: 'office-document/conversion-failed' })
    })
  }

  it('keeps an unrelated Error instance unchanged', async () => {
    const harness = await localHarness()
    const path = await copyFixture(harness, 'roundtrip.xlsx')
    const failure = new Error('unrelated')
    vi.spyOn(harness.ctx.fs, 'stat').mockRejectedValue(failure)

    await expect(harness.ctx.officeDocument.read(harness.scope, path, signal)).rejects.toBe(failure)
  })

  it('wraps a non-Error rejection as a conversion failure', async () => {
    const harness = await localHarness()
    const path = await copyFixture(harness, 'roundtrip.xlsx')
    vi.spyOn(harness.ctx.fs, 'stat').mockRejectedValue('not an error')

    await expect(harness.ctx.officeDocument.read(harness.scope, path, signal))
      .rejects.toMatchObject({ code: 'office-document/conversion-failed' })
  })

  it('refuses feature detection outside the workspace root', async () => {
    const harness = await localHarness()

    await expect(harness.ctx.officeDocument.detectUnsupported(harness.scope, '../outside.docx', signal))
      .rejects.toMatchObject({ code: 'office-document/out-of-scope' })
  })
})

describe('document lifecycle edges', () => {
  it('rebinds a document this process never read from its identity path', async () => {
    const harness = await localHarness()
    await copyFixture(harness, 'roundtrip.xlsx')
    await copyFile(join(harness.root, 'roundtrip.xlsx'), join(harness.root, 'renamed.xlsx'))
    const events = sessionEvents(harness)
    const documentId = documentIdFor(harness.scope.sessionId, 'moved-away.xlsx')

    await harness.ctx.officeDocument.rebindDocument(harness.scope, documentId, 'renamed.xlsx', signal)

    expect(events.filter(event => event.type === 'document/rebind').map(event => event.data)).toEqual([{
      documentId,
      fromPath: 'moved-away.xlsx',
      toPath: 'renamed.xlsx',
    }])
  })

  it('ignores a repeat observation of the same path', async () => {
    const harness = await localHarness()
    const path = await copyFixture(harness, 'roundtrip.xlsx')
    await harness.ctx.officeDocument.read(harness.scope, path, signal)
    const events = sessionEvents(harness)

    await harness.ctx.officeDocument.rebindDocument(harness.scope, documentIdFor(harness.scope.sessionId, path), path, signal)

    expect(events.filter(event => event.type === 'document/rebind')).toEqual([])
  })

  it('marks an unread document missing from its identity path', async () => {
    const harness = await localHarness()
    const events = sessionEvents(harness)
    const documentId = documentIdFor(harness.scope.sessionId, 'ghost.xlsx')

    await harness.ctx.officeDocument.markDocumentMissing(harness.scope, documentId, signal)

    expect(events.filter(event => event.type === 'document/missing').map(event => event.data)).toEqual([
      { documentId, path: 'ghost.xlsx' },
    ])
  })

  it('refuses a missing fact for another Session identity', async () => {
    const harness = await localHarness()

    await expect(harness.ctx.officeDocument.markDocumentMissing(harness.scope, 'other-session:x.xlsx', signal))
      .rejects.toMatchObject({ code: 'gateway/bad-request' })
  })

  it('refuses a lifecycle fact when the Session is not live', async () => {
    const harness = await localHarness()
    const scope: WorkspaceFileScope = { sessionId: SessionId('not-live'), workspaceRoot: harness.root }

    await expect(harness.ctx.officeDocument.closeDocument(scope, 'not-live:x.xlsx', 'closed', signal))
      .rejects.toMatchObject({ code: 'gateway/bad-request' })
  })

  it('opens a document again when a save finds no tracked entry', async () => {
    const harness = await localHarness()
    const path = await copyFixture(harness, 'roundtrip.xlsx')
    const read = await harness.ctx.officeDocument.read(harness.scope, path, signal)
    await harness.ctx.officeDocument.closeDocument(harness.scope, documentIdFor(harness.scope.sessionId, path), 'closed', signal)
    const events = sessionEvents(harness)

    await harness.ctx.officeDocument.save(harness.scope, path, read.snapshot, 'xlsx', read.version, signal)

    expect(events.filter(event => event.type === 'document/open')).toHaveLength(1)
    expect(events.filter(event => event.type === 'document/save')).toHaveLength(1)
  })
})

describe('wire snapshot validation', () => {
  async function readWith(produce: ReadFactory): Promise<Harness> {
    const harness = await stubHarness({ produce })
    await writeFile(join(harness.root, 'budget.xlsx'), 'bytes')
    return harness
  }

  function workbookWith(extra: Record<string, unknown>): UnitSnapshot {
    const data = workbookData()
    Object.assign(data, extra)
    return { kind: 'workbook', data }
  }

  it('refuses a non-finite number', async () => {
    const harness = await readWith(async () => readResult(workbookWith({ score: Infinity })))

    await expect(harness.ctx.officeDocument.read(harness.scope, 'budget.xlsx', signal))
      .rejects.toMatchObject({ code: 'office-document/conversion-failed' })
  })

  it('refuses a class instance', async () => {
    const harness = await readWith(async () => readResult(workbookWith({ nested: new Date() })))

    await expect(harness.ctx.officeDocument.read(harness.scope, 'budget.xlsx', signal))
      .rejects.toMatchObject({ code: 'office-document/conversion-failed' })
  })

  it('refuses a function member', async () => {
    const harness = await readWith(async () => readResult(workbookWith({ nested: () => 1 })))

    await expect(harness.ctx.officeDocument.read(harness.scope, 'budget.xlsx', signal))
      .rejects.toMatchObject({ code: 'office-document/conversion-failed' })
  })

  it('drops undefined members from the wire snapshot', async () => {
    const harness = await readWith(async () => readResult(workbookWith({ skipped: undefined, kept: 1 })))

    const read = await harness.ctx.officeDocument.read(harness.scope, 'budget.xlsx', signal)
    const data = read.snapshot.data as Record<string, JsonValue>

    expect(Object.hasOwn(data, 'skipped')).toBe(false)
    expect(data['kept']).toBe(1)
  })

  it('keeps a prototype-free object as a plain JSON object', async () => {
    const harness = await readWith(async () => readResult(workbookWith({ nested: Object.create(null) })))

    const read = await harness.ctx.officeDocument.read(harness.scope, 'budget.xlsx', signal)

    expect((read.snapshot.data as Record<string, JsonValue>)['nested']).toEqual({})
  })

  for (const format of ['pptx', 'doc', 'xls', 'ppt'] as const) {
    it(`refuses a .${format} read result`, async () => {
      const harness = await readWith(async () => readResult({ kind: 'workbook', data: workbookData() }, format))

      await expect(harness.ctx.officeDocument.read(harness.scope, 'budget.xlsx', signal))
        .rejects.toMatchObject({ code: 'office-document/unsupported-format' })
    })
  }
})

describe('save validation', () => {
  it('refuses a document snapshot saved as a workbook', async () => {
    const harness = await localHarness()

    await expect(harness.ctx.officeDocument.save(
      harness.scope, 'budget.xlsx', { kind: 'document', data: {} }, 'xlsx', 'v1', signal,
    )).rejects.toMatchObject({ code: 'office-document/conversion-failed' })
  })

  it('refuses a workbook snapshot saved as a document', async () => {
    const harness = await localHarness()

    await expect(harness.ctx.officeDocument.save(
      harness.scope, 'report.docx', { kind: 'workbook', data: {} }, 'docx', 'v1', signal,
    )).rejects.toMatchObject({ code: 'office-document/conversion-failed' })
  })

  for (const format of ['pptx', 'doc', 'xls', 'ppt'] as const) {
    it(`refuses a .${format} save`, async () => {
      const harness = await localHarness()

      await expect(harness.ctx.officeDocument.save(
        harness.scope, 'budget.xlsx', { kind: 'workbook', data: {} }, format, 'v1', signal,
      )).rejects.toMatchObject({ code: 'office-document/unsupported-format' })
    })
  }

  it('refuses a workbook payload without workbook identity', async () => {
    const harness = await localHarness()

    await expect(harness.ctx.officeDocument.save(
      harness.scope, 'budget.xlsx', { kind: 'workbook', data: [] }, 'xlsx', 'v1', signal,
    )).rejects.toMatchObject({ code: 'office-document/conversion-failed' })
  })

  it('refuses a document payload without document identity', async () => {
    const harness = await localHarness()

    await expect(harness.ctx.officeDocument.save(
      harness.scope, 'report.docx', { kind: 'document', data: [] }, 'docx', 'v1', signal,
    )).rejects.toMatchObject({ code: 'office-document/conversion-failed' })
  })

  it('saves without a command broker in the composition', async () => {
    const harness = await stubHarness({ produce: async () => readResult() })
    const path = 'budget.xlsx'
    await writeFile(join(harness.root, path), 'bytes')
    const read = await harness.ctx.officeDocument.read(harness.scope, path, signal)
    await harness.ctx.officeDocument.closeDocument(harness.scope, documentIdFor(harness.scope.sessionId, path), 'closed', signal)

    await expect(harness.ctx.officeDocument.save(harness.scope, path, read.snapshot, 'xlsx', read.version, signal))
      .resolves.toMatchObject({ version: expect.any(String) })
  })
})

describe('absent composition seams', () => {
  it('reports an absent Exchange seam', async () => {
    const harness = await stubHarness()
    await writeFile(join(harness.root, 'budget.xlsx'), 'bytes')

    await expect(harness.ctx.officeDocument.read(harness.scope, 'budget.xlsx', signal))
      .rejects.toMatchObject({ code: 'gateway/internal' })
  })

  it('reports an absent filesystem seam', async () => {
    const harness = await stubHarness({ fs: false, produce: async () => readResult() })

    await expect(harness.ctx.officeDocument.read(harness.scope, 'budget.xlsx', signal))
      .rejects.toMatchObject({ code: 'gateway/internal' })
  })
})
