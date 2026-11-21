// @vitest-environment jsdom
/**
 * The save that follows an accepted agent edit keeps the live unit (FR-014,
 * SC-012).
 *
 * The tab's real command bridge applies the edit through the mounted adapter,
 * and the Host's save is simulated by advancing the file version the resource
 * reports and the numeric version the `document` projection publishes. The
 * classification under test is the hook's own: a save this tab caused adopts the
 * published version in place, so the handle is never disposed and its native
 * history still reverses the agent edit. A change no applied command accounts
 * for still reloads.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { Mock } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import type { ReactElement } from 'react'
import { brandString } from '@forwork/brand'
import type { DocumentId } from '@forwork/session/src/document-events.ts'
import type { SessionId } from '@forwork/session/types'
import type { RemoteResult } from '@forwork/typert-protocol'
import type { WorkspaceFileStat } from '@forwork/api-workspace-files/types'
import type { DocumentProjectionView } from '@forwork/session-document/client'
import type {
  OfficeDocumentReadResult,
  OfficeDocumentSnapshot,
} from '@forwork/office-exchange/types'
import type { OfficePanelProps } from '../src/client/OfficePanel.tsx'
import type { OfficePanelInjected } from '../src/client/face.ts'
import { OfficePanel } from '../src/client/OfficePanel.tsx'
import type { OfficeEditorHandle, OfficeRuntime } from '../src/client/adapter.ts'
import { createAdapter } from '../src/client/adapter.ts'
import type {
  CommandEventChange,
  CommandEventEntry,
  CommandEventFeed,
  DocumentCommandRequest,
  OfficeCommandRemote,
} from '../src/client/command-bridge.ts'
import { attachSessionCommandBridge } from '../src/client/command-bridge.ts'
import { en } from '../src/client/locales.ts'
import { makeDraftRemote } from './draft-fixtures.client.ts'
import type { DraftRemote } from './draft-fixtures.client.ts'
import { SESSION_ID, panelFixture, stubMeasuredSurface } from './panel.tsx'

beforeEach(stubMeasuredSurface)

afterEach(cleanup)

const PATH = 'work/budget.xlsx'
const DOCUMENT_ID = brandString<DocumentId>(`${SESSION_ID}:${PATH}`)
const ADDRESS = `dsh-resource://file/session/s-1/${PATH}`
const STEPS = [{ kind: 'cell' as const, sheetId: 'sheet-1', row: 0, column: 0, value: 'after' }]

const BEFORE: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book', value: 'before' } }
const AFTER: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book', value: 'after' } }

/** The one Remote method the Host answer travels through. */
type CompleteCommand = NonNullable<OfficeCommandRemote['completeCommand']>

function success<T>(value: T): RemoteResult<T> {
  return { ok: true, value }
}

/** A controllable Session event feed, the same boundary the controller publishes to a bridge. */
class FakeFeed implements CommandEventFeed {
  private change: CommandEventChange = { kind: 'replace' }
  private readonly listeners = new Set<() => void>()

  getSnapshot(): { change: CommandEventChange } {
    return { change: this.change }
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener)
    return () => { this.listeners.delete(listener) }
  }

  appendCommand(request: DocumentCommandRequest): void {
    this.change = {
      kind: 'append',
      entries: [{ type: 'event', event: { type: 'document/command', data: request } } satisfies CommandEventEntry],
    }
    for (const listener of [...this.listeners]) listener()
  }
}

/**
 * A runtime whose history is a linear snapshot list the test drives through the
 * adapter handle.
 * @param start - the snapshot the unit opened from.
 * @returns the runtime behind the handle.
 */
function historyRuntime(start: OfficeDocumentSnapshot): OfficeRuntime {
  const states: OfficeDocumentSnapshot[] = [start]
  let index = 0
  const listeners = new Set<() => void>()
  const notify = (): void => { for (const listener of [...listeners]) listener() }
  return {
    getSnapshot: () => states[index] ?? start,
    apply() {
      states.length = index + 1
      states.push(AFTER)
      index = states.length - 1
      notify()
      return { status: 'applied' }
    },
    undo() {
      if (index === 0) return
      index -= 1
      notify()
    },
    redo() {
      if (index >= states.length - 1) return
      index += 1
      notify()
    },
    onMutation(listener) {
      listeners.add(listener)
      return () => { listeners.delete(listener) }
    },
    dispose() { listeners.clear() },
  }
}

/** The `document` projection entry a save or edit fact leaves for the open document. */
function projectionAt(contentVersion: number): DocumentProjectionView {
  return {
    order: [DOCUMENT_ID],
    documents: {
      [DOCUMENT_ID]: {
        documentId: DOCUMENT_ID,
        path: PATH,
        format: 'xlsx',
        contentVersion,
        status: 'open',
        unsupported: [],
      },
    },
  }
}

interface Live {
  meta: WorkspaceFileStat | undefined
  projection: DocumentProjectionView | undefined
  /** The version the next read reports. */
  readVersion: string
  /** A fresh object per read, so a reload always remounts the unit. */
  reads: number
}

interface Harness {
  readonly live: Live
  readonly feed: FakeFeed
  readonly remote: DraftRemote
  readonly handles: OfficeEditorHandle[]
  readonly completeCommand: Mock<CompleteCommand>
  readonly props: () => OfficePanelProps
}

/**
 * Build the panel fixture whose file version and projection the test advances,
 * with the production command bridge attached to the mounted target.
 * @returns the mutable live values, the bridge feed, and the props builder.
 */
function harness(): Harness {
  const live: Live = { meta: { absolutePath: `/${PATH}`, version: 'v1' }, projection: undefined, readVersion: 'v1', reads: 0 }
  const remote = makeDraftRemote()
  remote.read.mockImplementation(async (): Promise<RemoteResult<OfficeDocumentReadResult>> => {
    live.reads += 1
    return success({
      format: 'xlsx',
      snapshot: live.reads === 1 ? BEFORE : { kind: 'workbook', data: { id: 'book', reread: live.reads } },
      unsupported: [],
      version: live.readVersion,
    })
  })
  const feed = new FakeFeed()
  const completeCommand = vi.fn<CompleteCommand>(async () => undefined)
  const commandRemote: OfficeCommandRemote = {
    read: () => undefined,
    registerCommandTarget: vi.fn(async () => undefined),
    unregisterCommandTarget: vi.fn(async () => undefined),
    completeCommand,
  }
  const handles: OfficeEditorHandle[] = []
  const adapter = createAdapter(async mount => historyRuntime(mount.snapshot))
  const mountEditor: OfficePanelInjected['mountEditor'] = (container, input) =>
    adapter.mount(container, input).then((handle) => {
      handles.push(handle)
      return handle
    })
  const fixture = panelFixture({
    address: ADDRESS,
    remote: remote.face,
    mount: mountEditor,
    attachCommandBridge: (sessionId: SessionId, target) =>
      attachSessionCommandBridge({ eventSource: feed }, sessionId, target, commandRemote),
  })
  const props = (): OfficePanelProps => ({
    ...fixture.props,
    useResource: () => live.meta === undefined
      ? { status: 'loading', value: undefined, failure: undefined }
      : { status: 'live', value: live.meta, failure: undefined },
    useProjection: (key: string) => key === 'document' ? live.projection : undefined,
  }) as OfficePanelProps
  return { live, feed, remote, handles, completeCommand, props }
}

/**
 * Apply one edit through the attached bridge and return the version it reported.
 * @param harness - the panel fixture under test.
 * @param view - the rendered panel, so React state updates flush inside `act`.
 * @returns the content version the tab reported to the Host.
 */
async function applyAgentEdit(
  harness: Harness,
  view: { rerender: (element: ReactElement) => void },
): Promise<number> {
  act(() => {
    harness.feed.appendCommand({
      requestId: 'request-1',
      documentId: DOCUMENT_ID,
      operation: { kind: 'edit', steps: STEPS },
      fromVersion: 0,
    })
  })
  await waitFor(() => { expect(harness.completeCommand).toHaveBeenCalledOnce() })
  const result = harness.completeCommand.mock.calls[0]?.[2]
  if (result?.status !== 'applied') throw new Error('the bridge did not report an applied edit')
  expect(result.snapshot).toEqual(AFTER)
  expect(harness.handles).toHaveLength(1)
  view.rerender(<OfficePanel {...harness.props()} />)
  return result.contentVersion
}

describe('a local save', () => {
  it.each(['the watched file', 'the document projection'] as const)(
    'keeps one live unit when %s reports the save first',
    async (first) => {
      const h = harness()
      // The panel reads the file before the resource's first frame, so the read
      // has no observed baseline to compare a later version against.
      h.live.meta = undefined
      const view = render(<OfficePanel {...h.props()} />)
      await waitFor(() => { expect(h.handles).toHaveLength(1) })
      const mounted = h.handles[0]

      h.live.meta = { absolutePath: `/${PATH}`, version: 'v1' }
      view.rerender(<OfficePanel {...h.props()} />)
      expect(h.remote.read).toHaveBeenCalledOnce()

      fireEvent.click(view.getByText(en.save))
      await waitFor(() => { expect(h.remote.save).toHaveBeenCalledOnce() })

      const updateFile = (): void => { h.live.meta = { absolutePath: `/${PATH}`, version: 'v2' } }
      const updateProjection = (): void => { h.live.projection = projectionAt(1) }
      const updates = first === 'the watched file' ? [updateFile, updateProjection] : [updateProjection, updateFile]
      for (const update of updates) {
        update()
        view.rerender(<OfficePanel {...h.props()} />)
        expect(h.remote.read).toHaveBeenCalledOnce()
        expect(h.handles).toHaveLength(1)
        expect(h.handles[0]).toBe(mounted)
      }
    },
  )
  it('leaves the unit mounted when another save advanced the projection first', async () => {
    const h = harness()
    h.live.meta = undefined
    const view = render(<OfficePanel {...h.props()} />)
    await waitFor(() => { expect(h.handles).toHaveLength(1) })
    const mounted = h.handles[0]

    h.live.meta = { absolutePath: `/${PATH}`, version: 'v1' }
    view.rerender(<OfficePanel {...h.props()} />)
    fireEvent.click(view.getByText(en.save))
    await waitFor(() => { expect(h.remote.clearDraft).toHaveBeenCalledOnce() })

    // Another writer's save advanced the projection while the watched file
    // still reports the version this tab replaced.
    h.live.projection = projectionAt(2)
    view.rerender(<OfficePanel {...h.props()} />)
    expect(h.remote.read).toHaveBeenCalledOnce()
    expect(h.handles).toHaveLength(1)
    expect(h.handles[0]).toBe(mounted)
  })

  it('keeps the unit mounted when the file resource loses its value after a save', async () => {
    const h = harness()
    const view = render(<OfficePanel {...h.props()} />)
    await waitFor(() => { expect(h.handles).toHaveLength(1) })

    fireEvent.click(view.getByText(en.save))
    await waitFor(() => { expect(h.remote.clearDraft).toHaveBeenCalledOnce() })

    h.live.meta = undefined
    view.rerender(<OfficePanel {...h.props()} />)
    expect(h.remote.read).toHaveBeenCalledOnce()
    expect(h.handles).toHaveLength(1)
  })

  it('reloads a clean unit when a different file version follows the save', async () => {
    const h = harness()
    const view = render(<OfficePanel {...h.props()} />)
    await waitFor(() => { expect(h.handles).toHaveLength(1) })

    fireEvent.click(view.getByText(en.save))
    await waitFor(() => { expect(h.remote.clearDraft).toHaveBeenCalledOnce() })

    h.live.readVersion = 'v3'
    h.live.meta = { absolutePath: `/${PATH}`, version: 'v3' }
    view.rerender(<OfficePanel {...h.props()} />)
    await waitFor(() => { expect(h.remote.read).toHaveBeenCalledTimes(2) })
    await waitFor(() => { expect(h.handles).toHaveLength(2) })
  })
})

describe('an accepted agent edit', () => {
  it('keeps the mounted unit and its native history across the Host save that follows it', async () => {
    const h = harness()
    const view = render(<OfficePanel {...h.props()} />)
    await waitFor(() => { expect(h.handles).toHaveLength(1) })
    const mounted = h.handles[0]

    const contentVersion = await applyAgentEdit(h, view)
    expect(contentVersion).toBe(1)
    const readBeforeSave = h.remote.read.mock.calls.length

    // The Host saved the returned snapshot: the watched file moves first.
    h.live.meta = { absolutePath: `/${PATH}`, version: 'v2' }
    view.rerender(<OfficePanel {...h.props()} />)
    expect(h.remote.read).toHaveBeenCalledTimes(readBeforeSave)
    expect(h.handles).toHaveLength(1)

    // The projection then reports the version the tab itself handed the Host.
    h.live.projection = projectionAt(contentVersion)
    view.rerender(<OfficePanel {...h.props()} />)
    expect(h.remote.read).toHaveBeenCalledTimes(readBeforeSave)
    expect(h.handles).toHaveLength(1)
    expect(h.handles[0]).toBe(mounted)

    act(() => { mounted?.undo() })
    expect(h.handles[0]?.getSnapshot()).toEqual(BEFORE)
    act(() => { mounted?.redo() })
    expect(h.handles[0]?.getSnapshot()).toEqual(AFTER)
    act(() => { mounted?.undo() })
    expect(h.handles[0]?.getSnapshot()).toEqual(BEFORE)
    expect(h.handles).toHaveLength(1)
    expect(h.handles[0]).toBe(mounted)

    // Undoing the agent edit and saving writes the original value at the version
    // the Host published, so the version guard stays on the unit's own advance.
    fireEvent.click(view.getByText(en.save))
    await waitFor(() => { expect(h.remote.save).toHaveBeenCalledOnce() })
    expect(h.remote.save.mock.calls[0]?.[2]).toEqual(BEFORE)
    expect(h.remote.save.mock.calls[0]?.[4]).toBe('v2')
    expect(h.handles).toHaveLength(1)
  })

  it('adopts the projection version that arrives before the file version moves', async () => {
    const h = harness()
    const view = render(<OfficePanel {...h.props()} />)
    await waitFor(() => { expect(h.handles).toHaveLength(1) })
    const mounted = h.handles[0]

    const contentVersion = await applyAgentEdit(h, view)
    h.live.projection = projectionAt(contentVersion)
    view.rerender(<OfficePanel {...h.props()} />)
    expect(h.remote.read).toHaveBeenCalledOnce()

    h.live.meta = { absolutePath: `/${PATH}`, version: 'v2' }
    view.rerender(<OfficePanel {...h.props()} />)
    await waitFor(() => { expect(h.handles).toHaveLength(1) })
    expect(h.remote.read).toHaveBeenCalledOnce()
    expect(h.handles[0]).toBe(mounted)

    act(() => { mounted?.undo() })
    expect(h.handles[0]?.getSnapshot()).toEqual(BEFORE)
  })

  it('still reloads a clean unit when a genuinely external change follows the save', async () => {
    const h = harness()
    const view = render(<OfficePanel {...h.props()} />)
    await waitFor(() => { expect(h.handles).toHaveLength(1) })

    const contentVersion = await applyAgentEdit(h, view)
    h.live.meta = { absolutePath: `/${PATH}`, version: 'v2' }
    h.live.projection = projectionAt(contentVersion)
    view.rerender(<OfficePanel {...h.props()} />)
    expect(h.remote.read).toHaveBeenCalledOnce()

    h.live.readVersion = 'v3'
    h.live.meta = { absolutePath: `/${PATH}`, version: 'v3' }
    view.rerender(<OfficePanel {...h.props()} />)
    await waitFor(() => { expect(h.handles).toHaveLength(2) })
    expect(h.remote.read).toHaveBeenCalledTimes(2)
    expect(h.remote.read.mock.calls[1]?.[1]).toBe(PATH)
    expect(h.handles[1]).not.toBe(h.handles[0])
  })
})
