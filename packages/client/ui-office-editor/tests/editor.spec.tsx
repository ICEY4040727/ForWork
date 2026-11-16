// @vitest-environment jsdom
/**
 * The office tab type's registration, close veto, and mounted panel states.
 *
 * The registry is real, because "registered" means what it says a type claims;
 * the slot and locale faces are recorders, because what matters here is what
 * was handed to them and that every registration is gone after dispose. The
 * body renders with a fake Remote and adapter so no Univer code loads.
 */
import { afterEach, describe, expect, it, onTestFinished, vi } from 'vitest'
import type { Mock } from 'vitest'
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { Context } from '@forwork/cordis'
import { brandString } from '@forwork/brand'
import type { SessionId } from '@forwork/session/types'
import { SessionSeq } from '@forwork/session/types'
import type { DocumentId } from '@forwork/session/src/document-events.ts'
import type { WorkspaceFileStat } from '@forwork/api-workspace-files/types'
import type { TabId } from '@forwork/client-ui-dockkit'
import type { SidebarRightOpenTab } from '@forwork/client-ui-sidebar-right/client'
import { createSnapshotStore } from '@forwork/client-store'
import { MutableSessionEventSource } from '@forwork/api-session-controller/client'
import type { RemoteResult } from '@forwork/typert-protocol'
import { RemoteError } from '@forwork/typert-protocol'
import type { OfficeDocumentReadResult, OfficeDocumentSaveResult, OfficeDocumentSnapshot } from '@forwork/office-exchange/types'
import { SidebarRightTabRegistry } from '@forwork/client-ui-sidebar-right/src/client/tab-registry.ts'
import { createMockAdapter } from '../src/client/adapter.ts'
import type { OfficeMountInput } from '../src/client/adapter.ts'
import type { OfficeTabChangeSnapshot } from '../src/client/changes.ts'
import { OFFICE_ID, OFFICE_KIND } from '../src/client/definition.ts'
import { apply, inject } from '../src/client/index.ts'
import { OfficePanel } from '../src/client/OfficePanel.tsx'
import type { OfficePanelProps } from '../src/client/OfficePanel.tsx'
import type { OfficeCommandTarget } from '../src/client/command-bridge.ts'
import type { OfficePanelInjected } from '../src/client/face.ts'
import { OfficeTitle } from '../src/client/OfficeTitle.tsx'
import type { OfficeTitleProps } from '../src/client/OfficeTitle.tsx'
import { createOfficeRestore } from '../src/client/restore.ts'
import type { OfficeRestoreState } from '../src/client/restore.ts'
import type { OfficeDocumentRemote } from '../src/client/useOfficeDocument.ts'
import { en, zh, type OfficeEditorKey } from '../src/client/locales.ts'

const ADDRESS = 'dsh-resource://file/session/s-1/work/budget.xlsx'
const SESSION_ID = brandString<SessionId>('s-1')
const TAB_ID = brandString<TabId>('tab-1')
const SIGNAL = new AbortController().signal

afterEach(cleanup)

const snapshot: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book' } }

function readResult(version: string, overrides: Partial<OfficeDocumentReadResult> = {}): OfficeDocumentReadResult {
  return { format: 'xlsx', snapshot, unsupported: [], version, ...overrides }
}

function success<T>(value: T): RemoteResult<T> {
  return { ok: true, value }
}

function failure<T>(code: 'office-document/conflict' | 'office-document/out-of-scope'): RemoteResult<T> {
  return { ok: false, error: new RemoteError(code, code, { path: '/work/budget.xlsx' }) }
}

function fakeRemote(options: { factFailure?: boolean } = {}): {
  read: ReturnType<typeof vi.fn>
  convert: ReturnType<typeof vi.fn>
  save: ReturnType<typeof vi.fn>
  writeDraft: ReturnType<typeof vi.fn>
  readDraft: ReturnType<typeof vi.fn>
  clearDraft: ReturnType<typeof vi.fn>
  complete: ReturnType<typeof vi.fn>
  closeDocument: ReturnType<typeof vi.fn>
  markDocumentMissing: ReturnType<typeof vi.fn>
  face: OfficeDocumentRemote & {
    completeCommand: ReturnType<typeof vi.fn>
    closeDocument: ReturnType<typeof vi.fn>
    markDocumentMissing: ReturnType<typeof vi.fn>
  }
} {
  const read = vi.fn(async (): Promise<RemoteResult<OfficeDocumentReadResult>> => success(readResult('v1')))
  const convert = vi.fn(async (): Promise<RemoteResult<string>> => success('/work/report.docx'))
  const save = vi.fn(async (): Promise<RemoteResult<OfficeDocumentSaveResult>> => success({ version: 'v2' }))
  const writeDraft = vi.fn(async (): Promise<RemoteResult<{ savedAt: string }>> => success({ savedAt: 't' }))
  const readDraft = vi.fn(async (): Promise<RemoteResult<null>> => success(null))
  const clearDraft = vi.fn(async (): Promise<RemoteResult<void>> => success(undefined))
  const complete = vi.fn(async () => undefined)
  const fact = (): RemoteResult<void> => options.factFailure === true
    ? failure<undefined>('office-document/out-of-scope')
    : success(undefined)
  const closeDocument = vi.fn(async (): Promise<RemoteResult<void>> => fact())
  const markDocumentMissing = vi.fn(async (): Promise<RemoteResult<void>> => fact())
  return {
    read,
    convert,
    save,
    writeDraft,
    readDraft,
    clearDraft,
    complete,
    closeDocument,
    markDocumentMissing,
    face: {
      read,
      convertLegacy: convert,
      save,
      writeDraft,
      readDraft,
      clearDraft,
      completeCommand: complete,
      closeDocument,
      markDocumentMissing,
    },
  }
}

/** Draft Remote methods for panels whose behavior under test does not touch drafts. */
function idleDrafts(): Pick<OfficeDocumentRemote, 'writeDraft' | 'readDraft' | 'clearDraft'> {
  return {
    writeDraft: vi.fn(),
    readDraft: vi.fn(async () => success(null)),
    clearDraft: vi.fn(async () => success(undefined)),
  }
}

interface Recorded {
  name: string
  key: string
  locale?: string
  component: unknown
  inject?: (sessionId: unknown) => OfficePanelInjected
}

interface Boot {
  tabs: SidebarRightTabRegistry
  registered: Recorded[]
  dictionaries: Map<string, unknown>
  closeHandler: ((sessionId: unknown, tab: { id: TabId; title: string; contentId: string }) => void) | undefined
  fiber: ReturnType<Context['plugin']>
  commands: MutableSessionEventSource
  completeCommand: ReturnType<typeof vi.fn>
  warn: Mock<(message: string) => void>
}

async function boot(options: { factFailure?: boolean } = {}): Promise<Boot> {
  const ctx = new Context()
  const warn = vi.fn<(message: string) => void>()
  ctx.logger.warn = warn
  const tabs = new SidebarRightTabRegistry(ctx)
  const registered: Recorded[] = []
  const slots = {
    inject: vi.fn((_name: string, register: () => () => void) => register()),
    register: vi.fn((options: Omit<Recorded, 'component'>, component: unknown) => {
      const entry: Recorded = { ...options, component }
      registered.push(entry)
      return () => { registered.splice(registered.indexOf(entry), 1) }
    }),
  }
  const dictionaries = new Map<string, unknown>()
  const locale = {
    register: vi.fn((ns: string, dicts: unknown) => {
      dictionaries.set(ns, dicts)
      return () => { dictionaries.delete(ns) }
    }),
  }
  let closeHandler: Boot['closeHandler']
  const sidebarRight = {
    openTabs: createSnapshotStore<readonly SidebarRightOpenTab[]>([]),
    registerCloseHandler: vi.fn((_kind: string, handler: NonNullable<Boot['closeHandler']>) => {
      closeHandler = handler
      return () => { closeHandler = undefined }
    }),
  }
  const remote = fakeRemote(options)
  const commands = new MutableSessionEventSource()
  ctx.provide('sidebarRightTabs', tabs as never)
  ctx.provide('slots', slots as never)
  ctx.provide('locale', locale as never)
  ctx.provide('sessions', { binding: vi.fn(() => ({ eventSource: commands })) } as never)
  ctx.provide('sidebarRight', sidebarRight as never)
  ctx.provide('remote', { officeDocument: remote.face } as never)
  ctx.provide('remote.officeDocument', remote.face as never)
  const fiber = ctx.plugin({ inject: [...inject], apply })
  onTestFinished(async () => { await fiber.dispose() })
  await fiber.await()
  return {
    tabs, registered, dictionaries, closeHandler, fiber, commands, completeCommand: remote.complete, warn,
  }
}

interface PanelOptions {
  readonly value?: WorkspaceFileStat
  readonly title?: string
  readonly observedVersion?: string
  readonly dirty?: boolean
  readonly warned?: boolean
  readonly remote?: OfficeDocumentRemote
  readonly mount?: OfficePanelInjected['mountEditor']
  readonly onMarkDirty?: OfficePanelInjected['markTabDirty']
  readonly onDismissWarning?: OfficePanelInjected['dismissCloseWarning']
  readonly onClose?: () => void
}

function panelProps(options: PanelOptions = {}): OfficePanelProps {
  const adapter = createMockAdapter()
  const changes: OfficeTabChangeSnapshot = options.dirty === true || options.warned === true
    ? { [TAB_ID]: { dirty: options.dirty === true, warned: options.warned === true } }
    : {}
  return {
    useTabInfo: () => ({
      tab: {
        id: TAB_ID,
        contentId: ADDRESS,
        title: options.title ?? 'budget.xlsx',
        signal: SIGNAL,
        actions: { close: options.onClose ?? vi.fn(), openResource: vi.fn() },
      },
    }),
    useProjection: () => undefined,
    useRestore: (selector: (value: OfficeRestoreState) => unknown) => selector({ restored: [], missing: {} }),
    useResource: () => ({
      status: options.value === undefined ? 'loading' : 'live',
      value: options.value,
      failure: undefined,
    }),
    remote: options.remote ?? fakeRemote().face,
    mountEditor: options.mount ?? ((container, input) => adapter.mount(container, input)),
    attachCommandBridge: vi.fn(() => () => {}),
    markTabDirty: options.onMarkDirty ?? vi.fn(),
    dismissCloseWarning: options.onDismissWarning ?? vi.fn(),
    markMissing: vi.fn(),
    rememberLegacySource: vi.fn(),
    dismissLegacyNotice: vi.fn(),
    restore: createOfficeRestore(),
    dismissRestoreNotice: vi.fn(),
    useSidebarTabs: (selector: (value: readonly SidebarRightOpenTab[]) => unknown) => selector([]),
    useLegacyNotice: (selector: (value: Record<string, never>) => unknown) => selector({}),
    useOfficeTab: (selector: (value: OfficeTabChangeSnapshot) => unknown) => selector(changes),
    t: (key: OfficeEditorKey) => en[key],
  } as OfficePanelProps
}

describe('ui-office-editor apply', () => {
  it('registers the office type, its dictionaries, and the body and title seats under its id', async () => {
    const { tabs, registered, dictionaries } = await boot()
    const definition = tabs.get(OFFICE_KIND)
    expect(definition?.id).toBe(OFFICE_ID)
    expect(definition?.patterns).toEqual(['*.docx', '*.xlsx', '*.doc', '*.xls', '*.ppt'])
    expect(definition?.priority).toBe('builtin')
    expect(definition?.keepMounted).toBe(true)
    expect(dictionaries.get('officeEditor')).toEqual({ zh, en })
    expect(registered.map(entry => [entry.name, entry.key, entry.locale, entry.component])).toEqual([
      ['sidebar.right.pane.tab', OFFICE_ID, 'officeEditor', OfficePanel],
      ['sidebar.right.pane.tab.title', OFFICE_ID, undefined, OfficeTitle],
    ])
  })

  it('settles a Session document command through the officeDocument Remote', async () => {
    const { registered, commands, completeCommand } = await boot()
    const face = registered.find(entry => entry.name === 'sidebar.right.pane.tab')?.inject?.(SESSION_ID)
    expect(face).toBeDefined()
    const apply = vi.fn(() => ({ status: 'applied' as const }))
    const target: OfficeCommandTarget = {
      documentId: 'document-1',
      version: () => 2,
      isOpen: () => true,
      isDirty: () => false,
      unsupported: () => [],
      apply,
      snapshot: () => snapshot,
    }
    const detach = face!.attachCommandBridge(SESSION_ID, target)
    commands.append({
      type: 'event',
      event: {
        type: 'document/command',
        seq: SessionSeq(1),
        time: 1,
        data: {
          requestId: 'r-1',
          documentId: brandString<DocumentId>('document-1'),
          operation: { kind: 'read' },
          fromVersion: 2,
        },
      },
    })
    expect(apply).not.toHaveBeenCalled()
    expect(completeCommand).toHaveBeenCalledWith(SESSION_ID, 'r-1', {
      status: 'read',
      snapshot,
      contentVersion: 2,
    })
    detach()
  })

  it('takes every registration back when the plugin is disposed', async () => {    const { tabs, registered, dictionaries, fiber } = await boot()
    await fiber.dispose()
    expect(tabs.get(OFFICE_KIND)).toBeUndefined()
    expect(registered).toEqual([])
    expect(dictionaries.size).toBe(0)
  })

  it('vetoes a dirty tab close and flags the panel prompt', async () => {
    const { registered, closeHandler } = await boot()
    const registration = registered.find(entry => entry.name === 'sidebar.right.pane.tab')
    const face = registration?.inject?.(SESSION_ID)
    expect(face).toBeDefined()
    await face!.remote.read(SESSION_ID, 'work/budget.xlsx')
    await face!.remote.convertLegacy(SESSION_ID, 'work/legacy.doc')
    await face!.remote.save(SESSION_ID, 'work/budget.xlsx', snapshot, 'xlsx', 'v1')
    await face!.remote.writeDraft(SESSION_ID, 'doc-1', 'work/budget.xlsx', 'v1', snapshot)
    await face!.remote.readDraft(SESSION_ID, 'doc-1')
    await face!.remote.clearDraft(SESSION_ID, 'doc-1')
    face!.markTabDirty(TAB_ID, true)
    expect(() => closeHandler?.(SESSION_ID, { id: TAB_ID, title: 'budget.xlsx', contentId: ADDRESS }))
      .toThrow('unsaved changes')
    expect(face!.hooks.officeTab.getSnapshot()[TAB_ID]).toEqual({ dirty: true, warned: true })
    face!.dismissCloseWarning(TAB_ID)
    expect(face!.hooks.officeTab.getSnapshot()[TAB_ID]).toEqual({ dirty: true, warned: false })
    face!.markTabDirty(TAB_ID, false)
    face!.rememberLegacySource(ADDRESS, 'work/legacy.doc')
    expect(() => closeHandler?.(SESSION_ID, { id: TAB_ID, title: 'budget.xlsx', contentId: ADDRESS })).not.toThrow()
    expect(face!.hooks.legacyNotice.getSnapshot()).toEqual({})
  })

  it('logs a refused close or missing fact without blocking the tab', async () => {
    const { registered, closeHandler, warn } = await boot({ factFailure: true })
    const face = registered.find(entry => entry.name === 'sidebar.right.pane.tab')?.inject?.(SESSION_ID)
    expect(() => closeHandler?.(SESSION_ID, { id: TAB_ID, title: 'budget.xlsx', contentId: ADDRESS })).not.toThrow()
    face?.markMissing(SESSION_ID, `${SESSION_ID}:work/budget.xlsx`)
    await Promise.resolve()
    await Promise.resolve()
    expect(warn).toHaveBeenCalledTimes(2)
  })
})

describe('OfficePanel', () => {
  it('shows the loading line before the read settles', () => {
    const remote: OfficeDocumentRemote = {
      read: () => new Promise(() => {}),
      convertLegacy: vi.fn(),
      save: vi.fn(),
      ...idleDrafts(),
    }
    const view = render(<OfficePanel {...panelProps({ remote })} />)
    expect(view.getByText(en.loading)).toBeDefined()
  })

  it('reports a failed read with its localized reason', async () => {
    const remote: OfficeDocumentRemote = {
      read: async () => failure('office-document/out-of-scope'),
      convertLegacy: vi.fn(),
      save: vi.fn(),
      ...idleDrafts(),
    }
    const view = render(<OfficePanel {...panelProps({ remote, value: { absolutePath: '/work/budget.xlsx', version: 'v1' } })} />)
    expect(await view.findByText(en['error.outOfScope'])).toBeDefined()
  })

  it('mounts the adapter after a successful read and saves the live snapshot', async () => {
    const remote = fakeRemote()
    const adapter = createMockAdapter()
    const markDirty = vi.fn()
    const view = render(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v1' },
      remote: remote.face,
      mount: (container, input) => adapter.mount(container, input),
      onMarkDirty: markDirty,
    })} />)
    await waitFor(() => { expect(adapter.mounts).toHaveLength(1) })
    expect(adapter.mounts[0]?.kind).toBe('workbook')
    expect(adapter.mounts[0]?.snapshot).toBe(snapshot)
    adapter.setDirty(true)
    expect(markDirty).toHaveBeenCalledWith(TAB_ID, true)
    fireEvent.click(view.getByText(en.save))
    await waitFor(() => { expect(remote.save).toHaveBeenCalledOnce() })
    expect(remote.save.mock.calls[0]?.[4]).toBe('v1')
    await waitFor(() => { expect(markDirty).toHaveBeenCalledWith(TAB_ID, false) })
  })

  it('shows the named unsupported features and refuses to call save', async () => {
    const remote = fakeRemote()
    remote.read.mockResolvedValueOnce(success(readResult('v1', {
      unsupported: [{ kind: 'macro' }, { kind: 'chart', detail: 'Sheet1' }],
    })))
    const view = render(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v1' },
      remote: remote.face,
    })} />)
    expect(await view.findByText(en['feature.macro'])).toBeDefined()
    expect(view.getByText(en['feature.chart'])).toBeDefined()
    expect(view.getByText('Sheet1')).toBeDefined()
    fireEvent.click(view.getByText(en.save))
    expect(remote.save).not.toHaveBeenCalled()
  })

  it('prompts instead of reloading when the file changes while dirty', async () => {
    const remote = fakeRemote()
    const adapter = createMockAdapter()
    const mountEditor: OfficePanelInjected['mountEditor'] = (container, input) => adapter.mount(container, input)
    const markDirty = vi.fn()
    const view = render(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v1' },
      observedVersion: 'v1',
      remote: remote.face,
      mount: mountEditor,
      onMarkDirty: markDirty,
    })} />)
    await waitFor(() => { expect(adapter.mounts).toHaveLength(1) })
    view.rerender(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v2' },
      observedVersion: 'v2',
      dirty: true,
      remote: remote.face,
      mount: mountEditor,
      onMarkDirty: markDirty,
    })} />)
    expect(await view.findByText(en.externalChange)).toBeDefined()
    expect(remote.read).toHaveBeenCalledOnce()
    fireEvent.click(view.getByText(en.keepEditing))
    expect(view.queryByText(en.externalChange)).toBeNull()
    view.rerender(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v3' },
      observedVersion: 'v3',
      dirty: true,
      remote: remote.face,
      mount: mountEditor,
      onMarkDirty: markDirty,
    })} />)
    fireEvent.click(await view.findByText(en.reload))
    await waitFor(() => { expect(remote.read).toHaveBeenCalledTimes(2) })
  })

  it('offers the save-or-discard prompt after a vetoed close', async () => {
    const remote = fakeRemote()
    const close = vi.fn()
    const dismiss = vi.fn()
    const markDirty = vi.fn()
    const view = render(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v1' },
      dirty: true,
      warned: true,
      remote: remote.face,
      onClose: close,
      onDismissWarning: dismiss,
      onMarkDirty: markDirty,
    })} />)
    expect(await view.findByText(en.closeWarning)).toBeDefined()
    fireEvent.click(view.getByText(en.cancel))
    expect(dismiss).toHaveBeenCalledWith(TAB_ID)
    fireEvent.click(view.getByText(en.discardAndClose))
    expect(markDirty).toHaveBeenCalledWith(TAB_ID, false)
    expect(dismiss).toHaveBeenCalledTimes(2)
    expect(close).toHaveBeenCalledOnce()
  })

  it('saves before closing from the close prompt', async () => {
    const remote = fakeRemote()
    const close = vi.fn()
    const view = render(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v1' },
      dirty: true,
      warned: true,
      remote: remote.face,
      onClose: close,
    })} />)
    fireEvent.click(await view.findByText(en.saveAndClose))
    await waitFor(() => { expect(close).toHaveBeenCalledOnce() })
  })

  it('shows a save conflict and offers reload', async () => {
    const remote = fakeRemote()
    remote.save.mockResolvedValueOnce(failure('office-document/conflict'))
    const view = render(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v1' },
      remote: remote.face,
    })} />)
    fireEvent.click(await view.findByText(en.save))
    expect(await view.findByText(en['error.conflict'])).toBeDefined()
    fireEvent.click(view.getByText(en.reload))
    await waitFor(() => { expect(remote.read).toHaveBeenCalledTimes(2) })
  })

  it('shows a non-conflict save failure', async () => {
    const remote = fakeRemote()
    remote.save.mockResolvedValueOnce(failure('office-document/out-of-scope'))
    const view = render(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v1' },
      remote: remote.face,
    })} />)
    fireEvent.click(await view.findByText(en.save))
    expect(await view.findByText(en['error.outOfScope'])).toBeDefined()
  })

  it('reports a runtime that cannot start', async () => {
    const view = render(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v1' },
      mount: async (_container: HTMLElement, _input: OfficeMountInput) => { throw new Error('no canvas') },
    })} />)
    expect(await view.findByText(en.editorUnavailable)).toBeDefined()
  })

  it('disposes a runtime that resolves after the body unmounts', async () => {
    let resolve: (handle: Awaited<ReturnType<OfficePanelInjected['mountEditor']>>) => void = () => {}
    const dispose = vi.fn()
    const mount = vi.fn(() => new Promise<Awaited<ReturnType<OfficePanelInjected['mountEditor']>>>((done) => {
      resolve = done
    }))
    const view = render(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v1' },
      mount,
    })} />)
    await waitFor(() => { expect(mount).toHaveBeenCalledOnce() })
    view.unmount()
    resolve({
      dispose,
      getSnapshot: () => snapshot,
      apply: vi.fn(),
      undo: vi.fn(),
      redo: vi.fn(),
      isDirty: () => false,
      clearDirty: vi.fn(),
    })
    await waitFor(() => { expect(dispose).toHaveBeenCalledOnce() })
  })

  it('ignores a runtime failure that arrives after the body unmounts', async () => {
    let reject: (reason: unknown) => void = () => {}
    const mount = vi.fn(() => new Promise<Awaited<ReturnType<OfficePanelInjected['mountEditor']>>>((_done, fail) => {
      reject = fail
    }))
    const view = render(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v1' },
      mount,
    })} />)
    await waitFor(() => { expect(mount).toHaveBeenCalledOnce() })
    view.unmount()
    reject(new Error('late'))
    await Promise.resolve()
    await Promise.resolve()
  })

  it('disposes the runtime and clears dirty state when the body unmounts', async () => {
    const adapter = createMockAdapter()
    const markDirty = vi.fn()
    const view = render(<OfficePanel {...panelProps({
      value: { absolutePath: '/work/budget.xlsx', version: 'v1' },
      mount: (container, input) => adapter.mount(container, input),
      onMarkDirty: markDirty,
    })} />)
    await waitFor(() => { expect(adapter.mounts).toHaveLength(1) })
    view.unmount()
    expect(markDirty).toHaveBeenCalledWith(TAB_ID, false)
  })
})

describe('OfficeTitle', () => {
  it('renders the file type sheet before the captured tab title', () => {
    const props = { useTabInfo: () => ({ tab: { title: 'budget.xlsx' } }) } as OfficeTitleProps
    const view = render(<OfficeTitle {...props} />)
    expect(view.getByText('budget.xlsx')).toBeDefined()
  })
})
