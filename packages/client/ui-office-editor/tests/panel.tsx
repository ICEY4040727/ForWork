/**
 * Driven fixture for `OfficePanel`: the injected face, the framework-bound
 * selector hooks over real stores, and the panel's copy seat.
 *
 * Specs provide only what the behavior under test needs; everything else keeps
 * its production wiring so a store write reaches the rendered body exactly as
 * the slot renderer would deliver it.
 */
import { render } from '@testing-library/react'
import type { RenderResult } from '@testing-library/react'
import { vi, type Mock } from 'vitest'
import { brandString } from '@forwork/brand'
import type { SessionId } from '@forwork/session/types'
import type { TabId } from '@forwork/client-ui-dockkit'
import type { SidebarRightOpenTab, SidebarRightTabActions } from '@forwork/client-ui-sidebar-right/client'
import { createSnapshotStore, type SnapshotStore } from '@forwork/client-store'
import type { WorkspaceFileStat } from '@forwork/api-workspace-files/types'
import type { RemoteFailure } from '@forwork/typert-protocol'
import { createMockAdapter } from '../src/client/adapter.ts'
import { createOfficeTabChanges } from '../src/client/changes.ts'
import type { OfficeDocumentRebind } from '../src/client/changes.ts'
import { createOfficeLegacyNotices } from '../src/client/notices.ts'
import { createOfficeRestore } from '../src/client/restore.ts'
import { OfficePanel } from '../src/client/OfficePanel.tsx'
import type { OfficePanelProps } from '../src/client/OfficePanel.tsx'
import type { OfficeDocumentRemote } from '../src/client/useOfficeDocument.ts'
import type { OfficePanelInjected } from '../src/client/face.ts'
import { en, type OfficeEditorKey } from '../src/client/locales.ts'
import type { DocumentProjectionView } from '@forwork/session-document/client'
import { bindObservable } from './observable.ts'

/** The Session every fixture addresses. */
export const SESSION_ID = brandString<SessionId>('s-1')

/** The tab id the fixture body renders as. */
export const TAB_ID = brandString<TabId>('tab-1')

/** The signal the fixture tab was opened with. */
export const SIGNAL = new AbortController().signal

/** The fixture's copy seat, interpolating `{name}` parameters like the locale service. */
export function translate(key: string, params?: Record<string, unknown>): string {
  const template = key in en ? en[key as OfficeEditorKey] : key
  if (params === undefined) return template
  return template.replace(/\{(\w+)\}/g, (match, name: string) => name in params ? String(params[name]) : match)
}

/** What one panel fixture needs. */
export interface PanelInput {
  /** The tab's content id; defaults to an editable workbook address. */
  readonly address?: string
  /** The tab id; defaults to the fixture tab. */
  readonly tabId?: TabId
  /** The tab's captured title. */
  readonly title?: string
  /** The `file` resource value; absent while the resource is loading. */
  readonly meta?: WorkspaceFileStat
  /** The Session's `document` projection value; absent while the Host has not published it. */
  readonly projection?: DocumentProjectionView
  /** A restore record to share across bodies; a fresh one by default. */
  readonly restore?: ReturnType<typeof createOfficeRestore>
  /** The Remote the panel calls. */
  readonly remote: OfficeDocumentRemote
  /** Open tab metadata the panel lists; empty by default. */
  readonly openTabs?: readonly SidebarRightOpenTab[]
  /** Tab actions the panel takes; the spies are returned for assertions. */
  readonly actions?: SidebarRightTabActions
  /** Editor mount the panel calls; defaults to a recording mock adapter. */
  readonly mount?: OfficePanelInjected['mountEditor']
  /** Command-bridge attachment the panel calls; defaults to a recorder. */
  readonly attachCommandBridge?: OfficePanelInjected['attachCommandBridge']
  /** The resource failure the panel observes; absent while the resource is healthy. */
  readonly failure?: RemoteFailure
  /** The shared change surface; a fresh one by default. */
  readonly changes?: ReturnType<typeof createOfficeTabChanges>
  /** The sink a fresh change surface records an observed move through. */
  readonly rebind?: OfficeDocumentRebind
}

/**
 * Build the panel's composed props for one fixture.
 * @param input - the address, Remote, resource value, and tab inventory.
 * @returns the props plus the stores and actions the spec asserts on.
 */
export function panelFixture(input: PanelInput): {
  props: OfficePanelProps
  changes: ReturnType<typeof createOfficeTabChanges>
  notices: ReturnType<typeof createOfficeLegacyNotices>
  restore: ReturnType<typeof createOfficeRestore>
  openTabs: SnapshotStore<readonly SidebarRightOpenTab[]>
  openResource: Mock<SidebarRightTabActions['openResource']>
  adapter: ReturnType<typeof createMockAdapter>
  attachCommandBridge: Mock<OfficePanelInjected['attachCommandBridge']>
  markMissing: Mock<OfficePanelInjected['markMissing']>
  trackDocument: Mock<OfficePanelInjected['trackDocument']>
  observeFile: Mock<OfficePanelInjected['observeFile']>
  forgetDocument: Mock<OfficePanelInjected['forgetDocument']>
  resourceAddresses: string[]
} {
  const address = input.address ?? 'dsh-resource://file/session/s-1/work/budget.xlsx'
  const tabId = input.tabId ?? TAB_ID
  const changes = input.changes ?? createOfficeTabChanges(input.rebind)
  const notices = createOfficeLegacyNotices()
  const restore = input.restore ?? createOfficeRestore()
  const openTabs = createSnapshotStore<readonly SidebarRightOpenTab[]>(input.openTabs ?? [])
  const openResource = vi.fn<SidebarRightTabActions['openResource']>()
  const attachCommandBridge = vi.fn<OfficePanelInjected['attachCommandBridge']>(() => () => {})
  const markMissing = vi.fn<OfficePanelInjected['markMissing']>()
  const markTabDirty = vi.fn<OfficePanelInjected['markTabDirty']>((tabId, dirty) => {
    changes.markDirty(tabId, dirty)
  })
  const dismissCloseWarning = vi.fn<OfficePanelInjected['dismissCloseWarning']>((tabId) => {
    changes.dismissWarning(tabId)
  })
  const trackDocument = vi.fn<OfficePanelInjected['trackDocument']>((sessionId, documentId, path) => {
    changes.trackDocument(sessionId, documentId, path)
  })
  const observeFile = vi.fn<OfficePanelInjected['observeFile']>((sessionId, documentId, path, version) => {
    changes.observeFile(sessionId, documentId, path, version)
  })
  const forgetDocument = vi.fn<OfficePanelInjected['forgetDocument']>((documentId) => {
    changes.forgetDocument(documentId)
  })
  const resourceAddresses: string[] = []
  const actions: SidebarRightTabActions = input.actions ?? {
    bindCommands: vi.fn(() => () => {}),
    openResource,
    openTab: vi.fn(),
    close: vi.fn(),
  }
  const adapter = createMockAdapter()
  const props = {
    useTabInfo: () => ({
      tab: {
        id: tabId,
        contentId: address,
        title: input.title ?? 'budget.xlsx',
        signal: SIGNAL,
        actions,
      },
    }),
    useProjection: (key: string) => key === 'document' ? input.projection : undefined,
    useResource: (resourceAddress: string) => {
      resourceAddresses.push(resourceAddress)
      return input.failure !== undefined
        ? { status: 'failed', value: input.meta, failure: input.failure }
        : { status: input.meta === undefined ? 'loading' : 'live', value: input.meta, failure: undefined }
    },
    remote: input.remote,
    mountEditor: input.mount ?? ((container, mountInput) => adapter.mount(container, mountInput)),
    attachCommandBridge: input.attachCommandBridge ?? attachCommandBridge,
    markTabDirty,
    dismissCloseWarning,
    bindDocument: (address: string, documentId: string) => { changes.bindDocument(address, documentId) },
    documentIdAt: (address: string) => changes.documentIdAt(address),
    trackDocument,
    forgetDocument,
    observeFile,
    markMissing,
    rememberLegacySource: (contentId: string, source: string) => { notices.remember(contentId, source) },
    dismissLegacyNotice: (contentId: string) => { notices.dismiss(contentId) },
    restore,
    dismissRestoreNotice: (sessionId: SessionId) => { restore.dismissMissing(sessionId) },
    useOfficeTab: bindObservable(changes.source),
    useDocumentLocations: bindObservable(changes.locations),
    useLegacyNotice: bindObservable(notices.source),
    useSidebarTabs: bindObservable(openTabs),
    useRestore: bindObservable(restore.source),
    t: translate,
  } as OfficePanelProps
  return {
    props, changes, notices, restore, openTabs, openResource, adapter, attachCommandBridge, markMissing,
    trackDocument, observeFile, forgetDocument, resourceAddresses,
  }
}

/**
 * Render one panel fixture.
 * @param input - the address, Remote, resource value, and tab inventory.
 * @returns the render result plus the stores and actions the spec asserts on.
 */
export function renderPanel(input: PanelInput): RenderResult & {
  fixture: ReturnType<typeof panelFixture>
} {
  const fixture = panelFixture(input)
  return { ...render(<OfficePanel {...fixture.props} />), fixture }
}
