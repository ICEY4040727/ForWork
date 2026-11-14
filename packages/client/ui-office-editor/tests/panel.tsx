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
import { brandString } from '@deepseek-ai/dsh-brand'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import type { TabId } from '@deepseek-ai/dsh-client-ui-dockkit'
import type { SidebarRightOpenTab, SidebarRightTabActions } from '@deepseek-ai/dsh-client-ui-sidebar-right/client'
import { createSnapshotStore, type SnapshotStore } from '@deepseek-ai/dsh-client-store'
import type { WorkspaceFileStat } from '@deepseek-ai/dsh-api-workspace-files/types'
import { createMockAdapter } from '../src/client/adapter.ts'
import { createOfficeTabChanges } from '../src/client/changes.ts'
import { createOfficeLegacyNotices } from '../src/client/notices.ts'
import { OfficePanel } from '../src/client/OfficePanel.tsx'
import type { OfficePanelProps } from '../src/client/OfficePanel.tsx'
import type { OfficeDocumentRemote } from '../src/client/useOfficeDocument.ts'
import type { OfficePanelInjected } from '../src/client/face.ts'
import { en, type OfficeEditorKey } from '../src/client/locales.ts'
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
  /** The Remote the panel calls. */
  readonly remote: OfficeDocumentRemote
  /** Open tab metadata the panel lists; empty by default. */
  readonly openTabs?: readonly SidebarRightOpenTab[]
  /** Tab actions the panel takes; the spies are returned for assertions. */
  readonly actions?: SidebarRightTabActions
  /** Editor mount the panel calls; defaults to a recording mock adapter. */
  readonly mount?: OfficePanelInjected['mountEditor']
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
  openTabs: SnapshotStore<readonly SidebarRightOpenTab[]>
  openResource: Mock<SidebarRightTabActions['openResource']>
  adapter: ReturnType<typeof createMockAdapter>
} {
  const address = input.address ?? 'dsh-resource://file/session/s-1/work/budget.xlsx'
  const tabId = input.tabId ?? TAB_ID
  const changes = createOfficeTabChanges()
  const notices = createOfficeLegacyNotices()
  const openTabs = createSnapshotStore<readonly SidebarRightOpenTab[]>(input.openTabs ?? [])
  const openResource = vi.fn<SidebarRightTabActions['openResource']>()
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
    useResource: () => ({
      status: input.meta === undefined ? 'loading' : 'live',
      value: input.meta,
      failure: undefined,
    }),
    remote: input.remote,
    mountEditor: input.mount ?? ((container, mountInput) => adapter.mount(container, mountInput)),
    markTabDirty: vi.fn(),
    dismissCloseWarning: vi.fn(),
    rememberLegacySource: (contentId: string, source: string) => { notices.remember(contentId, source) },
    dismissLegacyNotice: (contentId: string) => { notices.dismiss(contentId) },
    useOfficeTab: bindObservable(changes.source),
    useLegacyNotice: bindObservable(notices.source),
    useSidebarTabs: bindObservable(openTabs),
    t: translate,
  } as OfficePanelProps
  return { props, changes, notices, openTabs, openResource, adapter }
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
