/** The panel's injected business face, shared by the registration and its tests. */
import type { HostObservable } from '@deepseek-ai/dsh-client-ui-slots'
import type { TabId } from '@deepseek-ai/dsh-client-ui-dockkit'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import type { SidebarRightOpenTab } from '@deepseek-ai/dsh-client-ui-sidebar-right/client'
import type { OfficeAdapter, OfficeEditorHandle, OfficeMountInput } from './adapter.ts'
import type { OfficeCommandTarget } from './command-bridge.ts'
import type { OfficeTabChangeSnapshot, OfficeTabChanges } from './changes.ts'
import type { OfficeLegacyNoticeSnapshot, OfficeLegacyNotices } from './notices.ts'
import type { OfficeDocumentRemote } from './useOfficeDocument.ts'
import type { OfficeRestore, OfficeRestoreActions, OfficeRestoreState } from './restore.ts'

/** The panel's injected face: the Remote, the runtime mount, the restore record, and the shared tab state. */
export interface OfficePanelInjected {
  /** The `officeDocument` Remote the hook calls. */
  readonly remote: OfficeDocumentRemote
  /** Mount the editor runtime behind the adapter seam. */
  readonly mountEditor: (container: HTMLElement, input: OfficeMountInput) => Promise<OfficeEditorHandle>
  /**
   * Subscribe one tab to its Session's `document/command` facts.
   * @param sessionId - the Session whose command feed is read.
   * @param target - the live tab state and adapter the bridge acts on.
   * @returns stops listening for commands for this tab.
   */
  readonly attachCommandBridge: (sessionId: SessionId, target: OfficeCommandTarget) => () => void
  /** Record whether this tab's live unit has unsaved edits. */
  readonly markTabDirty: (tabId: TabId, dirty: boolean) => void
  /** Dismiss the close prompt without saving or closing. */
  readonly dismissCloseWarning: (tabId: TabId) => void
  /**
   * Record that the open document's file no longer resolves.
   * @param sessionId - the Session whose open set holds the document.
   * @param documentId - Session-scoped identity of the missing document.
   */
  readonly markMissing: (sessionId: SessionId, documentId: string) => void
  /**
   * Record which legacy source a converted copy came from.
   * @param contentId - the converted tab's address.
   * @param source - the legacy workspace path the user opened.
   */
  readonly rememberLegacySource: (contentId: string, source: string) => void
  /**
   * Hide the converted-copy notice.
   * @param contentId - the converted tab's address.
   */
  readonly dismissLegacyNotice: (contentId: string) => void
  /** The plugin-owned restore operations this body may invoke. */
  readonly restore: OfficeRestoreActions
  /**
   * Hide one Session's missing-document notice.
   * @param sessionId - the Session whose notice the user dismissed.
   */
  readonly dismissRestoreNotice: (sessionId: SessionId) => void
  readonly hooks: {
    readonly officeTab: HostObservable<OfficeTabChangeSnapshot>
    readonly legacyNotice: HostObservable<OfficeLegacyNoticeSnapshot>
    readonly sidebarTabs: HostObservable<readonly SidebarRightOpenTab[]>
    readonly restore: HostObservable<OfficeRestoreState>
  }
}

/** The dependencies one session's panel registration binds. */
export interface OfficePanelDeps {
  /** The `officeDocument` Remote read call. */
  readonly read: OfficeDocumentRemote['read']
  /** The `officeDocument` Remote save call. */
  readonly save: OfficeDocumentRemote['save']
  /** The `officeDocument` Remote legacy conversion call. */
  readonly convert: OfficeDocumentRemote['convertLegacy']
  /** The `officeDocument` Remote draft write call. */
  readonly writeDraft: OfficeDocumentRemote['writeDraft']
  /** The `officeDocument` Remote draft read call. */
  readonly readDraft: OfficeDocumentRemote['readDraft']
  /** The `officeDocument` Remote draft removal call. */
  readonly clearDraft: OfficeDocumentRemote['clearDraft']
  /** The editor runtime adapter. */
  readonly adapter: OfficeAdapter
  /** Subscribe one tab to its Session's `document/command` facts. */
  readonly attachCommandBridge: OfficePanelInjected['attachCommandBridge']
  /** Record a file resource that reported the open document missing. */
  readonly markMissing: OfficePanelInjected['markMissing']
  /** The plugin-owned change surface. */
  readonly changes: OfficeTabChanges
  /** The plugin-owned converted-copy notices. */
  readonly notices: OfficeLegacyNotices
  /** The plugin-owned restore record. */
  readonly restore: OfficeRestore
  /** The Sidebar's open-tab inventory across adopted Sessions. */
  readonly openTabs: HostObservable<readonly SidebarRightOpenTab[]>
}

/**
 * Build the panel's injected face.
 * @param deps - the Remote, the editor adapter, the shared tab and restore state, and the open-tab inventory.
 * @returns the callbacks and observables the panel receives.
 */
export function officePanelInjected(deps: OfficePanelDeps): OfficePanelInjected {
  const {
    read, save, convert, writeDraft, readDraft, clearDraft, adapter, attachCommandBridge, markMissing, changes, notices,
    restore, openTabs,
  } = deps
  return {
    remote: { read, save, convertLegacy: convert, writeDraft, readDraft, clearDraft },
    mountEditor: (container, input) => adapter.mount(container, input),
    attachCommandBridge,
    markTabDirty: (tabId, dirty) => { changes.markDirty(tabId, dirty) },
    dismissCloseWarning: (tabId) => { changes.dismissWarning(tabId) },
    markMissing,
    rememberLegacySource: (contentId, source) => { notices.remember(contentId, source) },
    dismissLegacyNotice: (contentId) => { notices.dismiss(contentId) },
    restore: {
      claim: sessionId => restore.claim(sessionId),
      reportMissing: (sessionId, paths) => { restore.reportMissing(sessionId, paths) },
    },
    dismissRestoreNotice: (sessionId) => { restore.dismissMissing(sessionId) },
    hooks: {
      officeTab: changes.source,
      legacyNotice: notices.source,
      sidebarTabs: openTabs,
      restore: restore.source,
    },
  }
}
