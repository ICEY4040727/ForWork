/**
 * Browser half: register `office` as a right-Sidebar tab type.
 *
 * The type reaches the Sidebar through its public path only: the definition
 * into `ctx.sidebarRightTabs`, the body into the keyed `sidebar.right.pane.tab`
 * seat, the chip title into `sidebar.right.pane.tab.title`, and the selected
 * pane's header into `sidebar.right.pane.header`, all under the definition's
 * `id`. The body reads the file through the `officeDocument`
 * Remote, mounts the editor runtime behind the adapter seam, shares one
 * plugin-owned change surface with the tab close handler, and lists the
 * Session's open Office documents from the Sidebar's open-tab inventory. A
 * legacy source is converted and reopened as a copy through the tab's own
 * `openResource` action, with the original path kept in the plugin's notice
 * store. The same plugin-owned store records the one restore pass each Session
 * runs and the documents it could not reopen; that pass also runs when the
 * Session becomes active, so a Session with no Office tab restores its set.
 * A moved file observed through the open tabs' file resources is recorded as a
 * `document/rebind` through the same Remote.
 */
import type { Context as ClientContext } from '@forwork/cordis'
import type {} from '@forwork/api-remotes/client'
import type {} from '@forwork/api-session-controller/client'
import type {} from '@forwork/client-locale/client'
import type {} from '@forwork/client-ui-renderer/client'
import type {} from '@forwork/client-ui-sidebar-right/client'
// Type-only: pulls the Session standard useProjection seat and its identity.
import type {} from '@forwork/client-ui-session/client'
import type {} from '@forwork/client-ui-slots'
import type {} from '@forwork/office-exchange/remote'
import { officeAdapter } from './adapter.ts'
import { createOfficeTabChanges } from './changes.ts'
import { attachSessionCommandBridge, type OfficeCommandRemote } from './command-bridge.ts'
import { officePanelInjected } from './face.ts'
import { createOfficeLegacyNotices } from './notices.ts'
import { OfficeHeader } from './OfficeHeader.tsx'
import { OfficePanel } from './OfficePanel.tsx'
import { OfficeTitle } from './OfficeTitle.tsx'
import { createOfficeRestore, restoreOnSessionActivation } from './restore.ts'
import { OFFICE_ID, OFFICE_KIND, officeDefinition } from './definition.ts'
import { de, en, es, fr, ja, pt, ru, zh } from './locales.ts'
import { officeDocumentIdentity, openTargetOf } from './target.ts'

export type { OfficeEditorKey } from './locales.ts'

/** This package's copy namespace. */
const NS = 'officeEditor'

declare module '@forwork/client-ui-slots' {
  interface LocaleNamespaceMap {
    /** Editor body, save, refusal, external-change, and close-prompt copy. */
    officeEditor: import('./locales.ts').OfficeEditorKey
  }
}

/** Required browser services: the tab registry, the slots, copy, the Session feed, the Remote, and close cleanup. */
export const inject = ['slots', 'locale', 'sessions', 'sidebarRightTabs', 'sidebarRight', 'remote', 'remote.officeDocument']

/**
 * Client plugin body: register the type, its dictionaries, its body and chip
 * title, and the unsaved-close veto.
 * @param ctx - client root context carrying the registry, slots, copy, and Remote.
 */
export function apply(ctx: ClientContext): void {
  const changes = createOfficeTabChanges((sessionId, documentId, toPath) => {
    void ctx.remote.officeDocument.rebindDocument(sessionId, documentId, toPath).then((result) => {
      if (!result.ok) {
        ctx.logger.warn(`ui-office-editor: could not record the move of "${documentId}" to "${toPath}": ${result.error.message}`)
      }
    })
  })
  const notices = createOfficeLegacyNotices()
  const restore = createOfficeRestore()
  const commandRemote: OfficeCommandRemote = ctx.remote.officeDocument
  ctx.effect(() => ctx.sidebarRightTabs.register(officeDefinition()), 'ui-office-editor: type')
  ctx.effect(() => ctx.locale.register(NS, { zh, en, ja, es, fr, de, ru, pt }), 'ui-office-editor: dictionaries')
  ctx.effect(() => restoreOnSessionActivation({
    sessions: ctx.sessions.list,
    openTabs: ctx.sidebarRight.openTabs,
    restore,
    bindDocument: (address, documentId) => { changes.bindDocument(address, documentId) },
    openResourceIn: (sessionId, address) => { ctx.sidebarRight.openResourceIn(sessionId, address) },
  }), 'ui-office-editor: activation restore')
  ctx.effect(() => ctx.sidebarRight.registerCloseHandler(OFFICE_KIND, (sessionId, tab) => {
    // Start the pending draft before the warning, so the edits are on disk even
    // when the user cancels the prompt and the process then dies.
    changes.flushDraft(tab.id)
    if (changes.isDirty(tab.id)) {
      changes.warnClose(tab.id)
      throw new Error(`ui-office-editor: "${tab.title}" has unsaved changes`)
    }
    changes.forget(tab.id)
    notices.forget(tab.contentId)
    const target = openTargetOf(tab.contentId)
    // A rebound document keeps the identity it was opened with, so the close
    // fact names the same document instead of one re-derived from the old address.
    const documentId = changes.documentIdAt(tab.contentId)
      ?? officeDocumentIdentity(sessionId, target.path)
    changes.releaseDocument(tab.contentId, documentId)
    void ctx.remote.officeDocument.closeDocument(sessionId, documentId, 'closed').then((result) => {
      if (!result.ok) ctx.logger.warn(`ui-office-editor: could not record closing "${tab.title}": ${result.error.message}`)
    })
  }), 'ui-office-editor: close')
  ctx.effect(() => ctx.slots.inject('sidebar.right.pane.tab', () => ctx.slots.register(
    {
      name: 'sidebar.right.pane.tab', key: OFFICE_ID, locale: NS,
      inject: () => officePanelInjected({
        read: (sessionId, path, signal) => ctx.remote.officeDocument.read(sessionId, path, signal),
        save: (sessionId, path, snapshot, format, expectedVersion, attribution, signal) =>
          ctx.remote.officeDocument.save(sessionId, path, snapshot, format, expectedVersion, attribution, signal),
        convert: (sessionId, path, signal) => ctx.remote.officeDocument.convertLegacy(sessionId, path, signal),
        writeDraft: (sessionId, documentId, documentPath, contentVersion, snapshot, signal) =>
          ctx.remote.officeDocument.writeDraft(
            sessionId, documentId, documentPath, contentVersion, snapshot, signal,
          ),
        readDraft: (sessionId, documentId, signal) =>
          ctx.remote.officeDocument.readDraft(sessionId, documentId, signal),
        clearDraft: (sessionId, documentId, signal) =>
          ctx.remote.officeDocument.clearDraft(sessionId, documentId, signal),
        adapter: officeAdapter,
        attachCommandBridge: (sessionId, target) =>
          attachSessionCommandBridge(ctx.sessions.binding(sessionId), sessionId, target, commandRemote),
        markMissing: (sessionId, documentId) => {
          void ctx.remote.officeDocument.markDocumentMissing(sessionId, documentId).then((result) => {
            if (!result.ok) ctx.logger.warn(`ui-office-editor: could not record "${documentId}" missing: ${result.error.message}`)
          })
        },
        changes,
        notices,
        restore,
        openTabs: ctx.sidebarRight.openTabs,
      }),
    },
    OfficePanel,
  )), 'ui-office-editor: body')
  ctx.effect(() => ctx.slots.inject('sidebar.right.pane.tab.title', () => ctx.slots.register(
    { name: 'sidebar.right.pane.tab.title', key: OFFICE_ID },
    OfficeTitle,
  )), 'ui-office-editor: title')
  ctx.effect(() => ctx.slots.inject('sidebar.right.pane.header', () => ctx.slots.register(
    { name: 'sidebar.right.pane.header', key: OFFICE_ID, locale: NS },
    OfficeHeader,
  )), 'ui-office-editor: pane header')
}
