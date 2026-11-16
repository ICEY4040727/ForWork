/**
 * Restore a Session's durable open-document set into the right Sidebar.
 *
 * The `document` projection is the only authority for which documents were
 * open, in which order, and at which path after any rebind; the Sidebar's own
 * layout storage is per-window and the Session log outlives it. Restore runs
 * once per Session per window, from the first trigger in that window: an Office
 * tab body that mounts in the Session, or the Session becoming active in the
 * main view, which restores a Session with no Office tab open. The trigger asks
 * the plugin-owned record below whether this Session still needs its set, and
 * only then opens the missing documents in projection order through an
 * `openResource` action.
 *
 * A `missing` document keeps its place in the Session (FR-015), so restore
 * never opens a dead tab for it: its path is recorded in the same store and
 * the panel lists them in a localized notice until the user dismisses it.
 *
 * Restore never reads or applies a recovery draft. Each document it opens
 * mounts its own panel, and that panel's existing read-then-offer path asks
 * the `officeDocument` Remote for the stored draft, so recovery has one
 * implementation (see `useOfficeDocument`).
 */
import { useEffect } from 'react'
import { createSnapshotStore, type ObservableSnapshot, type SnapshotStore } from '@forwork/client-store'
import type { SessionListState } from '@forwork/api-session-controller/client'
import type {
  SidebarRightOpenTab,
  SidebarRightTabActions,
} from '@forwork/client-ui-sidebar-right/client'
import type { SessionId } from '@forwork/session/types'
import type { DocumentProjectionView } from '@forwork/session-document/client'
import { fileAddressFor } from '@forwork/util-workspace-path'
import { OFFICE_KIND } from './definition.ts'

/** The documents one restore pass acts on, split by what the tab layer can show. */
export interface DocumentRestorePlan {
  /** Paths of open documents, in the projection's order. */
  readonly open: readonly string[]
  /** Paths recorded as missing; the Session keeps them open, so they are reported instead of opened. */
  readonly missing: readonly string[]
}

/**
 * Split the projection into the paths to open and the paths to report.
 *
 * Order is the projection's own `order`, so a restored tab sequence matches the
 * one the user left. A closed document is absent from the projection already; an
 * empty open set yields two empty lists. This mirrors the Host
 * `restoreSet` ordering rule; the two planes cannot share a value module, so the
 * Client splits the same projected entries itself.
 * @param view - the Session's `document` projection value.
 * @returns the ordered paths to open and the missing paths to report.
 */
export function restorePlan(view: DocumentProjectionView): DocumentRestorePlan {
  const open: string[] = []
  const missing: string[] = []
  for (const documentId of view.order) {
    const entry = view.documents[documentId]
    if (entry === undefined) continue
    if (entry.status === 'missing') missing.push(entry.path)
    else open.push(entry.path)
  }
  return { open, missing }
}

/** Restore bookkeeping for one window: which Sessions ran it and what they could not reopen. */
export interface OfficeRestoreState {
  /** Sessions whose open set this window has already restored. */
  readonly restored: readonly SessionId[]
  /** Missing document paths per Session, until the notice is dismissed. */
  readonly missing: Readonly<Record<string, readonly string[]>>
}

/** The restore operations one tab body may invoke. */
export interface OfficeRestoreActions {
  /**
   * Claim the one restore pass for a Session.
   * @param sessionId - the Session whose bodies are mounting.
   * @returns `true` for the first claim in this window, `false` for every later one.
   */
  claim(sessionId: SessionId): boolean
  /**
   * Record the documents a restore pass could not reopen.
   * @param sessionId - the Session whose open set was restored.
   * @param paths - the missing paths, in projection order; an empty list records nothing.
   */
  reportMissing(sessionId: SessionId, paths: readonly string[]): void
}

/** The plugin-owned restore record shared by every Office tab body. */
export interface OfficeRestore extends OfficeRestoreActions {
  /** Observable state the panel reads through the framework-bound hook. */
  readonly source: SnapshotStore<OfficeRestoreState>
  /**
   * Hide one Session's missing-document notice.
   * @param sessionId - the Session whose notice the user dismissed.
   */
  dismissMissing(sessionId: SessionId): void
}

/**
 * Create the restore record for one plugin instance.
 * @returns the observable state and the write set the panel and its hook share.
 */
export function createOfficeRestore(): OfficeRestore {
  const source = createSnapshotStore<OfficeRestoreState>({ restored: [], missing: {} })
  return {
    source,
    claim(sessionId) {
      const current = source.getSnapshot()
      if (current.restored.includes(sessionId)) return false
      source.set({ ...current, restored: [...current.restored, sessionId] })
      return true
    },
    reportMissing(sessionId, paths) {
      if (paths.length === 0) return
      const current = source.getSnapshot()
      source.set({ ...current, missing: { ...current.missing, [sessionId]: paths } })
    },
    dismissMissing(sessionId) {
      const current = source.getSnapshot()
      if (current.missing[sessionId] === undefined) return
      const missing: Record<string, readonly string[]> = {}
      for (const [id, paths] of Object.entries(current.missing)) {
        if (id !== sessionId) missing[id] = paths
      }
      source.set({ ...current, missing })
    },
  }
}

/** One restore pass: the projection, the tabs it reconciles, and where opens land. */
export interface DocumentRestorePassInput {
  /** Session whose open set this pass restores. */
  readonly sessionId: SessionId
  /** The live `document` projection value, absent until the Host publishes it. */
  readonly projection: DocumentProjectionView | undefined
  /** Open tab metadata, so a document already in a tab is not opened again. */
  readonly openTabs: readonly SidebarRightOpenTab[]
  /**
   * Open one missing document.
   * @param address - the file resource address to open.
   */
  readonly openResource: (address: string) => void
  /** The plugin-owned record deciding which pass runs. */
  readonly restore: OfficeRestoreActions
}

/**
 * Run one Session's restore pass if it still needs it.
 *
 * The pass is claimed before any tab opens, so later projection updates and
 * every other trigger are no-ops. A document that already has an Office tab is
 * skipped rather than revealed; one that is missing is reported, not opened.
 * @param input - the Session, its projection, the open tabs, and the open action.
 */
export function restoreDocuments(input: DocumentRestorePassInput): void {
  const { sessionId, projection, openTabs, openResource, restore } = input
  if (projection === undefined) return
  if (!restore.claim(sessionId)) return
  const plan = restorePlan(projection)
  const opened = new Set(openTabs
    .filter(tab => tab.sessionId === sessionId && tab.kind === OFFICE_KIND)
    .map(tab => tab.contentId))
  for (const path of plan.open) {
    const address = fileAddressFor(sessionId, undefined, path)
    if (opened.has(address)) continue
    openResource(address)
  }
  restore.reportMissing(sessionId, plan.missing)
}

/** One tab body's restore inputs. */
export interface DocumentRestoreInput {
  /** Session whose open set this body restores. */
  readonly sessionId: SessionId
  /** The live `document` projection value, absent until the Host publishes it. */
  readonly projection: DocumentProjectionView | undefined
  /** Open tab metadata, so a document already in a tab is not opened again. */
  readonly openTabs: readonly SidebarRightOpenTab[]
  /** This tab's own open action; it acts on the Session the tab is in. */
  readonly actions: Pick<SidebarRightTabActions, 'openResource'>
  /** The plugin-owned record deciding which pass runs. */
  readonly restore: OfficeRestoreActions
}

/**
 * Run this Session's one restore pass when its projection first arrives.
 * @param input - the Session, its projection, the open tabs, and the tab's actions.
 */
export function useDocumentRestore(input: DocumentRestoreInput): void {
  const { sessionId, projection, openTabs, actions, restore } = input
  useEffect(() => {
    restoreDocuments({
      sessionId,
      projection,
      openTabs,
      openResource: (address) => { actions.openResource(address) },
      restore,
    })
  }, [sessionId, projection, openTabs, actions, restore])
}

/** The Session-list and open-tab sources a Session activation pass reads. */
export interface SessionActivationRestoreInput {
  /** Host list snapshots carrying main-view retention and cached projection values. */
  readonly sessions: ObservableSnapshot<SessionListState>
  /** Open tab metadata across adopted Sessions. */
  readonly openTabs: ObservableSnapshot<readonly SidebarRightOpenTab[]>
  /** The plugin-owned record deciding which pass runs. */
  readonly restore: OfficeRestoreActions
  /**
   * Open one missing document in the Session that holds it.
   * @param sessionId - the Session whose open set is restored.
   * @param address - the file resource address to open.
   */
  readonly openResourceIn: (sessionId: SessionId, address: string) => void
}

/**
 * Restore every active Session's open set, including one with no Office tab.
 *
 * The main-view Session is the active one: the Session on screen has a retained
 * main view while the Conversation fills the main column. Its cached `document`
 * projection carries the durable set, so no Office body is needed to learn what
 * to reopen. Each notification reconciles every active Session; the shared
 * restore record makes all passes after the first a no-op. A Session whose set
 * is empty claims nothing, leaving a projection that arrives later able to run.
 * @param input - the Session list, open tabs, restore record, and per-Session open action.
 * @returns stops following Session activation.
 */
export function restoreOnSessionActivation(input: SessionActivationRestoreInput): () => void {
  const sync = (): void => {
    const list = input.sessions.getSnapshot()
    const openTabs = input.openTabs.getSnapshot()
    for (const sessionId of list.ids) {
      if ((list.byId[sessionId]?.retainedBy.mainView ?? 0) === 0) continue
      const projection = list.projectionsBySession[sessionId]?.values.document
      if (projection === undefined) continue
      const plan = restorePlan(projection)
      if (plan.open.length === 0 && plan.missing.length === 0) continue
      restoreDocuments({
        sessionId,
        projection,
        openTabs,
        openResource: (address) => { input.openResourceIn(sessionId, address) },
        restore: input.restore,
      })
    }
  }
  const unsubscribe = input.sessions.subscribe(sync)
  sync()
  return unsubscribe
}
