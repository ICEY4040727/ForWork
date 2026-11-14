/**
 * Where each converted tab came from.
 *
 * Opening a legacy `.doc`, `.xls`, or `.ppt` converts it and opens the copy in a
 * new tab, so the converted body must still name the file the user opened
 * and say the original is unchanged. The Sidebar's resource-open owner data
 * carries no plugin field, and the tab is replaced in the same step, so the
 * plugin keeps one record per converted content id: written before the open and
 * read by the converted body, which may remount through hiding and Session
 * switches. Records live only for this window's session, keyed by the address
 * the converted tab was opened under.
 */
import { createSnapshotStore, type SnapshotStore } from '@deepseek-ai/dsh-client-store'

/** One converted tab's notice. */
export interface OfficeLegacyNotice {
  /** The workspace path of the legacy source the user opened. */
  readonly source: string
  /** The user dismissed the notice for this converted copy. */
  readonly dismissed: boolean
}

/** Every converted tab's notice, keyed by the converted tab's content id. */
export type OfficeLegacyNoticeSnapshot = Readonly<Record<string, OfficeLegacyNotice>>

/** The plugin-owned notice store shared by the converting and converted bodies. */
export interface OfficeLegacyNotices {
  /** Observable records the converted body reads through the framework-bound hook. */
  readonly source: SnapshotStore<OfficeLegacyNoticeSnapshot>
  /**
   * Record the source a converted copy was made from.
   * @param contentId - the converted tab's address.
   * @param sourcePath - the legacy source path to name in the notice.
   */
  remember(contentId: string, sourcePath: string): void
  /**
   * Hide the notice for one converted copy.
   * @param contentId - the converted tab's address; an address with no record is unchanged.
   */
  dismiss(contentId: string): void
  /**
   * Drop one converted copy's record once its tab is gone.
   * @param contentId - the closed tab's address.
   */
  forget(contentId: string): void
}

/**
 * Create the notice store for one plugin instance.
 * @returns the observable records and their write set.
 */
export function createOfficeLegacyNotices(): OfficeLegacyNotices {
  const source = createSnapshotStore<OfficeLegacyNoticeSnapshot>({})
  const write = (contentId: string, notice: OfficeLegacyNotice): void => {
    source.set({ ...source.getSnapshot(), [contentId]: notice })
  }
  return {
    source,
    remember(contentId, sourcePath) {
      write(contentId, { source: sourcePath, dismissed: false })
    },
    dismiss(contentId) {
      const current = source.getSnapshot()[contentId]
      if (current === undefined) return
      write(contentId, { source: current.source, dismissed: true })
    },
    forget(contentId) {
      const current = source.getSnapshot()
      if (current[contentId] === undefined) return
      const next: Record<string, OfficeLegacyNotice> = {}
      for (const [id, notice] of Object.entries(current)) {
        if (id !== contentId) next[id] = notice
      }
      source.set(next)
    },
  }
}
