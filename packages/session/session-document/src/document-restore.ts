/**
 * The open-document set a client restores after the editor starts again.
 *
 * Restore reads the durable `document` projection, never the filesystem and
 * never a per-window list: the projection is the single authority for which
 * documents were open, in which order, and at which path after any rebind.
 *
 * A `missing` document is part of the restore set. The projection keeps it
 * open with its identity and history (FR-015), so dropping it here would lose
 * the fact that the Session still holds the document; the caller opens a tab
 * only for `open` entries and reports the missing ones instead.
 *
 * @module @forwork/session-document/document-restore
 */

import type { DocumentProjectionEntry, DocumentProjectionView } from './types.ts'

/**
 * One document the client must restore.
 *
 * `status` is carried so a caller can split the set without re-reading the
 * projection: an `open` entry becomes a tab, a `missing` entry becomes a
 * report naming its path.
 */
export type DocumentRestoreEntry = Pick<
  DocumentProjectionEntry,
  'documentId' | 'path' | 'format' | 'status'
>

/**
 * Expand the projection into the ordered restore set.
 *
 * Order is the projection's own `order`, which preserves each document's
 * position across saves, edits, rebinds, and missing observations. A closed
 * document is absent from the projection, so it is absent here; an empty open
 * set yields an empty list.
 *
 * @param state - the folded `document` projection, or its client view.
 * @returns one entry per open document, in open order, with `missing` entries retained.
 */
export function restoreSet(state: DocumentProjectionView): readonly DocumentRestoreEntry[] {
  const entries: DocumentRestoreEntry[] = []
  for (const documentId of state.order) {
    const entry = state.documents[documentId]
    if (entry === undefined) continue
    entries.push({
      documentId: entry.documentId,
      path: entry.path,
      format: entry.format,
      status: entry.status,
    })
  }
  return entries
}
