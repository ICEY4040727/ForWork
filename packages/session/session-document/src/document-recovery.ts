/**
 * How a stored recovery draft relates to the file an editor just read.
 *
 * The draft store (the `officeDocument` Remote) and the client capture are
 * separate owners. This module owns only the decision, so both sides apply one
 * rule: a draft is offered when it was captured from the version the file
 * currently holds, labelled stale otherwise, and never applied without the
 * user's explicit choice.
 *
 * @module @forwork/session-document/document-recovery
 */

/** What the editor may do with one stored draft. */
export type DocumentRecoveryDecision = 'none' | 'offer' | 'stale'

/**
 * The part of a stored draft the decision reads.
 *
 * This mirrors the `officeDocument` wire draft structurally, so the decision
 * depends on no document-format package. `contentVersion` is the file version
 * the draft was captured from, as the filesystem seam reported it.
 */
export interface DocumentRecoveryDraft {
  /** File version the draft's snapshot was captured from. */
  readonly contentVersion: string
}

/**
 * Decide whether the editor offers a stored draft, labels it stale, or ignores it.
 *
 * A draft whose captured version equals the file's current version restores
 * exactly what the user left unsaved; any other version may be built on
 * content that changed since, so it is offered only behind an explicit
 * "recover anyway" choice.
 *
 * @param draft - the stored draft, or `null`/`undefined` when the document has none.
 * @param currentVersion - the version the successful read observed for the file.
 * @returns `none` without a draft, `offer` when the versions match, `stale` otherwise.
 */
export function recoveryDecision(
  draft: DocumentRecoveryDraft | null | undefined,
  currentVersion: string,
): DocumentRecoveryDecision {
  if (draft === null || draft === undefined) return 'none'
  return draft.contentVersion === currentVersion ? 'offer' : 'stale'
}
