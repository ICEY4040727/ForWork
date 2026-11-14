/**
 * Stable document identity bindings for open office documents.
 *
 * A binding pairs the Session-local {@link DocumentId} with the current
 * locator and tracks whether that locator still resolves. Rename, move, and a
 * temporary missing file change the binding, never the identity: every
 * function here returns a new binding and preserves `documentId`.
 *
 * The module is pure. It performs no filesystem access; callers report path
 * existence and observation times.
 *
 * @module @deepseek-ai/dsh-session-document/document-identity
 */

import type { DocumentId } from '@deepseek-ai/dsh-session'

/**
 * The current locator of one open document and its availability.
 *
 * `path` is an opaque locator owned by the filesystem seam, not identity.
 * `reboundAt` records only the latest path change; a missing binding keeps
 * the document and its history open.
 */
export interface DocumentIdentityBinding {
  /** Stable identity retained across every rebind. */
  readonly documentId: DocumentId
  /** Current path of the file; never used as identity. */
  readonly path: string
  /** Whether the current path failed to resolve at the last observation. */
  readonly missing: boolean
  /** Observation time of the latest path change, when the path has changed. */
  readonly reboundAt?: number
}

/**
 * Create the initial binding for a newly opened document.
 * @param documentId - stable Session-local document identity.
 * @param path - path the file was opened from.
 * @returns a binding whose path resolves and that has no rebind timestamp.
 */
export function createDocumentIdentity(documentId: DocumentId, path: string): DocumentIdentityBinding {
  return { documentId, path, missing: false }
}

/**
 * Rebind a document to a new path after a rename or move.
 *
 * The identity and history are preserved; only the locator, its availability,
 * and the rebind timestamp change. A caller that cannot observe the new path
 * passes `exists: false`, and the binding stays missing until a later
 * observation resolves it.
 *
 * @param binding - the binding to rebind.
 * @param toPath - path after the rename or move.
 * @param at - observation time of the path change.
 * @param exists - whether the caller observed the new path resolving.
 * @returns a new binding with the same document id and the new locator.
 */
export function rebindDocument(
  binding: DocumentIdentityBinding,
  toPath: string,
  at: number,
  exists: boolean,
): DocumentIdentityBinding {
  return { documentId: binding.documentId, path: toPath, missing: !exists, reboundAt: at }
}

/**
 * Mark a document's current path as no longer resolving.
 *
 * The document stays open: identity, path, and any rebind history are kept so
 * a later observation can resolve the same id again.
 *
 * @param binding - the binding whose path failed to resolve.
 * @param _at - observation time of the missing fact. A binding records only
 * path changes, so this time is accepted for call-site symmetry with the
 * `document/missing` fact and is not stored.
 * @returns a new binding with `missing: true` and every other field preserved.
 */
export function markDocumentMissing(binding: DocumentIdentityBinding, _at: number): DocumentIdentityBinding {
  return { ...binding, missing: true }
}

/**
 * Find the binding for a document id.
 * @param bindings - bindings to search, in registry order.
 * @param documentId - identity to resolve.
 * @returns the matching binding, or `undefined` when no binding carries that id.
 */
export function resolveDocumentBinding(
  bindings: readonly DocumentIdentityBinding[],
  documentId: DocumentId,
): DocumentIdentityBinding | undefined {
  return bindings.find(binding => binding.documentId === documentId)
}
