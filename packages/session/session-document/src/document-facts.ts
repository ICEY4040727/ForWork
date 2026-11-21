/**
 * Durable `document/*` fact writers for one Session.
 *
 * Every helper appends exactly one ordinary Session event and returns the
 * committed event. The log is append-only: no helper rewrites or removes an
 * earlier fact, and a caller that must report a later observation appends a new
 * one. The `document` projection is the only reader that decides what the open
 * set means.
 *
 * @module @forwork/session-document/document-facts
 */

import type {
  DocumentCloseEventData,
  DocumentFormat,
  DocumentId,
  DocumentMissingEventData,
  DocumentOpenEventData,
  DocumentRebindEventData,
  DocumentSaveDisposition,
  DocumentSaveEventData,
  Session,
  SessionEvent,
} from '@forwork/session'

/** One document that became open in a Session. */
export interface DocumentOpenFact {
  /** Stable identity of the opened document. */
  readonly documentId: DocumentId
  /** Path the document was opened from; a locator, never identity. */
  readonly path: string
  /** Detected editable format. */
  readonly format: DocumentFormat
  /** Document version established by the open. */
  readonly contentVersion: number
}

/** One accepted save whose in-memory changes stay inside the editor. */
export interface DocumentSaveFact {
  /** Stable identity of the saved document. */
  readonly documentId: DocumentId
  /** Who authorized the save. */
  readonly actor: 'user' | 'agent'
  /** Document version the save published. */
  readonly contentVersion: number
  /** Turn that caused the save, when a turn caused it. */
  readonly turn?: number
  /** Preservation summary produced by the writer; forwarded unchanged. */
  readonly disposition?: DocumentSaveDisposition
}

/** One document that left the open set. */
export interface DocumentCloseFact {
  /** Stable identity of the closed document. */
  readonly documentId: DocumentId
  /** Producer-supplied closure reason; the projection does not interpret it. */
  readonly reason: string
}

/** One open document whose file moved or was renamed. */
export interface DocumentRebindFact {
  /** Stable identity retained across the move. */
  readonly documentId: DocumentId
  /** Path before the move. */
  readonly fromPath: string
  /** Path after the move. */
  readonly toPath: string
}

/** One open document whose file no longer resolves. */
export interface DocumentMissingFact {
  /** Stable identity retained while the file is unavailable. */
  readonly documentId: DocumentId
  /** Path that no longer resolves. */
  readonly path: string
}

/**
 * Append `document/open` for a document that became open.
 * @param session - Session that owns the document.
 * @param fact - identity, path, format, and starting version.
 * @returns the committed `document/open` event.
 */
export function recordDocumentOpen(session: Session, fact: DocumentOpenFact): SessionEvent<'document/open'> {
  return session.append('document/open', {
    documentId: fact.documentId,
    path: fact.path,
    format: fact.format,
    contentVersion: fact.contentVersion,
  } satisfies DocumentOpenEventData)
}

/**
 * Append `document/save` for an accepted save with no exposed operation steps.
 * @param session - Session that owns the document.
 * @param fact - identity, author, published content version, causing turn, and optional preservation summary.
 * @returns the committed `document/save` event.
 */
export function recordDocumentSave(session: Session, fact: DocumentSaveFact): SessionEvent<'document/save'> {
  return session.append('document/save', {
    documentId: fact.documentId,
    actor: fact.actor,
    contentVersion: fact.contentVersion,
    ...fact.turn === undefined ? {} : { turn: fact.turn },
    ...fact.disposition === undefined ? {} : { disposition: fact.disposition },
  } satisfies DocumentSaveEventData)
}

/**
 * Append `document/close` for a document that left the open set.
 * @param session - Session that owns the document.
 * @param fact - identity and producer-supplied reason.
 * @returns the committed `document/close` event.
 */
export function recordDocumentClose(session: Session, fact: DocumentCloseFact): SessionEvent<'document/close'> {
  return session.append('document/close', {
    documentId: fact.documentId,
    reason: fact.reason,
  } satisfies DocumentCloseEventData)
}

/**
 * Append `document/rebind` for an open document whose locator changed.
 * @param session - Session that owns the document.
 * @param fact - retained identity and both paths.
 * @returns the committed `document/rebind` event.
 */
export function recordDocumentRebind(session: Session, fact: DocumentRebindFact): SessionEvent<'document/rebind'> {
  return session.append('document/rebind', {
    documentId: fact.documentId,
    fromPath: fact.fromPath,
    toPath: fact.toPath,
  } satisfies DocumentRebindEventData)
}

/**
 * Append `document/missing` for an open document whose file no longer resolves.
 * @param session - Session that owns the document.
 * @param fact - retained identity and the unresolved path.
 * @returns the committed `document/missing` event.
 */
export function recordDocumentMissing(
  session: Session,
  fact: DocumentMissingFact,
): SessionEvent<'document/missing'> {
  return session.append('document/missing', {
    documentId: fact.documentId,
    path: fact.path,
  } satisfies DocumentMissingEventData)
}
