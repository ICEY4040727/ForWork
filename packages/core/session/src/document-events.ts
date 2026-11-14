/**
 * Durable `document/*` Session facts: the open-document vocabulary consumed by
 * the `document` projection, restore, and the document tools.
 *
 * Every member is an ordinary required log event. None carries the envelope's
 * `ignorable` marker: the projection cannot rebuild the open-document set
 * without these facts, and `Session.append` writes no marker for
 * `SessionEventMap` members. Adding an ordinary event type remains a
 * same-version change because a reader that predates these types refuses the
 * Session at admission against `KNOWN_SESSION_EVENT_TYPES` instead of silently
 * dropping required facts, so `SESSION_FORMAT_VERSION` stays unchanged.
 */

import { brandString, type Branded } from '@deepseek-ai/dsh-brand'
import type { JsonValue } from '@deepseek-ai/dsh-util-values'

/** Stable identity of one open document, unique within its authorizing Session. */
export type DocumentId = Branded<'DocumentId'>

/**
 * Brand a Session-local document identity.
 * @param id - opaque document id issued by the document identity registry.
 * @returns the same string with the document-id brand.
 */
export function DocumentId(id: string): DocumentId {
  return brandString<DocumentId>(id)
}

/** Who caused one document fact. Content changes are made by `user` or `agent`; `system` covers lifecycle actions. */
export type DocumentEventActor = 'user' | 'agent' | 'system'

/** Format of an open document. Legacy `.doc`/`.xls`/`.ppt` files are converted before opening. */
export type DocumentFormat = 'docx' | 'xlsx' | 'pptx'

/**
 * One spreadsheet cell value after an edit. `value` is the JSON form of the
 * cell data, never a whole unit snapshot.
 */
export interface DocumentCellEdit {
  /** Sheet identity inside the workbook snapshot. */
  readonly sheetId: string
  /** Zero-based row index in that sheet. */
  readonly row: number
  /** Zero-based column index in that sheet. */
  readonly column: number
  /** JSON cell value after the edit. */
  readonly value: JsonValue
}

/** One paragraph replacement in a text document. */
export interface DocumentParagraphEdit {
  /** Zero-based paragraph index inside the document snapshot. */
  readonly index: number
  /** Paragraph text after the edit. */
  readonly text: string
}

/** One primitive change in a document edit or its inverse. */
export type DocumentEditStep =
  | ({
    /** Marks a spreadsheet cell replacement. */
    readonly kind: 'cell'
  } & DocumentCellEdit)
  | ({
    /** Marks a paragraph replacement. */
    readonly kind: 'paragraph'
  } & DocumentParagraphEdit)

/**
 * Structured, JSON-serializable content change stored by `document/edit`.
 * `document/edit` carries both this operation and its inverse instead of a raw
 * unit snapshot, so document undo is reconstructable from the Session log.
 */
export type DocumentEditOperation = readonly DocumentEditStep[]

/** Payload of `document/open`: a document became open in the authorizing Session. */
export interface DocumentOpenEventData {
  /** Stable identity of the opened document. */
  readonly documentId: DocumentId
  /** Absolute path of the file at open time; a locator that rebinding may replace, not identity. */
  readonly path: string
  /** Detected format of the file. */
  readonly format: DocumentFormat
  /** Document version established by the open; no accepted edit has advanced it yet. */
  readonly contentVersion: number
  /** Turn that opened the document, when a turn caused the open. */
  readonly turn?: number
}

/**
 * Payload of `document/edit`: one accepted, attributed change. The event is
 * appended only after the version guard accepts the write, and `toVersion` is
 * always `fromVersion + 1`.
 */
export interface DocumentEditEventData {
  /** Stable identity of the edited document. */
  readonly documentId: DocumentId
  /** Who caused the change. */
  readonly actor: DocumentEventActor
  /** Document version the change was applied to. */
  readonly fromVersion: number
  /** Document version after the change: `fromVersion + 1`. */
  readonly toVersion: number
  /** Structured changes that were applied. */
  readonly operation: DocumentEditOperation
  /** Structured changes that exactly reverse `operation`. */
  readonly inverse: DocumentEditOperation
  /** Turn that made the change, when a turn caused it. */
  readonly turn?: number
}

/** Payload of `document/rebind`: the file moved or was renamed while the document stayed open. */
export interface DocumentRebindEventData {
  /** Stable identity retained across the move. */
  readonly documentId: DocumentId
  /** Absolute path before the move. */
  readonly fromPath: string
  /** Absolute path after the move. */
  readonly toPath: string
  /** Turn that observed the move, when a turn caused it. */
  readonly turn?: number
}

/** Payload of `document/missing`: the file no longer resolves; the document and its history stay open. */
export interface DocumentMissingEventData {
  /** Stable identity retained while the file is unavailable. */
  readonly documentId: DocumentId
  /** Absolute path that no longer resolves. */
  readonly path: string
  /** Turn that observed the missing file, when a turn caused it. */
  readonly turn?: number
}

/** Payload of `document/close`: the document left the open set. History is not deleted. */
export interface DocumentCloseEventData {
  /** Stable identity of the closed document. */
  readonly documentId: DocumentId
  /** Producer-supplied closure reason. The projection removes the document and does not interpret the value. */
  readonly reason: string
  /** Turn that closed the document, when a turn caused it. */
  readonly turn?: number
}

declare module '@deepseek-ai/dsh-session/types' {
  interface SessionEventMap {
    /**
     * A document became open in this Session. The projection records the
     * document id, path, format, and starting content version; later facts
     * reference the same id.
     */
    'document/open': DocumentOpenEventData
    /**
     * One accepted document change with its actor, version transition, forward
     * operation, and exact inverse. Raw unit snapshots are never persisted.
     */
    'document/edit': DocumentEditEventData
    /**
     * The open document's file moved or was renamed. The document id and all
     * earlier history are preserved.
     */
    'document/rebind': DocumentRebindEventData
    /**
     * The open document's file no longer resolves. The document stays open and
     * keeps its history until it is closed.
     */
    'document/missing': DocumentMissingEventData
    /**
     * The document left the open set. The fact does not delete earlier history.
     */
    'document/close': DocumentCloseEventData
  }
}
