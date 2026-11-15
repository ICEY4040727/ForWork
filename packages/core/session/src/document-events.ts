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
 *
 * `document/command` is a correlated request rather than an accepted change.
 * Its loss cannot corrupt the reconstructed document content: the accepted
 * result is `document/edit`, which carries the operation and its inverse. It is
 * still a known required event so a reader cannot resume a Session that holds
 * an unanswered request without understanding that type.
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

/**
 * One row insertion or deletion in a spreadsheet. `row` is the first affected
 * zero-based index; `count` rows move or are removed as one step so undo costs
 * do not grow with the number of edited cells.
 */
export interface DocumentRowEdit {
  /** Sheet identity inside the workbook snapshot. */
  readonly sheetId: string
  /** Zero-based index of the first affected row. */
  readonly row: number
  /** Number of rows inserted or deleted. */
  readonly count: number
  /** Whether the rows are inserted before `row` or deleted from `row`. */
  readonly action: 'insert' | 'delete'
}

/** One column insertion or deletion in a spreadsheet. */
export interface DocumentColumnEdit {
  /** Sheet identity inside the workbook snapshot. */
  readonly sheetId: string
  /** Zero-based index of the first affected column. */
  readonly column: number
  /** Number of columns inserted or deleted. */
  readonly count: number
  /** Whether the columns are inserted before `column` or deleted from `column`. */
  readonly action: 'insert' | 'delete'
}

/** One formula assignment to a spreadsheet cell. */
export interface DocumentFormulaEdit {
  /** Sheet identity inside the workbook snapshot. */
  readonly sheetId: string
  /** Zero-based row index in that sheet. */
  readonly row: number
  /** Zero-based column index in that sheet. */
  readonly column: number
  /** Formula text after the edit, or `null` to clear the formula. */
  readonly formula: string | null
}

/** One style assignment over a rectangular spreadsheet range. */
export interface DocumentStyleEdit {
  /** Sheet identity inside the workbook snapshot. */
  readonly sheetId: string
  /** Zero-based first row of the range, inclusive. */
  readonly startRow: number
  /** Zero-based first column of the range, inclusive. */
  readonly startColumn: number
  /** Zero-based last row of the range, inclusive. */
  readonly endRow: number
  /** Zero-based last column of the range, inclusive. */
  readonly endColumn: number
  /** JSON style object applied to every cell in the range. */
  readonly style: JsonValue
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
  | ({
    /** Marks a spreadsheet row insertion or deletion. */
    readonly kind: 'row'
  } & DocumentRowEdit)
  | ({
    /** Marks a spreadsheet column insertion or deletion. */
    readonly kind: 'column'
  } & DocumentColumnEdit)
  | ({
    /** Marks a spreadsheet formula assignment. */
    readonly kind: 'formula'
  } & DocumentFormulaEdit)
  | ({
    /** Marks a spreadsheet style assignment over a range. */
    readonly kind: 'style'
  } & DocumentStyleEdit)

/**
 * Structured, JSON-serializable content change stored by `document/edit`.
 * `document/edit` carries both this operation and its inverse instead of a raw
 * unit snapshot, so document undo is reconstructable from the Session log.
 *
 * The step union grows additively: a row or column change stays one step
 * instead of expanding into one step per affected cell.
 */
export type DocumentEditOperation = readonly DocumentEditStep[]

/**
 * One operation dispatched to the live editor unit.
 *
 * A read requests the current unit without changing it; an edit carries the
 * same steps a `document/edit` fact stores. The command never carries a unit
 * snapshot, so a persisted request cannot duplicate document content.
 */
export type DocumentCommandOperation =
  | {
    /** Requests the live unit's current state without editing it. */
    readonly kind: 'read'
  }
  | {
    /** Applies edit steps to the live unit. */
    readonly kind: 'edit'
    /** Steps to apply, in order. */
    readonly steps: DocumentEditOperation
  }

/**
 * Payload of `document/command`: a correlated request for the live editor.
 *
 * The Host appends one command and waits for the matching
 * `officeDocument.completeCommand` result. The command is a request, not an
 * accepted change: only a successful save appends `document/edit`, so a
 * command that times out or is refused leaves no content fact behind.
 */
export interface DocumentCommandEventData {
  /** Correlation id chosen by the Host for this request. */
  readonly requestId: string
  /** Stable identity of the open document the request targets. */
  readonly documentId: DocumentId
  /** Operation the live editor applies or answers. */
  readonly operation: DocumentCommandOperation
  /** Document content version the caller believes the live unit holds. */
  readonly fromVersion: number
}

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
    /**
     * A correlated request for the live editor unit: apply edit steps or read
     * the current state. The command is not an accepted change; only a
     * successful save appends `document/edit`.
     */
    'document/command': DocumentCommandEventData
  }
}
