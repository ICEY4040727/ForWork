/**
 * Wire and seam vocabulary for the Office Exchange capability. This module is
 * the package's JSON-safe surface: it carries no Univer snapshot types, so the
 * generated Remote Client can import it without loading the editor core.
 * @module @deepseek-ai/dsh-office-exchange/types
 */

// Import the protocol module so the declaration at the end of this file
// augments its error map rather than defining an unrelated ambient module.
import type { JsonValue } from '@deepseek-ai/dsh-util-values'
import type {} from '@deepseek-ai/dsh-typert-protocol'

/** Office document formats handled by the Exchange layer. */
export type OfficeFormat = 'docx' | 'xlsx' | 'pptx' | 'doc' | 'xls' | 'ppt'

/** Formats the editable surface accepts and the agent can operate on. */
export type EditableOfficeFormat = 'docx' | 'xlsx'

/** Legacy binary formats converted to an editable format before opening. */
export type LegacyOfficeFormat = 'doc' | 'xls' | 'ppt'

/**
 * Kinds of editable unit shared by the editor and the agent. `UnitSnapshot`
 * narrows with these literal kinds; the alias remains exported for consumers
 * that name the discriminant.
 */
export type UnitSnapshotKind = 'workbook' | 'document'

/** A feature present in a source file that the editable model cannot preserve. */
export interface UnsupportedFeature {
  /** Coarse classification of the feature. */
  readonly kind: UnsupportedFeatureKind
  /** Human-readable location or context, when known. */
  readonly detail?: string
}

/** Known feature classes the editable model cannot round-trip losslessly. */
export type UnsupportedFeatureKind =
  | 'macro'
  | 'tracked-changes'
  | 'comment'
  | 'chart'
  | 'drawing'
  | 'shape'
  | 'header-footer'
  | 'formula-recalc'
  | 'embedded-object'
  | 'other'

/**
 * A document snapshot on the Remote wire. Univer's native snapshot interfaces
 * admit unconstrained custom fields, which a JSON boundary cannot carry, so the
 * Remote projects and restores the payload as lossless JSON around the kind
 * discriminant. The Host validates this value before serializing it.
 */
export type OfficeDocumentSnapshot =
  | {
    /** A spreadsheet unit. */
    readonly kind: 'workbook'
    /** JSON projection of an `IWorkbookData` snapshot. */
    readonly data: JsonValue
  }
  | {
    /** A word-processing unit. */
    readonly kind: 'document'
    /** JSON projection of an `IDocumentData` snapshot. */
    readonly data: JsonValue
  }

/** One document read as the Client receives it from the `officeDocument` Remote. */
export interface OfficeDocumentReadResult {
  /** Detected source format. */
  readonly format: OfficeFormat
  /** Editable content at {@link version}. */
  readonly snapshot: OfficeDocumentSnapshot
  /** Features the editable model cannot preserve on save. */
  readonly unsupported: readonly UnsupportedFeature[]
  /** Opaque content version the returned snapshot was read from. */
  readonly version: string
}

/** The version a successful save published. */
export interface OfficeDocumentSaveResult {
  /** Opaque content version of the written file. */
  readonly version: string
}

/**
 * One recovery draft persisted for an open document. The snapshot is bound to
 * the file version it was captured from, so a restart can tell a current draft
 * from one based on an older file.
 */
export interface OfficeDocumentDraft {
  /** Caller-owned identity of the open document. */
  readonly documentId: string
  /** Document path the draft was captured for, recorded as the caller supplied it. */
  readonly path: string
  /** File version the snapshot was captured from. */
  readonly contentVersion: string
  /** Captured editable content. */
  readonly snapshot: OfficeDocumentSnapshot
  /** ISO-8601 time the draft was written. */
  readonly savedAt: string
}

/** The timestamp a successful draft write persisted. */
export interface OfficeDocumentDraftWriteResult {
  /** ISO-8601 time the draft was written. */
  readonly savedAt: string
}

declare module '@deepseek-ai/dsh-typert-protocol' {
  interface RemoteErrorDetailsMap {
    /** The extension or file content is not an editable Office format. */
    'office-document/unsupported-format': { readonly path: string }
    /** The source contains a feature an edit-and-save would drop. */
    'office-document/unsupported-feature': { readonly path: string }
    /** The path resolves outside the Session's workspace root. */
    'office-document/out-of-scope': { readonly path: string }
    /** A configured byte bound was exceeded. */
    'office-document/bound-exceeded': { readonly path: string }
    /** Parsing, serialization, or legacy conversion failed. */
    'office-document/conversion-failed': { readonly path: string }
    /** The file changed after the caller read it, so the guarded save was refused. */
    'office-document/conflict': { readonly path: string }
  }
}
