/**
 * Wire and seam vocabulary for the Office Exchange capability. This module is
 * the package's JSON-safe surface: it carries no Univer snapshot types, so the
 * generated Remote Client can import it without loading the editor core.
 * @module @forwork/office-exchange/types
 */

// Import the protocol module so the declaration at the end of this file
// augments its error map rather than defining an unrelated ambient module.
import type { DocumentCommandOperation } from '@forwork/session/types'
import type { JsonValue } from '@forwork/util-values'
import type {} from '@forwork/typert-protocol'

/** Office document formats handled by the Exchange layer. */
export type OfficeFormat = 'docx' | 'docm' | 'xlsx' | 'xlsm' | 'pptx' | 'doc' | 'xls' | 'ppt'

/** Formats the editable surface accepts and the agent can operate on. */
export type EditableOfficeFormat = 'docx' | 'docm' | 'xlsx' | 'xlsm'

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
  | 'formula-recalc'
  | 'embedded-object'
  | 'other'

/**
 * Unmodeled construct classes an in-place save carries verbatim. Detection
 * reports the classes present in a source package; presence alone never
 * refuses a save.
 */
export type PreservedConstructKind =
  | 'macro'
  | 'chart'
  | 'pivot-table'
  | 'excel-table'
  | 'defined-names'
  | 'calc-chain'
  | 'comment'
  | 'embedded-object'
  | 'drawing'
  | 'shape'
  | 'theme'
  | 'font-table'
  | 'web-settings'
  | 'custom-xml'
  | 'external-link'
  | 'field'
  | 'bookmark'
  | 'content-control'
  | 'document-properties'

/**
 * Construct classes whose meaning or anchors depend on the content an edit
 * would change. An edited unit intersecting one of these is a refusal.
 */
export type CoupledConstructKind =
  | 'tracked-change'
  | 'comment-range'
  | 'content-control-wrap'
  | 'field-owned-text'
  | 'rich-text-shared-string'
  | 'shared-formula'
  | 'array-formula'
  | 'ambiguous-alternate-content'

/** One coupled-construct occurrence located in a source part. */
export interface CoupledConstructMarker {
  /** Construct class whose anchor or meaning the edit would disturb. */
  readonly kind: CoupledConstructKind
  /** OPC part name, exactly as stored in the package. */
  readonly part: string
  /** Byte offset of the construct's start tag in the uncompressed part. */
  readonly offset: number
}

/**
 * Byte-based classification of a source package for an in-place save. It never
 * refuses on its own: the caller intersects {@link CoupledConstructMarker}
 * offsets with the edited unit spans.
 */
export interface PreservedConstructs {
  /** Construct classes carried verbatim; never a whole-file refusal list. */
  readonly preserved: ReadonlySet<PreservedConstructKind>
  /** Coupled-construct occurrences, sorted by part name then offset. */
  readonly coupled: readonly CoupledConstructMarker[]
}

/** Entry names affected one way by a save, with the exact count before truncation. */
export interface SaveDispositionGroup {
  /** Entries in this group, counted exactly even when `names` is truncated. */
  readonly count: number
  /** Sorted ASCII entry names, truncated to the disposition's `nameLimit`. */
  readonly names: readonly string[]
}

/**
 * What an in-place save carried, spliced, added, and removed. The counts are
 * exact; each group's `names` holds a sorted ASCII prefix of the affected
 * entries. A legacy regeneration cannot account for individual entries and
 * reports no group members.
 */
export interface SaveDisposition {
  /** Payload version. */
  readonly schemaVersion: 1
  /** Entries copied unchanged from the source package. */
  readonly verbatimCount: number
  /** Entries that received at least one span operation. */
  readonly spliced: SaveDispositionGroup
  /** New entries the save created. */
  readonly added: SaveDispositionGroup
  /** Source entries the save dropped. */
  readonly removed: SaveDispositionGroup
  /** Cap applied to each group's `names` array. */
  readonly nameLimit: number
}

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
  /**
   * Construct classes the source carries verbatim, sorted and deduplicated.
   * Absent from a Host that does not classify them; presence never blocks a save.
   */
  readonly preserved?: readonly PreservedConstructKind[]
  /**
   * Features a whole-file regeneration would drop. Presence alone never blocks
   * a save: the in-place writer refuses only an edit that intersects a coupled
   * construct.
   */
  readonly unsupported: readonly UnsupportedFeature[]
  /** Opaque content version the returned snapshot was read from. */
  readonly version: string
}

/** The version a successful save published and, when the Host reports it, what it changed. */
export interface OfficeDocumentSaveResult {
  /** Opaque content version of the written file. */
  readonly version: string
  /** Writer-computed account of the entries the save carried or changed. */
  readonly disposition?: SaveDisposition
}

/**
 * Who caused one save and, for an agent command, the turn it ran in.
 *
 * The client's user save passes `actor: 'user'`; an agent tool passes
 * `actor: 'agent'` with the open turn of the Session it is acting for.
 */
export interface DocumentSaveAttribution {
  /** Who authorized the save; the Session fact actor is `user` or `agent`. */
  readonly actor: 'user' | 'agent'
  /** Turn that caused the save, when a turn exists. */
  readonly turn?: number
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

/**
 * One operation a Host tool dispatches to the live editor unit through the
 * `document/command` bridge. A read asks for the current unit; an edit applies
 * the same structured steps a `document/edit` fact stores.
 */
export type OfficeCommandOperation = DocumentCommandOperation

/**
 * The live editor's answer to one command.
 *
 * `contentVersion` is the document's numeric content version: it starts at 0
 * when the document opens and advances by one for each accepted save. A
 * `read` or `applied` answer reports the live value so the Host can hand the
 * next command a current `fromVersion`.
 */
export type OfficeCommandResult =
  | {
    /** The live unit was read without changing it. */
    readonly status: 'read'
    /** Current editable content. */
    readonly snapshot: OfficeDocumentSnapshot
    /** Live document content version. */
    readonly contentVersion: number
  }
  | {
    /** The operation was applied to the live unit; nothing is saved yet. */
    readonly status: 'applied'
    /** Editable content after the operation. */
    readonly snapshot: OfficeDocumentSnapshot
    /** Live document content version after the operation. */
    readonly contentVersion: number
  }
  | {
    /** The live unit holds unsaved user edits, so an agent save is refused. */
    readonly status: 'pending-edits'
  }
  | {
    /** The live unit's version differs from the command's `fromVersion`. */
    readonly status: 'conflict'
    /** Live document content version. */
    readonly currentVersion: number
  }
  | {
    /** Applying or saving the operation would drop unsupported content. */
    readonly status: 'unsupported-feature'
    /** Features the editable model cannot preserve. */
    readonly features: readonly UnsupportedFeature[]
  }
  | {
    /** No live editor instance matched the document. */
    readonly status: 'document-not-open'
  }

/**
 * The Host-side outcome of one command: the editor's answer or the elapsed
 * timeout. A timeout writes nothing and is reported as the tool's `timeout`
 * failure rather than a live-editor status. The broker answers
 * `document-not-open` immediately when the document has no registered live
 * target, so only a registered target that stays silent reaches the timeout.
 */
export type OfficeCommandOutcome = OfficeCommandResult | {
  /** No result arrived before the configured timeout. */
  readonly status: 'timeout'
}

declare module '@forwork/typert-protocol' {
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
    /** The file does not exist or is not a regular file. */
    'office-document/not-found': { readonly path: string }
    /** The file changed after the caller read it, so the guarded save was refused. */
    'office-document/conflict': { readonly path: string }
  }
}
