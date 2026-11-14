/**
 * Attributable, reversible document changes and the version guard.
 *
 * Every accepted edit becomes a {@link ChangeRecord} carrying its actor, the
 * document version transition, and the inverse operation that restores the
 * prior content. This module is the guard: it owns no storage and reads no
 * file. A caller passes the live document version as `currentVersion`, and the
 * change is accepted only when its `fromVersion` matches that value.
 *
 * All functions are pure and return new values; inputs are never mutated.
 *
 * @module @deepseek-ai/dsh-session-document/document-changes
 */

import type { DocumentEditOperation, DocumentId } from '@deepseek-ai/dsh-session'

/** Who caused one content change. Lifecycle facts may be `system`; content changes never are. */
export type ChangeActor = 'user' | 'agent'

/**
 * One attributable, reversible edit accepted against a document version.
 *
 * `operation` and `inverse` store structured edit steps, never raw unit
 * snapshots, so the Session log can reconstruct document undo.
 */
export interface ChangeRecord {
  /** Deterministic change identity, unique within the document. */
  readonly changeId: string
  /** Document the change belongs to. */
  readonly documentId: DocumentId
  /** Who made the change. */
  readonly actor: ChangeActor
  /** Document version the change was applied to. */
  readonly fromVersion: number
  /** Document version after the change: `fromVersion + 1`. */
  readonly toVersion: number
  /** Structured steps that were applied. */
  readonly operation: DocumentEditOperation
  /** Structured steps that exactly reverse `operation`. */
  readonly inverse: DocumentEditOperation
  /** Observation time of the accepted change. */
  readonly at: number
}

/** Stable failure codes consumers branch on instead of message text. */
export type OfficeDocumentErrorCode =
  | 'stale-version'
  | 'invalid-transition'
  | 'invalid-operation'

/** A refused document change with a stable classification. */
export class OfficeDocumentError extends Error {
  /**
   * @param code - stable classification for branching consumers.
   * @param message - localized, actionable explanation.
   */
  constructor(
    readonly code: OfficeDocumentErrorCode,
    message: string,
  ) {
    super(message)
    this.name = 'OfficeDocumentError'
  }
}

/** Request to apply one change against a live document version. */
export interface ApplyChangeInput {
  /** Document the change belongs to. */
  readonly documentId: DocumentId
  /** Who made the change. */
  readonly actor: ChangeActor
  /** Actual version of the live document at the time of the request. */
  readonly currentVersion: number
  /** Version the caller believes it is editing; it must equal `currentVersion`. */
  readonly fromVersion: number
  /** Structured steps to apply. */
  readonly operation: DocumentEditOperation
  /** Structured steps that exactly reverse `operation`. */
  readonly inverse: DocumentEditOperation
  /** Observation time of the change. */
  readonly at: number
  /** Explicit change identity; defaults to `documentId:toVersion`. */
  readonly changeId?: string
}

/** An accepted change and the document version it produced. */
export interface AppliedChange {
  /** The accepted record, ready for the `document/edit` fact. */
  readonly record: ChangeRecord
  /** Document version after the change. */
  readonly version: number
}

/** The document-level undo primitive: the inverse steps and the version to restore. */
export interface ChangeUndo {
  /** Structured steps that exactly reverse the accepted change. */
  readonly operation: DocumentEditOperation
  /** Document version the change was applied to. */
  readonly version: number
}

/**
 * Refuse a version that cannot describe a real document.
 * @param version - candidate document version.
 * @param label - field name used in the failure message.
 * @throws OfficeDocumentError with code `invalid-transition` when the version
 * is not a non-negative safe integer.
 */
function assertVersion(version: number, label: string): void {
  if (!Number.isSafeInteger(version) || version < 0) {
    throw new OfficeDocumentError(
      'invalid-transition',
      `${label} ${version} is not a non-negative safe integer.`,
    )
  }
}

/**
 * Apply one attributed change when it matches the live document version.
 *
 * The version guard is pure: this module does not read storage, so the caller
 * supplies `currentVersion` from the document it is editing. A mismatch is a
 * conflict and is rejected as `stale-version`; nothing is overwritten.
 *
 * @param input - the change, its actor, and the live document version.
 * @returns the accepted record and the resulting document version.
 * @throws OfficeDocumentError with code `stale-version` when `fromVersion`
 * differs from `currentVersion`, `invalid-operation` when either the operation
 * or its inverse is empty, or `invalid-transition` when a version cannot be
 * advanced safely.
 */
export function applyChange(input: ApplyChangeInput): AppliedChange {
  assertVersion(input.currentVersion, 'currentVersion')
  assertVersion(input.fromVersion, 'fromVersion')
  if (input.fromVersion !== input.currentVersion) {
    throw new OfficeDocumentError(
      'stale-version',
      `Change targets version ${input.fromVersion} but the document is at ${input.currentVersion}.`,
    )
  }
  if (input.operation.length === 0) {
    throw new OfficeDocumentError('invalid-operation', 'A change requires at least one operation step.')
  }
  if (input.inverse.length === 0) {
    throw new OfficeDocumentError(
      'invalid-operation',
      'A change requires a non-empty inverse so document undo can restore the prior content.',
    )
  }
  const toVersion = input.fromVersion + 1
  assertVersion(toVersion, 'toVersion')
  return {
    record: {
      changeId: input.changeId ?? `${input.documentId}:${toVersion}`,
      documentId: input.documentId,
      actor: input.actor,
      fromVersion: input.fromVersion,
      toVersion,
      operation: input.operation,
      inverse: input.inverse,
      at: input.at,
    },
    version: toVersion,
  }
}

/**
 * Undo one accepted change at the document level.
 *
 * The caller applies the returned `operation` and records the returned
 * `version` as the live document version; this function mutates nothing.
 *
 * @param record - an accepted change.
 * @returns the record's inverse steps and the version to restore.
 */
export function undoChange(record: ChangeRecord): ChangeUndo {
  return { operation: record.inverse, version: record.fromVersion }
}
