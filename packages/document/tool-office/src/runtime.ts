/**
 * Narrow seams the Office document tools consume. Production wiring backs the
 * document port with `ctx.officeDocument` and the command executor with its
 * Host-only `dispatchCommand`; tests supply fakes without a renderer.
 * @module @forwork/tool-office/runtime
 */

import type { WorkspaceFileScope } from '@forwork/api-workspace-files'
import type {
  DocumentCommandRequest,
  DocumentSaveAttribution,
  OfficeCommandOutcome,
  OfficeDocumentReadResult,
  OfficeDocumentSaveResult,
  OfficeDocumentSnapshot,
  OfficeFormat,
} from '@forwork/office-exchange'
import type { DocumentId, Session } from '@forwork/session'

/** File reads, saves, and legacy conversion over the `officeDocument` Remote. */
export interface OfficeDocumentPort {
  /**
   * Read an editable file and its current filesystem version.
   * @param scope - Session identity and workspace root.
   * @param path - confined path.
   * @param signal - caller cancellation.
   * @returns detected format, snapshot, unsupported features, and file version.
   */
  read(scope: WorkspaceFileScope, path: string, signal: AbortSignal): Promise<OfficeDocumentReadResult>

  /**
   * Replace the file only at the observed version.
   * @param scope - Session identity and workspace root.
   * @param path - confined path.
   * @param snapshot - content to serialize.
   * @param format - target editable format.
   * @param expectedVersion - version returned by {@link read}.
   * @param attribution - the agent that caused the save and its open turn.
   * @param signal - caller cancellation.
   * @returns the version the guarded write published.
   */
  save(
    scope: WorkspaceFileScope,
    path: string,
    snapshot: OfficeDocumentSnapshot,
    format: OfficeFormat,
    expectedVersion: string,
    attribution: DocumentSaveAttribution,
    signal: AbortSignal,
  ): Promise<OfficeDocumentSaveResult>

  /**
   * Convert a legacy binary file inside the workspace.
   * @param scope - Session identity and workspace root.
   * @param path - confined legacy path.
   * @param signal - caller cancellation.
   * @returns the converted editable path.
   */
  convertLegacy(scope: WorkspaceFileScope, path: string, signal: AbortSignal): Promise<string>
}

/** Correlated command transport to the live editor. */
export interface DocumentCommandExecutor {
  /**
   * Append one `document/command` request and await the editor's answer.
   * @param request - Session, document, operation, and expected version.
   * @param signal - caller cancellation.
   * @returns the editor's result or `timeout`.
   */
  execute(request: DocumentCommandRequest, signal: AbortSignal): Promise<OfficeCommandOutcome>

  /**
   * Read the last content version the Host observed for one document.
   * @param documentId - stable identity of the document.
   * @returns the observed version, or `0` before any command observed one.
   */
  currentVersion(documentId: DocumentId): number

  /**
   * Record the content version the document now holds.
   * @param documentId - stable identity of the document.
   * @param version - non-negative version observed from the editor.
   */
  remember(documentId: DocumentId, version: number): void
}

/** Host services the document tools consume. */
export interface DocumentToolRuntime {
  /** File read, save, and conversion surface. */
  readonly documents: OfficeDocumentPort
  /** Live editor command transport and version observations. */
  readonly commands: DocumentCommandExecutor
  /**
   * The open turn of a Session, when one is running.
   * @param session - Session whose durable turn boundary is read.
   * @returns the turn number, or `undefined` between turns.
   */
  turnOf(session: Session): number | undefined
}

/** Deployment caps applied to one tool call. */
export interface DocumentToolConfig {
  /** Largest cell preview a read summary may list. */
  readonly summaryCells: number
  /** Largest single edit operation. */
  readonly maxEditSteps: number
}

/** One resolved tool call: runtime, caller Session, signal, and caps. */
export interface DocumentToolCall {
  /** Host services consumed by the operation. */
  readonly runtime: DocumentToolRuntime
  /** Deployment caps. */
  readonly config: DocumentToolConfig
  /** Calling Session that records the command and the resulting edit. */
  readonly session: Session
  /** Session identity and workspace root for the Remote calls. */
  readonly scope: WorkspaceFileScope
  /** Caller cancellation. */
  readonly signal: AbortSignal
}
