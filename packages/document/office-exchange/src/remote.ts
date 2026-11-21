/**
 * Session-scoped `officeDocument` Remote surface over `ctx.officeExchange`.
 * The surface resolves the Session's workspace root through the same
 * `workspaceFileScope` lookup as `workspaceFiles`, confines every path to that
 * root, threads the filesystem content version, and persists saves through the
 * version-guarded binary write.
 * @module @forwork/office-exchange/remote
 */

import { createHash } from 'node:crypto'
import { mkdir, mkdtemp, rm } from 'node:fs/promises'
import { dirname, join, relative, sep } from 'node:path'
import type { Context } from '@forwork/cordis'
import z from '@forwork/schemastery'
import type { IDocumentData, IWorkbookData } from '@univerjs/core'
import type { WorkspaceFileScope } from '@forwork/api-workspace-files'
import { brandString } from '@forwork/brand'
import type { FileSystem, FsInfo, FsTarget, FsVersion } from '@forwork/fs'
import type { SandboxExecutionPolicy } from '@forwork/sandbox'
import type {} from '@forwork/sandbox-policy'
import { DocumentId } from '@forwork/session'
import type { DocumentFormat, DocumentSaveDisposition, Session } from '@forwork/session'
import {
  createDocumentIdentity,
  documentIdFor,
  markDocumentMissing as markIdentityMissing,
  openDocumentAtPath,
  rebindDocument as rebindIdentity,
  recordDocumentClose,
  recordDocumentMissing,
  recordDocumentOpen,
  recordDocumentRebind,
  recordDocumentSave,
} from '@forwork/session-document'
import type { DocumentIdentityBinding, DocumentProjectionEntry } from '@forwork/session-document'
import { Remote, RemoteError, TypertRemoteService } from '@forwork/typert-protocol'
import type { JsonValue } from '@forwork/util-values'
import { assertWithinRoots } from './authority.ts'
import type { DocumentCommandRequest, OfficeCommandBroker } from './commands.ts'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from './config.ts'
import { OfficeExchangeError } from './errors.ts'
import { open } from './package/ooxml-package.ts'
import type { SourcePackage } from './package/types.ts'
import { ZipError } from './package/zip-format.ts'
import type { OfficeExchange } from './service.ts'
import type { ExchangeReadResult, UnitSnapshot } from './snapshots.ts'
import type {
  DocumentSaveAttribution,
  EditableOfficeFormat,
  OfficeCommandOutcome,
  OfficeCommandResult,
  OfficeDocumentDraft,
  OfficeDocumentDraftWriteResult,
  OfficeDocumentReadResult,
  OfficeDocumentSaveResult,
  OfficeDocumentSnapshot,
  OfficeFormat,
  SaveDisposition,
  UnsupportedFeature,
} from './types.ts'

declare module '@forwork/cordis' {
  interface Context {
    /** Host owner of the `officeDocument` Remote namespace. */
    officeDocument: OfficeDocumentRemote
  }
}

/** Deployment bounds for the document Remote surface. */
export interface RemoteConfig {
  /**
   * Largest file one read may load, in bytes. The filesystem seam refuses a
   * larger target with `FS_TOO_LARGE` instead of truncating it.
   */
  readonly maxBytes: number
}

/** Validated bounds applied to every Remote read. */
export const RemoteConfig: z<Partial<RemoteConfig>, RemoteConfig> = z.object({
  maxBytes: z.natural().min(1).max(Number.MAX_SAFE_INTEGER - 1)
    .default(DEFAULT_OFFICE_EXCHANGE_CONFIG.maxFileBytes),
})

/** One open document the Remote tracks for the current Host process. */
interface OpenDocument {
  /** Identity binding across rename and missing observations. */
  readonly binding: DocumentIdentityBinding
  /** Editable format detected at open. */
  readonly format: DocumentFormat
  /** Latest content version this process observed for the document. */
  readonly contentVersion: number
}

/**
 * Host Remote service backing the generated `remote.officeDocument` namespace.
 * Reads return a JSON snapshot bound to the file's observed version; saves
 * replace only that version, so a concurrent external edit is reported as a
 * conflict rather than overwritten.
 *
 * The service is the producer of `document/*` facts for the editor: a first read
 * appends `document/open` once per Session, a successful save appends
 * `document/save`, and the lifecycle endpoints append their matching facts. The
 * open set is owned by this service per Session; no module-level state is used.
 */
export class OfficeDocumentRemote extends TypertRemoteService {
  static Config = RemoteConfig
  static inject = ['sessions']

  /** Open documents per Session, from this process's first read to its close. */
  private readonly openDocuments = new WeakMap<Session, Map<DocumentId, OpenDocument>>()

  /**
   * @param ctx - Host context carrying the Exchange seam, the filesystem, and Sessions.
   * @param config - deployment bounds for one Remote read.
   */
  constructor(ctx: Context, private readonly config: RemoteConfig) {
    super(ctx, 'officeDocument')
  }

  /**
   * Read an editable Office file into a JSON snapshot bound to its version.
   * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
   * @param path - absolute or workspace-relative path; it must resolve inside the workspace root.
   * The recorded document identity uses the same workspace-relative locator the editor tab uses,
   * so two spellings of one file never open two documents.
   * @param signal - caller cancellation.
   * @returns the detected format, snapshot, preserved construct classes,
   * unsupported features, and observed version.
   * @throws RemoteError when the path is out of scope, the format is not editable, or the read fails.
   */
  @Remote
  async read(
    workspaceFileScope: WorkspaceFileScope, path: string, signal: AbortSignal,
  ): Promise<OfficeDocumentReadResult> {
    try {
      const { target, info } = await this.locate(workspaceFileScope, path, signal)
      const fs = this.files()
      const sourceBytes = await fs.readBytes(target, signal, this.config.maxBytes)
      this.requireReadBound(sourceBytes.byteLength, path)
      const result = await this.exchange().read(fs.processPath(target), signal)
      const after = await fs.stat(target, signal)
      if (after === undefined || after.version !== info.version) {
        throw new OfficeExchangeError('conversion-failed', `"${path}" changed while it was being read.`)
      }
      const format = editableDocumentFormat(result.format)
      this.recordOpen(workspaceFileScope, this.identityPath(workspaceFileScope, target), format)
      return {
        format,
        snapshot: wireSnapshot(result),
        ...result.preserved === undefined ? {} : { preserved: result.preserved },
        unsupported: result.unsupported,
        version: info.version,
      }
    } catch (cause) {
      throw this.failure(cause, path, signal)
    }
  }

  /**
   * Serialize a snapshot and replace the file only at the caller's version.
   *
   * The file is read and stat-verified at `expectedVersion` before the writer
   * runs, so the bytes the writer receives are the bytes the caller read. A
   * preserved construct never refuses the save by itself: the writer returns a
   * typed refusal only when the edit intersects a construct whose meaning or
   * anchors depend on the changed content. A refusal writes nothing and records
   * no fact; the final `replaceIfVersion` write refuses a change that happened
   * during serialization.
   * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
   * @param path - absolute or workspace-relative path; it must resolve inside the workspace root.
   * @param snapshot - snapshot returned by {@link read}, optionally modified.
   * @param format - target editable format; it must match the snapshot kind.
   * @param expectedVersion - version returned by {@link read}; a changed file is refused.
   * @param attribution - who caused the save and the agent turn, when one caused it.
   * @param signal - caller cancellation.
   * @returns the version the guarded write published and the writer's disposition.
   * @throws RemoteError with `office-document/unsupported-feature` when the edit
   * intersects a coupled construct, naming that construct in the message.
   * @throws RemoteError with `office-document/conflict` when the file changed after the read.
   */
  @Remote
  async save(
    workspaceFileScope: WorkspaceFileScope,
    path: string,
    snapshot: OfficeDocumentSnapshot,
    format: OfficeFormat,
    expectedVersion: string,
    attribution: DocumentSaveAttribution,
    signal: AbortSignal,
  ): Promise<OfficeDocumentSaveResult> {
    try {
      const target = await this.confine(workspaceFileScope, path, signal)
      const native = nativeSnapshot(snapshot, format)
      const fs = this.files()
      this.requireVersion(path, await fs.stat(target, signal), expectedVersion)
      const sourceBytes = await fs.readBytes(target, signal, this.config.maxBytes)
      this.requireVersion(path, await fs.stat(target, signal), expectedVersion)
      const written = await this.exchange().write(native, format, sourcePackageOf(sourceBytes), signal)
      if (written.bytes.byteLength > this.config.maxBytes) {
        throw new OfficeExchangeError(
          'too-large',
          `"${path}" would be written as ${written.bytes.byteLength} bytes, exceeding the ${this.config.maxBytes} byte limit.`,
        )
      }
      const outcome = await fs.writeBytes(
        target,
        written.bytes,
        { kind: 'replaceIfVersion', version: brandString<FsVersion>(expectedVersion) },
        signal,
        this.sandboxPolicyFor(workspaceFileScope),
      )
      this.recordSave(
        workspaceFileScope,
        this.identityPath(workspaceFileScope, target),
        editableDocumentFormat(format),
        attribution,
        written.disposition,
      )
      return { version: outcome.version, disposition: written.disposition }
    } catch (cause) {
      throw this.failure(cause, path, signal)
    }
  }

  /**
   * Record that an open document left the editor.
   *
   * The fact is appended even when this process no longer tracks the document,
   * so a tab closed after a Host restart still leaves a durable close. History
   * is never deleted; the open set is derived by folding the log.
   * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
   * @param documentId - Session-scoped identity of the closing document.
   * @param reason - producer-supplied closure reason; the projection stores it unread.
   * @param signal - caller cancellation.
   * @throws RemoteError when the identity belongs to another Session or no live Session exists.
   */
  @Remote
  closeDocument(
    workspaceFileScope: WorkspaceFileScope, documentId: string, reason: string, signal: AbortSignal,
  ): Promise<void> {
    try {
      signal.throwIfAborted()
      const session = this.sessionOf(workspaceFileScope)
      const id = this.ownedDocumentId(workspaceFileScope, documentId)
      this.openDocumentsFor(session).delete(id)
      recordDocumentClose(session, { documentId: id, reason })
      return Promise.resolve()
    } catch (cause) {
      const failure: Error = this.failure(cause, documentId, signal)
      return Promise.reject(failure)
    }
  }

  /**
   * Record that an open document's file moved or was renamed.
   *
   * The identity is preserved: only the locator changes, and the new path must
   * resolve inside the Session workspace root.
   * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
   * @param documentId - Session-scoped identity retained across the move.
   * @param toPath - new path; it must resolve inside the workspace root.
   * @param signal - caller cancellation.
   * @throws RemoteError when the identity belongs to another Session, either path is out of scope, or no live Session exists.
   */
  @Remote
  async rebindDocument(
    workspaceFileScope: WorkspaceFileScope, documentId: string, toPath: string, signal: AbortSignal,
  ): Promise<void> {
    try {
      const session = this.sessionOf(workspaceFileScope)
      const id = this.ownedDocumentId(workspaceFileScope, documentId)
      const toTarget = await this.confine(workspaceFileScope, toPath, signal)
      const documents = this.openDocumentsFor(session)
      const entry = documents.get(id)
      const fromPath = entry?.binding.path ?? this.pathOf(documentId, workspaceFileScope)
      await this.confine(workspaceFileScope, fromPath, signal)
      signal.throwIfAborted()
      const exists = (await this.files().stat(toTarget, signal)) !== undefined
      // A repeat observation of the same path is not a rebind.
      if (fromPath === toPath && entry?.binding.missing !== true) return
      recordDocumentRebind(session, { documentId: id, fromPath, toPath })
      const binding = rebindIdentity(
        entry?.binding ?? createDocumentIdentity(id, fromPath),
        toPath,
        Date.now(),
        exists,
      )
      if (entry !== undefined) {
        documents.set(id, { ...entry, binding })
      }
    } catch (cause) {
      throw this.failure(cause, documentId, signal)
    }
  }

  /**
   * Record that an open document's file no longer resolves.
   *
   * The document stays open with its identity and history; a later save, edit,
   * or rebind marks it available again.
   * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
   * @param documentId - Session-scoped identity that stays open.
   * @param signal - caller cancellation.
   * @throws RemoteError when the identity belongs to another Session or no live Session exists.
   */
  @Remote
  markDocumentMissing(
    workspaceFileScope: WorkspaceFileScope, documentId: string, signal: AbortSignal,
  ): Promise<void> {
    try {
      signal.throwIfAborted()
      const session = this.sessionOf(workspaceFileScope)
      const id = this.ownedDocumentId(workspaceFileScope, documentId)
      const documents = this.openDocumentsFor(session)
      const entry = documents.get(id)
      if (entry?.binding.missing === true) return Promise.resolve()
      const path = entry?.binding.path ?? this.pathOf(documentId, workspaceFileScope)
      recordDocumentMissing(session, { documentId: id, path })
      if (entry !== undefined) {
        documents.set(id, { ...entry, binding: markIdentityMissing(entry.binding, Date.now()) })
      }
      return Promise.resolve()
    } catch (cause) {
      const failure: Error = this.failure(cause, documentId, signal)
      return Promise.reject(failure)
    }
  }

  /**
   * Report the features a whole-file regeneration would drop without returning content.
   * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
   * @param path - absolute or workspace-relative path; it must resolve inside the workspace root.
   * @param signal - caller cancellation.
   * @returns every detected unsupported feature; empty when regeneration loses no mapped content.
   * @throws RemoteError when the path is out of scope or its format cannot be inspected.
   */
  @Remote
  async detectUnsupported(
    workspaceFileScope: WorkspaceFileScope, path: string, signal: AbortSignal,
  ): Promise<readonly UnsupportedFeature[]> {
    try {
      const { target } = await this.locate(workspaceFileScope, path, signal)
      return await this.exchange().detectUnsupported(this.files().processPath(target), signal)
    } catch (cause) {
      throw this.failure(cause, path, signal)
    }
  }

  /**
   * Convert a legacy binary Office file to an editable one.
   * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
   * @param path - absolute or workspace-relative legacy path; it must resolve inside the workspace root.
   * @param signal - caller cancellation.
   * @returns the path of the converted editable file inside the workspace root.
   * @throws RemoteError when the path is out of scope or conversion fails.
   */
  @Remote
  async convertLegacy(workspaceFileScope: WorkspaceFileScope, path: string, signal: AbortSignal): Promise<string> {
    try {
      const { target } = await this.locate(workspaceFileScope, path, signal)
      const inputPath = this.files().processPath(target)
      const outputDir = await this.conversionDirectory(workspaceFileScope.workspaceRoot, inputPath)
      try {
        await assertWithinRoots(outputDir, [workspaceFileScope.workspaceRoot], { allowMissing: false })
        signal.throwIfAborted()
        return await this.exchange().convertLegacy(inputPath, signal, outputDir)
      } catch (cause) {
        await Promise.allSettled([rm(outputDir, { recursive: true, force: true })])
        throw cause
      }
    } catch (cause) {
      throw this.failure(cause, path, signal)
    }
  }

  /**
   * Persist one recovery draft for an open document.
   * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
   * @param documentId - caller-owned identity of the open document; it selects the draft file.
   * @param documentPath - document path the draft is captured for; it is recorded, never written.
   * @param contentVersion - file version the snapshot was captured from.
   * @param snapshot - wire snapshot of the editable content.
   * @param signal - caller cancellation.
   * @returns the time the draft was written.
   * @throws RemoteError when the draft store is out of scope or cannot be written.
   */
  @Remote
  async writeDraft(
    workspaceFileScope: WorkspaceFileScope,
    documentId: string,
    documentPath: string,
    contentVersion: string,
    snapshot: OfficeDocumentSnapshot,
    signal: AbortSignal,
  ): Promise<OfficeDocumentDraftWriteResult> {
    const draftPath = this.draftPath(workspaceFileScope.workspaceRoot, documentId)
    try {
      const target = await this.draftTarget(workspaceFileScope.workspaceRoot, draftPath, true, signal)
      const savedAt = new Date().toISOString()
      const draft: OfficeDocumentDraft = { documentId, path: documentPath, contentVersion, snapshot, savedAt }
      await this.files().writeText(
        target,
        JSON.stringify(draft),
        undefined,
        signal,
        this.sandboxPolicyFor(workspaceFileScope),
      )
      return { savedAt }
    } catch (cause) {
      throw this.failure(cause, draftPath, signal)
    }
  }

  /**
   * Read the recovery draft persisted for one open document.
   * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
   * @param documentId - caller-owned identity of the open document; it selects the draft file.
   * @param signal - caller cancellation.
   * @returns the stored draft, or `null` when the document has no draft.
   * @throws RemoteError when the draft store is out of scope or its content is malformed.
   */
  @Remote
  async readDraft(
    workspaceFileScope: WorkspaceFileScope, documentId: string, signal: AbortSignal,
  ): Promise<OfficeDocumentDraft | null> {
    const draftPath = this.draftPath(workspaceFileScope.workspaceRoot, documentId)
    try {
      const fs = this.files()
      const entry = await this.draftEntry(workspaceFileScope.workspaceRoot, draftPath, fs, signal)
      if (entry === undefined) return null
      if (entry.info.type !== 'file') {
        throw new OfficeExchangeError('conversion-failed', `The draft at "${draftPath}" is not a regular file.`)
      }
      const bytes = await fs.readBytes(entry.target, signal, this.config.maxBytes)
      return parseDraft(new TextDecoder().decode(bytes), draftPath)
    } catch (cause) {
      throw this.failure(cause, draftPath, signal)
    }
  }

  /**
   * Remove the recovery draft persisted for one open document.
   * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
   * @param documentId - caller-owned identity of the open document; it selects the draft file.
   * @param signal - caller cancellation.
   * @throws RemoteError when the draft store is out of scope or its entry cannot be removed.
   */
  @Remote
  async clearDraft(workspaceFileScope: WorkspaceFileScope, documentId: string, signal: AbortSignal): Promise<void> {
    const draftPath = this.draftPath(workspaceFileScope.workspaceRoot, documentId)
    try {
      const fs = this.files()
      const entry = await this.draftEntry(workspaceFileScope.workspaceRoot, draftPath, fs, signal)
      if (entry === undefined) return
      await fs.deleteFile(entry.target, undefined, signal, this.sandboxPolicyFor(workspaceFileScope))
    } catch (cause) {
      throw this.failure(cause, draftPath, signal)
    }
  }

  /**
   * Answer one pending document command with the live editor's result.
   *
   * A result for an unknown request id is ignored and logged by the broker;
   * the call itself succeeds so a late or duplicate answer never becomes a
   * transport error.
   * @param workspaceFileScope - Session identity that ran the command on the wire.
   * @param requestId - correlation id copied from the `document/command` event.
   * @param result - the live editor's status and, for read or applied commands, content.
   * @param signal - caller cancellation; an aborted call drops the answer.
   * @throws RemoteError with `gateway/bad-request` when a reported version is not a non-negative safe integer.
   */
  @Remote
  completeCommand(
    workspaceFileScope: WorkspaceFileScope,
    requestId: string,
    result: OfficeCommandResult,
    signal: AbortSignal,
  ): Promise<void> {
    if (signal.aborted) return Promise.resolve()
    return Promise.resolve().then(() => {
      assertCommandResult(result, requestId)
      this.commands().complete(workspaceFileScope, requestId, result)
    })
  }

  /**
   * Register one live editor target for a document.
   *
   * The renderer calls this when its command bridge attaches and releases the
   * matching registration when that bridge detaches. A command dispatched while
   * no registration exists is answered as `document-not-open` instead of
   * waiting for the timeout.
   * @param workspaceFileScope - Session identity that owns the document.
   * @param documentId - Session-scoped identity of the document.
   * @param registrationId - caller-minted identity of this attachment.
   * @param signal - caller cancellation; an aborted call registers nothing.
   * @throws RemoteError when the identity belongs to another Session or no live Session exists.
   */
  @Remote
  registerCommandTarget(
    workspaceFileScope: WorkspaceFileScope,
    documentId: string,
    registrationId: string,
    signal: AbortSignal,
  ): Promise<void> {
    return this.changeCommandTarget(workspaceFileScope, documentId, registrationId, signal, 'attach')
  }

  /**
   * Release one live editor target registration.
   * @param workspaceFileScope - Session identity that owns the document.
   * @param documentId - Session-scoped identity of the document.
   * @param registrationId - identity passed to {@link registerCommandTarget}.
   * @param signal - caller cancellation; an aborted call releases nothing.
   * @throws RemoteError when the identity belongs to another Session or no live Session exists.
   */
  @Remote
  unregisterCommandTarget(
    workspaceFileScope: WorkspaceFileScope,
    documentId: string,
    registrationId: string,
    signal: AbortSignal,
  ): Promise<void> {
    return this.changeCommandTarget(workspaceFileScope, documentId, registrationId, signal, 'detach')
  }

  /**
   * Apply one command-target registration change after validating its identity.
   * @param workspaceFileScope - Session identity that owns the document.
   * @param documentId - Session-scoped identity of the document.
   * @param registrationId - caller-minted identity of this attachment.
   * @param signal - caller cancellation; an aborted call changes nothing.
   * @param change - register (`attach`) or release (`detach`) the target.
   * @returns a resolved promise after the broker accepts the change.
   * @throws RemoteError when the identity belongs to another Session or no live Session exists.
   */
  private changeCommandTarget(
    workspaceFileScope: WorkspaceFileScope,
    documentId: string,
    registrationId: string,
    signal: AbortSignal,
    change: 'attach' | 'detach',
  ): Promise<void> {
    try {
      signal.throwIfAborted()
      this.requireRegistrationId(registrationId)
      const id = this.ownedDocumentId(workspaceFileScope, documentId)
      this.sessionOf(workspaceFileScope)
      const commands = this.commands()
      if (change === 'attach') commands.attachTarget(id, registrationId)
      else commands.detachTarget(id, registrationId)
      return Promise.resolve()
    } catch (cause) {
      const failure: Error = this.failure(cause, documentId, signal)
      return Promise.reject(failure)
    }
  }

  /**
   * Dispatch one correlated command from a Host document tool.
   *
   * The broker appends `document/command` and waits for the live editor; this
   * method is the Host-only entry point behind the tool's command executor and
   * is not a Remote endpoint.
   * @param request - session, document, operation, and expected version.
   * @param signal - caller cancellation; aborting withdraws the pending request.
   * @returns the editor's result, `document-not-open` without a registered target, or
   * `timeout` when a registered target did not answer in time.
   * @throws RemoteError with `gateway/internal` when the broker is not mounted.
   */
  async dispatchCommand(request: DocumentCommandRequest, signal: AbortSignal): Promise<OfficeCommandOutcome> {
    return await this.commands().request(request, signal)
  }

  /** Resolve the live Session that owns the workspace root on the wire. */
  private sessionOf(scope: WorkspaceFileScope): Session {
    const session = this.ctx.sessions.get(scope.sessionId)
    if (session === undefined) {
      throw new RemoteError(
        'gateway/bad-request',
        `Session "${scope.sessionId}" is not live, so its document facts cannot be recorded.`,
        {},
      )
    }
    return session
  }

  /**
   * Resolve the per-call sandbox policy for one Session's filesystem mutations.
   *
   * A composition without a policy service leaves the backend its own default.
   * A live Session contributes its mode and workspace root; a draft write for a
   * Session this process no longer holds live still fences to the workspace
   * root on the wire, under the deployment default mode.
   * @param scope - Session identity and workspace root on the wire.
   * @returns the policy for this call, or `undefined` when none is mounted.
   */
  private sandboxPolicyFor(scope: WorkspaceFileScope): SandboxExecutionPolicy | undefined {
    const service = this.ctx.get('sandboxPolicy')
    if (service === undefined) return undefined
    const session = this.ctx.sessions.get(scope.sessionId)
    if (session !== undefined) return service.resolve({ session })
    return { mode: service.defaultMode, workspaceRoot: scope.workspaceRoot }
  }

  /** The open-document table for one Session, created on first use. */
  private openDocumentsFor(session: Session): Map<DocumentId, OpenDocument> {
    let documents = this.openDocuments.get(session)
    if (documents === undefined) {
      documents = new Map()
      this.openDocuments.set(session, documents)
    }
    return documents
  }

  /**
   * Append `document/open` for the first read of a document in this process.
   * A repeat read of the same identity updates no fact: the open set already
   * names the document. A locator a rebind already tracks also opens nothing.
   *
   * After a restart the in-process table is empty, so the Session projection is
   * the only remaining record of a rebind. A document the projection still holds
   * at this locator keeps its identity and its content version; a missing entry
   * reopens with that identity, and a locator the projection does not hold opens
   * under its own path-derived identity.
   */
  private recordOpen(scope: WorkspaceFileScope, path: string, format: DocumentFormat): void {
    const session = this.sessionOf(scope)
    const documents = this.openDocumentsFor(session)
    const canonicalId = documentIdFor(scope.sessionId, path)
    if (this.trackedIdAt(documents, canonicalId, path) !== undefined) return
    const projected = this.projectedOpenAt(session, path)
    const documentId = projected?.documentId ?? canonicalId
    const contentVersion = projected?.contentVersion ?? 0
    if (projected?.status !== 'open') {
      recordDocumentOpen(session, { documentId, path, format, contentVersion })
    }
    documents.set(documentId, { binding: createDocumentIdentity(documentId, path), format, contentVersion })
  }

  /**
   * The document the Session projection still holds at one locator.
   *
   * The projection is optional in a Host composition that records no document
   * facts; without it, identity stays path-derived as before.
   * @param session - Session whose committed log owns the projection.
   * @param path - workspace-relative locator the read resolved.
   * @returns the open entry at that locator, or `undefined` when the projection is absent or holds none.
   */
  private projectedOpenAt(session: Session, path: string): DocumentProjectionEntry | undefined {
    const projections = this.ctx.get('sessionProjections')
    if (projections === undefined) return undefined
    const view = projections.snapshot(session, ['document']).values.document
    return view === undefined ? undefined : openDocumentAtPath(view, path)
  }

  /**
   * Append `document/save` after a guarded save published a new file.
   *
   * The caller's in-memory changes stay in the editor, so no operation steps
   * are available and no inverse is fabricated. The writer's disposition is
   * recorded unchanged except for the writer-only name cap, which the event
   * payload does not carry. A document this process has not read is opened
   * first so the save has an entry to advance. The identity is the
   * workspace-relative locator, so an agent save of a file the editor opened
   * advances that entry instead of opening a second one.
   */
  private recordSave(
    scope: WorkspaceFileScope,
    path: string,
    format: DocumentFormat,
    attribution: DocumentSaveAttribution,
    disposition: SaveDisposition,
  ): void {
    const session = this.sessionOf(scope)
    const documents = this.openDocumentsFor(session)
    const canonicalId = documentIdFor(scope.sessionId, path)
    const documentId = this.trackedIdAt(documents, canonicalId, path) ?? canonicalId
    const entry = documents.get(documentId)
    const broker = this.ctx.get('officeCommandBroker')
    const from = Math.max(entry?.contentVersion ?? 0, broker?.versionOf(documentId) ?? 0)
    const contentVersion = from + 1
    if (entry === undefined) {
      recordDocumentOpen(session, { documentId, path, format, contentVersion: 0 })
    }
    recordDocumentSave(session, {
      documentId,
      actor: attribution.actor,
      contentVersion,
      ...attribution.turn === undefined ? {} : { turn: attribution.turn },
      disposition: documentSaveDisposition(disposition),
    })
    documents.set(documentId, {
      binding: entry?.binding ?? createDocumentIdentity(documentId, path),
      format,
      contentVersion,
    })
    broker?.remember(documentId, contentVersion)
  }

  /**
   * The identity this Session already tracks for one locator.
   *
   * The canonical id names the locator, and a rebind keeps the original id
   * while its binding carries the new locator. Both resolve here, so a later
   * read or save of the moved file advances the same document instead of
   * opening a second one.
   * @param documents - open-document table for the Session.
   * @param canonicalId - identity derived from the current locator.
   * @param path - workspace-relative locator.
   * @returns the tracked identity, or `undefined` when this process has not read the file.
   */
  private trackedIdAt(
    documents: Map<DocumentId, OpenDocument>,
    canonicalId: DocumentId,
    path: string,
  ): DocumentId | undefined {
    if (documents.has(canonicalId)) return canonicalId
    for (const [documentId, entry] of documents) {
      if (entry.binding.path === path) return documentId
    }
    return undefined
  }

  /**
   * The workspace-relative locator every document consumer shares for one file.
   *
   * The editor tab addresses the file relative to the workspace, and the
   * Session facts name documents by that locator. Deriving it from the confined
   * target makes an absolute read and a relative read of the same file resolve
   * to one `documentId`.
   * @param scope - Session identity and workspace root on the wire.
   * @param target - resolved target already confined to the workspace root.
   * @returns the locator with `/` separators.
   */
  private identityPath(scope: WorkspaceFileScope, target: FsTarget): string {
    const absolute = this.files().processPath(target)
    return relative(scope.workspaceRoot, absolute).split(sep).join('/')
  }

  /** Require the caller to name its attachment with a non-empty id. */
  private requireRegistrationId(registrationId: string): void {
    if (registrationId.length === 0) {
      throw new RemoteError('gateway/bad-request', 'A command target registration requires a non-empty id.', {})
    }
  }

  /** Require a document identity that belongs to the Session on the wire. */
  private ownedDocumentId(scope: WorkspaceFileScope, documentId: string): DocumentId {
    const prefix = `${scope.sessionId}:`
    if (!documentId.startsWith(prefix) || documentId.length === prefix.length) {
      throw new RemoteError(
        'gateway/bad-request',
        `Document id "${documentId}" does not belong to Session "${scope.sessionId}".`,
        {},
      )
    }
    return DocumentId(documentId)
  }

  /** The path recorded inside a Session-scoped document identity. */
  private pathOf(documentId: string, scope: WorkspaceFileScope): string {
    return documentId.slice(`${scope.sessionId}:`.length)
  }

  /**
   * The draft file for one document identity: a digest keeps a caller-supplied
   * id from naming a path outside the draft directory.
   */
  private draftPath(workspaceRoot: string, documentId: string): string {
    const digest = createHash('sha256').update(documentId).digest('hex')
    return join(workspaceRoot, '.dsh', 'office-tmp', 'drafts', `${digest}.json`)
  }

  /** Resolve one draft path inside the workspace root, creating its directory for a write. */
  private async draftTarget(
    workspaceRoot: string, path: string, create: boolean, signal: AbortSignal,
  ): Promise<FsTarget> {
    const directory = dirname(path)
    await assertWithinRoots(directory, [workspaceRoot])
    if (create) {
      try {
        await mkdir(directory, { recursive: true })
      } catch (cause) {
        throw new OfficeExchangeError(
          'conversion-failed',
          `The draft store at "${directory}" could not be created (${String(cause)}).`,
        )
      }
      // Re-check after creation: a link placed between the first check and
      // `mkdir` resolves outside the root here instead of being written through.
      await assertWithinRoots(directory, [workspaceRoot], { allowMissing: false })
    }
    const fs = this.files()
    const target = await fs.resolve(path, { cwd: workspaceRoot, signal })
    await assertWithinRoots(fs.processPath(target), [workspaceRoot])
    return target
  }

  /**
   * Observe one existing draft file at its confined target.
   * @param workspaceRoot - Workspace root the draft store is confined to.
   * @param draftPath - Digest-derived draft path.
   * @param fs - Filesystem seam that already resolved the target.
   * @param signal - Caller cancellation.
   * @returns The confined target and latest observation, or `undefined` when absent.
   */
  private async draftEntry(
    workspaceRoot: string, draftPath: string, fs: FileSystem, signal: AbortSignal,
  ): Promise<{ target: FsTarget; info: FsInfo } | undefined> {
    const target = await this.draftTarget(workspaceRoot, draftPath, false, signal)
    const info = await fs.stat(target, signal)
    if (info === undefined) return undefined
    return { target, info }
  }

  /**
   * Create a fresh conversion directory inside the workspace root so the
   * converted file passes the same root check as every other Remote path. The
   * hash names the source and `mkdtemp` keeps concurrent conversions apart.
   */
  private async conversionDirectory(workspaceRoot: string, inputPath: string): Promise<string> {
    const parent = join(workspaceRoot, '.dsh', 'office-tmp')
    await mkdir(parent, { recursive: true })
    const digest = createHash('sha256').update(inputPath).digest('hex').slice(0, 16)
    return await mkdtemp(join(parent, `convert-${digest}-`))
  }

  /** Resolve one path inside the Session's workspace root and require a regular file. */
  private async locate(
    workspaceFileScope: WorkspaceFileScope, path: string, signal: AbortSignal,
  ): Promise<{ target: FsTarget; info: FsInfo }> {
    const target = await this.confine(workspaceFileScope, path, signal)
    const info = await this.files().stat(target, signal)
    if (info === undefined || info.type !== 'file') {
      throw new RemoteError('office-document/not-found', `No regular file at "${path}".`, { path })
    }
    return { target, info }
  }

  /** Resolve one path inside the Session's workspace root; absence is left to the caller. */
  private async confine(workspaceFileScope: WorkspaceFileScope, path: string, signal: AbortSignal): Promise<FsTarget> {
    const fs = this.files()
    const target = await fs.resolve(path, { cwd: workspaceFileScope.workspaceRoot, signal })
    await assertWithinRoots(fs.processPath(target), [workspaceFileScope.workspaceRoot])
    return target
  }

  /**
   * Require the source to still be the regular file the caller read.
   * @param path - caller-supplied path used in the conflict report.
   * @param info - latest observation of the target, or `undefined` when it is absent.
   * @param expectedVersion - version the caller read the source at.
   * @throws RemoteError with `office-document/conflict` when the file is absent, is not a
   * regular file, or changed version.
   */
  private requireVersion(path: string, info: FsInfo | undefined, expectedVersion: string): void {
    if (info !== undefined && info.type === 'file' && info.version === expectedVersion) return
    throw new RemoteError(
      'office-document/conflict',
      `"${path}" changed after it was read; reopen it before saving.`,
      { path },
    )
  }

  /**
   * Require the bytes this read retained to stay inside the configured bound.
   *
   * The filesystem seam already refuses a larger target, so this guard also
   * covers a seam that returns oversized bytes instead of refusing them.
   * @param byteLength - Bytes the read retained.
   * @param path - caller-supplied path used in the bound report.
   * @throws {OfficeExchangeError} `too-large` when the retained bytes exceed `maxBytes`.
   */
  private requireReadBound(byteLength: number, path: string): void {
    if (byteLength <= this.config.maxBytes) return
    throw new OfficeExchangeError(
      'too-large',
      `"${path}" is ${byteLength} bytes, exceeding the ${this.config.maxBytes} byte limit.`,
    )
  }

  /** The Exchange seam, required for every operation. */
  private exchange(): OfficeExchange {
    const exchange = this.ctx.get('officeExchange')
    if (exchange === undefined) {
      throw new RemoteError(
        'gateway/internal',
        'officeExchange is absent: mount @forwork/office-exchange in this composition',
        {},
      )
    }
    return exchange
  }

  /** The correlated command broker, required for command dispatch and completion. */
  private commands(): OfficeCommandBroker {
    const broker = this.ctx.get('officeCommandBroker')
    if (broker === undefined) {
      throw new RemoteError(
        'gateway/internal',
        'officeCommandBroker is absent: mount @forwork/office-exchange in this composition',
        {},
      )
    }
    return broker
  }

  /** The filesystem seam, required for resolution, reads, versions, and saves. */
  private files(): FileSystem {
    const fs = this.ctx.get('fs')
    if (fs === undefined) {
      throw new RemoteError(
        'gateway/internal',
        'fs is absent: mount a filesystem provider in this composition',
        {},
      )
    }
    return fs
  }

  /** Map a seam, filesystem, or cancellation failure onto the Remote error vocabulary. */
  private failure(cause: unknown, path: string, signal: AbortSignal): Error {
    if (cause instanceof RemoteError) return cause
    if (signal.aborted) {
      return new RemoteError('gateway/cancelled', `"${path}" was not read or written.`, {}, { cause })
    }
    if (cause instanceof OfficeExchangeError) {
      switch (cause.code) {
        case 'unsupported-format':
          return new RemoteError('office-document/unsupported-format', cause.message, { path }, { cause })
        case 'unsupported-feature':
          return new RemoteError('office-document/unsupported-feature', cause.message, { path }, { cause })
        case 'out-of-scope':
          return new RemoteError('office-document/out-of-scope', cause.message, { path }, { cause })
        case 'too-large':
          return new RemoteError('office-document/bound-exceeded', cause.message, { path }, { cause })
        case 'conversion-failed':
        case 'not-implemented':
          return new RemoteError('office-document/conversion-failed', cause.message, { path }, { cause })
      }
    }
    if (cause instanceof ZipError) {
      const code = cause.failure.type === 'entry-limit' || cause.failure.type === 'too-large'
        ? 'office-document/bound-exceeded'
        : 'office-document/conversion-failed'
      return new RemoteError(code, cause.message, { path }, { cause })
    }
    if (isFsCode(cause, 'FS_TOO_LARGE')) {
      return new RemoteError(
        'office-document/bound-exceeded',
        `"${path}" exceeds the ${this.config.maxBytes} byte limit.`,
        { path },
        { cause },
      )
    }
    if (isFsCode(cause, 'FS_STALE_VERSION')) {
      return new RemoteError(
        'office-document/conflict',
        `"${path}" changed after it was read; reopen it before saving.`,
        { path },
        { cause },
      )
    }
    if (isFsCode(cause, 'FS_NOT_FOUND') || isFsCode(cause, 'FS_NOT_REGULAR_FILE')
      || isFsCode(cause, 'FS_PERMISSION_DENIED') || isFsCode(cause, 'FS_SANDBOX_DENIED')
      || isFsCode(cause, 'FS_IO_ERROR')) {
      return new RemoteError('office-document/conversion-failed', `"${path}" could not be read or written.`, { path }, { cause })
    }
    return cause instanceof Error
      ? cause
      : new RemoteError('office-document/conversion-failed', `"${path}" could not be read or written.`, { path }, { cause })
  }
}

export default OfficeDocumentRemote

/** Require the version fields a command result reports to be non-negative safe integers. */
function assertCommandResult(result: OfficeCommandResult, requestId: string): void {
  switch (result.status) {
    case 'read':
    case 'applied':
      assertCommandVersion(result.contentVersion, requestId, 'contentVersion')
      return
    case 'conflict':
      assertCommandVersion(result.currentVersion, requestId, 'currentVersion')
      return
    case 'pending-edits':
    case 'unsupported-feature':
    case 'document-not-open':
      return
  }
}

/** Refuse a wire-reported version a document cannot hold. */
function assertCommandVersion(version: number, requestId: string, field: string): void {
  if (!Number.isSafeInteger(version) || version < 0) {
    throw new RemoteError(
      'gateway/bad-request',
      `Document command "${requestId}" reported ${field} ${String(version)}.`,
      { issues: [{ field, value: version }] },
    )
  }
}

/** The backend's stable failure code, recognized structurally across package copies. */
function isFsCode(error: unknown, code: string): boolean {
  return typeof error === 'object' && error !== null && 'code' in error && error.code === code
}

/**
 * Central-directory entry bound applied when a save indexes its source. It
 * matches the package scan bound detection uses, so detection and the writer
 * accept the same archives.
 */
const MAX_SOURCE_ENTRIES = 10_000

/**
 * Index the source bytes a guarded save will splice.
 * @param bytes - Source file bytes read at the caller's expected version.
 * @returns The indexed source package with its content digest.
 * @throws {ZipError} When the container is unsupported or malformed.
 */
function sourcePackageOf(bytes: Uint8Array): SourcePackage {
  return open(bytes, { maxEntries: MAX_SOURCE_ENTRIES }).source
}

/**
 * Project the writer's disposition onto the durable save-fact payload.
 *
 * The event payload carries the exact counts and each capped name list; the
 * writer-only name-limit field is not part of the session vocabulary.
 * @param disposition - disposition the writer computed from its splice plan.
 * @returns the same counts and names without writer-only fields.
 */
function documentSaveDisposition(disposition: SaveDisposition): DocumentSaveDisposition {
  return {
    schemaVersion: disposition.schemaVersion,
    verbatimCount: disposition.verbatimCount,
    spliced: { count: disposition.spliced.count, names: [...disposition.spliced.names] },
    added: { count: disposition.added.count, names: [...disposition.added.names] },
    removed: { count: disposition.removed.count, names: [...disposition.removed.names] },
  }
}

/** Project a native snapshot onto the JSON-safe wire value. */
function wireSnapshot(result: ExchangeReadResult): OfficeDocumentSnapshot {
  return { kind: result.snapshot.kind, data: jsonValueOf(result.snapshot.data) }
}

/**
 * Validate and detach one native snapshot field as lossless JSON. Univer adds
 * values JSON cannot carry (class instances, functions, `undefined` members) to
 * the native interfaces through their custom extension points, so this walks
 * the value once instead of asserting the type.
 */
function jsonValueOf(value: unknown): JsonValue {
  if (value === null || typeof value === 'boolean' || typeof value === 'string') return value
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new OfficeExchangeError('conversion-failed', 'The snapshot contains a non-finite number.')
    return value
  }
  if (Array.isArray(value)) return value.map(jsonValueOf)
  if (typeof value === 'object') {
    const prototype: unknown = Object.getPrototypeOf(value)
    if (prototype !== Object.prototype && prototype !== null) {
      throw new OfficeExchangeError('conversion-failed', 'The snapshot contains a non-JSON object.')
    }
    const record: Record<string, JsonValue> = {}
    const entries: [string, unknown][] = Object.entries(value)
    for (const [key, member] of entries) {
      if (member === undefined) continue
      record[key] = jsonValueOf(member)
    }
    return record
  }
  throw new OfficeExchangeError('conversion-failed', `The snapshot contains a non-JSON ${typeof value} value.`)
}

/**
 * Narrow a detected format to one the document facts admit.
 * @param format - format reported by the Exchange read or save.
 * @returns the editable document or workbook format.
 * @throws {OfficeExchangeError} `unsupported-format` for a legacy or presentation format.
 */
function editableDocumentFormat(format: OfficeFormat): EditableOfficeFormat {
  switch (format) {
    case 'docx':
    case 'docm':
    case 'xlsx':
    case 'xlsm':
      return format
    case 'pptx':
    case 'doc':
    case 'xls':
    case 'ppt':
      throw new OfficeExchangeError(
        'unsupported-format',
        `Cannot record a .${format} document; only .docx, .docm, .xlsx, and .xlsm are editable.`,
      )
  }
}

/** Restore the native unit snapshot a save must serialize. */
function nativeSnapshot(snapshot: OfficeDocumentSnapshot, format: OfficeFormat): UnitSnapshot {
  switch (format) {
    case 'xlsx':
    case 'xlsm':
      if (snapshot.kind !== 'workbook') {
        throw new OfficeExchangeError('conversion-failed', `A document snapshot cannot be saved as .${format}.`)
      }
      return { kind: 'workbook', data: workbookData(snapshot.data) }
    case 'docx':
    case 'docm':
      if (snapshot.kind !== 'document') {
        throw new OfficeExchangeError('conversion-failed', `A workbook snapshot cannot be saved as .${format}.`)
      }
      return { kind: 'document', data: documentData(snapshot.data) }
    case 'pptx':
    case 'doc':
    case 'xls':
    case 'ppt':
      throw new OfficeExchangeError(
        'unsupported-format',
        `Cannot save .${format} files; only .docx, .docm, .xlsx, and .xlsm are editable.`,
      )
  }
}

/** Require the identity fields every `IWorkbookData` mapping dereferences. */
function workbookData(value: JsonValue): IWorkbookData {
  if (!isWorkbookData(value)) {
    throw new OfficeExchangeError('conversion-failed', 'The workbook snapshot is not a workbook data object.')
  }
  return value
}

/** Require the identity fields every `IDocumentData` mapping dereferences. */
function documentData(value: JsonValue): IDocumentData {
  if (!isDocumentData(value)) {
    throw new OfficeExchangeError('conversion-failed', 'The document snapshot is not a document data object.')
  }
  return value
}

/**
 * Whether a JSON value carries the required workbook identity and containers.
 * Content beyond this envelope is validated by the provider's typed OOXML
 * mapping, which reports `conversion-failed` instead of writing a partial file.
 */
function isWorkbookData(value: unknown): value is IWorkbookData {
  if (!isJsonObject(value)) return false
  return typeof value['id'] === 'string'
    && typeof value['name'] === 'string'
    && typeof value['appVersion'] === 'string'
    && typeof value['locale'] === 'string'
    && isJsonObject(value['styles'])
    && isJsonObject(value['sheets'])
    && Array.isArray(value['sheetOrder'])
    && value['sheetOrder'].every(entry => typeof entry === 'string')
}

/**
 * Whether a JSON value carries the required document identity and containers.
 * Content beyond this envelope is validated by the provider's typed OOXML
 * mapping, which reports `conversion-failed` instead of writing a partial file.
 */
function isDocumentData(value: unknown): value is IDocumentData {
  if (!isJsonObject(value)) return false
  return typeof value['id'] === 'string'
    && isJsonObject(value['documentStyle'])
    && (value['body'] === undefined || isJsonObject(value['body']))
}

/** Whether a JSON value is a plain object the mapping helpers can index. */
function isJsonObject(value: unknown): value is Record<string, JsonValue> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

/**
 * Parse persisted draft text. Malformed JSON, a non-object root, or a missing
 * required field is a `conversion-failed` draft rather than a partial draft.
 */
function parseDraft(text: string, path: string): OfficeDocumentDraft {
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch (error) {
    throw new OfficeExchangeError('conversion-failed', `The draft at "${path}" is not valid JSON (${String(error)}).`)
  }
  if (!isJsonObject(parsed)) {
    throw new OfficeExchangeError('conversion-failed', `The draft at "${path}" is not a JSON object.`)
  }
  const draft = draftRecord(parsed)
  if (draft === undefined) {
    throw new OfficeExchangeError('conversion-failed', `The draft at "${path}" is missing a required field.`)
  }
  return draft
}

/** Project one parsed draft object onto the wire draft, or refuse an incomplete record. */
function draftRecord(value: Record<string, JsonValue>): OfficeDocumentDraft | undefined {
  const { documentId, path, contentVersion, savedAt, snapshot } = value
  if (
    typeof documentId !== 'string'
    || typeof path !== 'string'
    || typeof contentVersion !== 'string'
    || typeof savedAt !== 'string'
    || !isDraftSnapshot(snapshot)
  ) return undefined
  return { documentId, path, contentVersion, savedAt, snapshot }
}

/** Whether a parsed value carries an edit snapshot with its kind and payload. */
function isDraftSnapshot(value: JsonValue | undefined): value is OfficeDocumentSnapshot {
  if (!isJsonObject(value)) return false
  if (value['kind'] !== 'workbook' && value['kind'] !== 'document') return false
  return isJsonObject(value['data'])
}
