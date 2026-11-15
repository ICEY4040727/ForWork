/**
 * The open document's Remote-backed state.
 *
 * Reads run through the injected `officeDocument` Remote — never the Host
 * service — and a save sends the live snapshot with the version the read
 * observed. A refused version is reported as a conflict; a file that reports
 * an unsupported feature cannot be saved at all. An external version change
 * reloads a clean unit and prompts instead of reloading a dirty one.
 *
 * Edits are captured as idle-debounced recovery drafts through the same
 * Remote. A draft found after a successful read is offered only as a recovery
 * choice: it is never applied automatically, and a draft captured from an
 * older file version is labelled stale so the user chooses it explicitly. A
 * successful save and a discard both clear the stored draft.
 *
 * A legacy `.doc`, `.xls`, or `.ppt` address is never read: it is converted
 * through the same Remote, and the converted copy opens in a new tab in this
 * one's place. The original is never modified, and this body reports
 * conversion state until the new tab takes over.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import type { RemoteFailure, RemoteResult } from '@deepseek-ai/dsh-typert-protocol'
import type {
  OfficeDocumentDraft,
  OfficeDocumentDraftWriteResult,
  OfficeDocumentReadResult,
  OfficeDocumentSaveResult,
  OfficeDocumentSnapshot,
  OfficeFormat,
  UnsupportedFeature,
} from '@deepseek-ai/dsh-office-exchange/types'
import { externalChangeVerdict } from './changes.ts'
import { createDraftScheduler } from './drafts.ts'
import { convertedAddress, legacyFormatOf, officeDocumentIdentity } from './target.ts'

/** The `officeDocument` Remote methods this hook calls. */
export interface OfficeDocumentRemote {
  /**
   * Read one editable file into a JSON snapshot.
   * @param sessionId - the Session whose workspace resolves `path`.
   * @param path - workspace path of the file.
   * @param signal - cancels the call.
   * @returns the snapshot, unsupported features, and observed version, or the declared failure.
   */
  read(sessionId: SessionId, path: string, signal?: AbortSignal): Promise<RemoteResult<OfficeDocumentReadResult>>
  /**
   * Convert one legacy binary Office file into an editable copy inside the workspace.
   * @param sessionId - the Session whose workspace resolves `path`.
   * @param path - workspace path of the legacy source.
   * @param signal - cancels the call and stops the conversion engine.
   * @returns the converted file's path, or the declared failure.
   */
  convertLegacy(sessionId: SessionId, path: string, signal?: AbortSignal): Promise<RemoteResult<string>>
  /**
   * Replace one file at the caller's version.
   * @param sessionId - the Session whose workspace resolves `path`.
   * @param path - workspace path of the file.
   * @param snapshot - the live unit snapshot.
   * @param format - the target editable format.
   * @param expectedVersion - the version the read observed.
   * @param signal - cancels the call.
   * @returns the version the write published, or the declared failure.
   */
  save(
    sessionId: SessionId,
    path: string,
    snapshot: OfficeDocumentSnapshot,
    format: OfficeFormat,
    expectedVersion: string,
    signal?: AbortSignal,
  ): Promise<RemoteResult<OfficeDocumentSaveResult>>
  /**
   * Persist the recovery draft for one open document.
   * @param sessionId - the Session whose workspace owns the draft store.
   * @param documentId - caller-owned identity of the open document.
   * @param documentPath - document path the draft was captured for.
   * @param contentVersion - file version the snapshot was captured from.
   * @param snapshot - wire snapshot of the editable content.
   * @param signal - cancels the call.
   * @returns the time the draft was written, or the declared failure.
   */
  writeDraft(
    sessionId: SessionId,
    documentId: string,
    documentPath: string,
    contentVersion: string,
    snapshot: OfficeDocumentSnapshot,
    signal?: AbortSignal,
  ): Promise<RemoteResult<OfficeDocumentDraftWriteResult>>
  /**
   * Read the recovery draft persisted for one open document.
   * @param sessionId - the Session whose workspace owns the draft store.
   * @param documentId - caller-owned identity of the open document.
   * @param signal - cancels the call.
   * @returns the stored draft, `null` when none exists, or the declared failure.
   */
  readDraft(
    sessionId: SessionId,
    documentId: string,
    signal?: AbortSignal,
  ): Promise<RemoteResult<OfficeDocumentDraft | null>>
  /**
   * Remove the recovery draft persisted for one open document.
   * @param sessionId - the Session whose workspace owns the draft store.
   * @param documentId - caller-owned identity of the open document.
   * @param signal - cancels the call.
   * @returns the declared failure, if any.
   */
  clearDraft(
    sessionId: SessionId,
    documentId: string,
    signal?: AbortSignal,
  ): Promise<RemoteResult<void>>
}

/** A stored draft offered as a recovery choice after a successful read. */
export interface OfficeDraftRecovery {
  /** ISO-8601 time the offered draft was written. */
  readonly savedAt: string
  /** The draft was captured from an older file version. */
  readonly stale: boolean
  /** The captured editable content. */
  readonly snapshot: OfficeDocumentSnapshot
}

/** Stable failure classes the panel localizes. */
export type OfficeDocumentFailureCode =
  | 'conflict'
  | 'unsupported-format'
  | 'unsupported-feature'
  | 'out-of-scope'
  | 'bound-exceeded'
  | 'conversion-failed'
  | 'cancelled'
  | 'unknown'

/** One failed read or save. */
export interface OfficeDocumentFailure {
  /** Stable class deciding the visible copy and available recourse. */
  readonly code: OfficeDocumentFailureCode
  /** Host diagnostic, kept for logs; product copy comes from the locale. */
  readonly message: string
}

/** What the panel is showing. */
export type OfficeDocumentStatus = 'loading' | 'converting' | 'ready' | 'unsupported' | 'conflict' | 'error'

/** One open document's observable state. */
export interface OfficeDocumentState {
  /** Load, edit-ready, refusal, save-conflict, or failed state. */
  readonly status: OfficeDocumentStatus
  /** The read snapshot; absent while loading or after a failed read. */
  readonly snapshot: OfficeDocumentSnapshot | undefined
  /** Version the snapshot was read from; absent until a read succeeds. */
  readonly version: string | undefined
  /** Features a save would drop. */
  readonly unsupported: readonly UnsupportedFeature[]
  /** The last failure, if any. */
  readonly error: OfficeDocumentFailure | undefined
}

/** Result of one save attempt. */
export type OfficeSaveOutcome = 'saved' | 'refused' | 'conflict' | 'error'

/** Everything the hook needs from its owning panel. */
export interface UseOfficeDocumentInput {
  /** Injected Remote face. */
  readonly remote: OfficeDocumentRemote
  /** Session whose workspace owns the path. */
  readonly sessionId: SessionId
  /** Workspace path of the open file. */
  readonly path: string
  /** Tab lifetime; aborts reads and writes when the record ends. */
  readonly signal: AbortSignal
  /** File version the `file` resource currently reports. */
  readonly observedVersion: string | undefined
  /** Workspace root derived from the opened address and its absolute path, when known. */
  readonly workspaceRoot: string | undefined
  /**
   * Open a converted legacy copy in this tab's place and record where it came from.
   * @param address - the converted copy's file resource address.
   * @param source - the legacy workspace path the user opened.
   */
  readonly openConverted: (address: string, source: string) => void
  /** Whether the live unit has unsaved edits. */
  readonly dirty: boolean
}

/** The hook's state plus the actions the panel calls. */
export interface UseOfficeDocumentResult {
  /** Current visible state. */
  readonly status: OfficeDocumentStatus
  /** Read snapshot. */
  readonly snapshot: OfficeDocumentSnapshot | undefined
  /** Version of the displayed snapshot. */
  readonly version: string | undefined
  /** Features a save would drop. */
  readonly unsupported: readonly UnsupportedFeature[]
  /** Last failure. */
  readonly error: OfficeDocumentFailure | undefined
  /** Whether the panel may offer save. */
  readonly canSave: boolean
  /** A save is in flight. */
  readonly saving: boolean
  /** The file changed externally while the unit was dirty. */
  readonly externalChange: boolean
  /** A stored draft the user may recover instead of starting from the file. */
  readonly recovery: OfficeDraftRecovery | undefined
  /** The last draft write or draft-store failure. */
  readonly draftError: OfficeDocumentFailure | undefined
  /** The live unit currently holds an accepted recovery draft. */
  readonly recovered: boolean
  /**
   * Capture the live unit as a recovery draft after the idle period.
   * @param snapshot - the live unit snapshot to capture.
   */
  scheduleDraft(snapshot: OfficeDocumentSnapshot): void
  /** Replace the live unit with the offered recovery draft. */
  acceptRecovery(): void
  /** Dismiss the offered draft and remove it from the draft store. */
  discardRecovery(): void
  /** Hide the last draft failure message. */
  dismissDraftError(): void
  /** Drop any pending draft and remove the stored draft without closing. */
  discardDraft(): void
  /**
   * Save the live unit at the version the read observed.
   * @param edited - the live snapshot; defaults to the read snapshot.
   * @returns whether the write landed, was refused, conflicted, or failed.
   */
  save(edited?: OfficeDocumentSnapshot): Promise<OfficeSaveOutcome>
  /** Re-read the file, discarding the live unit. */
  reload(): void
  /** Dismiss an external-change prompt and keep editing. */
  dismissExternalChange(): void
}

/** State before the first read settles. */
const INITIAL: OfficeDocumentState = {
  status: 'loading',
  snapshot: undefined,
  version: undefined,
  unsupported: [],
  error: undefined,
}

/** Format a save writes for one snapshot kind. */
function formatOf(snapshot: OfficeDocumentSnapshot): OfficeFormat {
  return snapshot.kind === 'workbook' ? 'xlsx' : 'docx'
}

/** Render a non-Error rejection as a diagnostic message. */
function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

/**
 * Map a Remote failure onto the stable classes the panel localizes.
 * @param failure - the settled Remote failure.
 * @returns its product-facing class and diagnostic message.
 */
export function mapOfficeFailure(failure: RemoteFailure): OfficeDocumentFailure {
  switch (failure.code) {
    case 'office-document/conflict':
      return { code: 'conflict', message: failure.message }
    case 'office-document/unsupported-format':
      return { code: 'unsupported-format', message: failure.message }
    case 'office-document/unsupported-feature':
      return { code: 'unsupported-feature', message: failure.message }
    case 'office-document/out-of-scope':
      return { code: 'out-of-scope', message: failure.message }
    case 'office-document/bound-exceeded':
      return { code: 'bound-exceeded', message: failure.message }
    case 'office-document/conversion-failed':
      return { code: 'conversion-failed', message: failure.message }
    case 'gateway/cancelled':
      return { code: 'cancelled', message: failure.message }
    default:
      return { code: 'unknown', message: failure.message }
  }
}

/**
 * Read, track, and save one Office document through the injected Remote.
 * @param input - the Remote face, open target, tab lifetime, and dirty state.
 * @returns the visible state and the panel's actions.
 */
export function useOfficeDocument(input: UseOfficeDocumentInput): UseOfficeDocumentResult {
  const { remote, sessionId, path, signal, observedVersion, workspaceRoot, openConverted, dirty } = input
  const legacy = legacyFormatOf(path)
  const [state, setState] = useState<OfficeDocumentState>(INITIAL)
  const [saving, setSaving] = useState(false)
  const [externalChange, setExternalChange] = useState(false)
  const [generation, setGeneration] = useState(0)
  const [recovery, setRecovery] = useState<OfficeDraftRecovery | undefined>(undefined)
  const [draftError, setDraftError] = useState<OfficeDocumentFailure | undefined>(undefined)
  const [recovered, setRecovered] = useState(false)
  const savingRef = useRef(false)
  const observedRef = useRef(observedVersion)
  const observedAtRead = useRef<string | undefined>(undefined)
  const rootRef = useRef(workspaceRoot)
  const draftCheckedFor = useRef<string | undefined>(undefined)
  const statusRef = useRef(state.status)
  const versionRef = useRef(state.version)
  observedRef.current = observedVersion
  rootRef.current = workspaceRoot
  statusRef.current = state.status
  versionRef.current = state.version
  const documentId = useMemo(() => officeDocumentIdentity(sessionId, path), [sessionId, path])

  const reportDraftError = useCallback((failure: RemoteFailure): void => {
    setDraftError(mapOfficeFailure(failure))
  }, [])

  const scheduler = useMemo(() => createDraftScheduler({
    write: (snapshot, contentVersion, lifetime) =>
      remote.writeDraft(sessionId, documentId, path, contentVersion, snapshot, lifetime),
    signal,
    onError: reportDraftError,
  }), [remote, sessionId, documentId, path, signal, reportDraftError])

  const clearDraftRecord = useCallback(async (lifetime: AbortSignal = signal): Promise<void> => {
    await scheduler.cancel()
    try {
      const result = await remote.clearDraft(sessionId, documentId, lifetime)
      if (!result.ok) setDraftError(mapOfficeFailure(result.error))
    } catch (error: unknown) {
      setDraftError({ code: 'unknown', message: messageOf(error) })
    }
  }, [scheduler, remote, sessionId, documentId, signal])

  useEffect(() => () => {
    void scheduler.cancel()
  }, [scheduler])

  useEffect(() => {
    if (legacy !== undefined || signal.aborted) return
    const controller = new AbortController()
    const lifetime = AbortSignal.any([signal, controller.signal])
    let live = true
    observedAtRead.current = observedRef.current
    setState(current => ({ ...current, status: 'loading', error: undefined }))
    void remote.read(sessionId, path, lifetime).then((result) => {
      if (!live) return
      if (!result.ok) {
        setState({ ...INITIAL, status: 'error', error: mapOfficeFailure(result.error) })
        return
      }
      setState({
        status: result.value.unsupported.length > 0 ? 'unsupported' : 'ready',
        snapshot: result.value.snapshot,
        version: result.value.version,
        unsupported: result.value.unsupported,
        error: undefined,
      })
      setExternalChange(false)
      setRecovery(undefined)
      setRecovered(false)
      if (result.value.unsupported.length > 0) return
      if (draftCheckedFor.current === documentId) return
      draftCheckedFor.current = documentId
      void remote.readDraft(sessionId, documentId, lifetime).then((draft) => {
        if (!live) return
        if (!draft.ok) {
          setDraftError(mapOfficeFailure(draft.error))
          return
        }
        if (draft.value === null) return
        setRecovery({
          savedAt: draft.value.savedAt,
          stale: draft.value.contentVersion !== result.value.version,
          snapshot: draft.value.snapshot,
        })
      }, (error: unknown) => {
        if (!live) return
        setDraftError({ code: 'unknown', message: messageOf(error) })
      })
    }, (error: unknown) => {
      if (!live) return
      setState({
        ...INITIAL,
        status: 'error',
        error: { code: 'unknown', message: messageOf(error) },
      })
    })
    return () => {
      live = false
      controller.abort()
    }
  }, [remote, sessionId, path, signal, generation, legacy, documentId])

  useEffect(() => {
    if (legacy === undefined || signal.aborted) return
    const controller = new AbortController()
    const lifetime = AbortSignal.any([signal, controller.signal])
    let live = true
    setState({ ...INITIAL, status: 'converting' })
    void remote.convertLegacy(sessionId, path, lifetime).then((result) => {
      if (!live) return
      if (!result.ok) {
        setState({ ...INITIAL, status: 'error', error: mapOfficeFailure(result.error) })
        return
      }
      const address = convertedAddress(sessionId, result.value, rootRef.current)
      if (address === undefined) {
        setState({
          ...INITIAL,
          status: 'error',
          error: {
            code: 'conversion-failed',
            message: `The converted copy of "${path}" is outside the Session workspace.`,
          },
        })
        return
      }
      openConverted(address, path)
    }, (error: unknown) => {
      if (!live) return
      setState({
        ...INITIAL,
        status: 'error',
        error: { code: 'unknown', message: messageOf(error) },
      })
    })
    return () => {
      live = false
      controller.abort()
    }
  }, [remote, sessionId, path, signal, legacy, openConverted])

  useEffect(() => {
    if (state.status === 'loading') return
    const verdict = externalChangeVerdict({
      observedVersion,
      version: state.version,
      observedAtRead: observedAtRead.current,
      dirty,
    })
    if (verdict === 'prompt') setExternalChange(true)
    else if (verdict === 'reload') setGeneration(value => value + 1)
  }, [observedVersion, state.status, state.version, dirty])

  const save = useCallback(async (edited?: OfficeDocumentSnapshot): Promise<OfficeSaveOutcome> => {
    const snapshot = state.snapshot
    const version = state.version
    if (snapshot === undefined || version === undefined) return 'refused'
    if (state.unsupported.length > 0 || externalChange || savingRef.current) return 'refused'
    savingRef.current = true
    setSaving(true)
    try {
      const result = await remote.save(sessionId, path, edited ?? snapshot, formatOf(snapshot), version, signal)
      if (signal.aborted) return 'refused'
      if (!result.ok) {
        const error = mapOfficeFailure(result.error)
        setState(current => ({
          ...current,
          error,
          status: error.code === 'conflict' ? 'conflict' : current.status,
        }))
        return error.code === 'conflict' ? 'conflict' : 'error'
      }
      setState(current => ({
        ...current,
        version: result.value.version,
        error: undefined,
        status: 'ready',
      }))
      setRecovery(undefined)
      setRecovered(false)
      setDraftError(undefined)
      await clearDraftRecord()
      return 'saved'
    } finally {
      savingRef.current = false
      setSaving(false)
    }
  }, [
    remote, sessionId, path, signal, state.snapshot, state.version, state.unsupported, externalChange,
    clearDraftRecord,
  ])

  const scheduleDraft = useCallback((snapshot: OfficeDocumentSnapshot): void => {
    if (statusRef.current !== 'ready' || versionRef.current === undefined) return
    scheduler.schedule(snapshot, versionRef.current)
  }, [scheduler])

  const acceptRecovery = useCallback((): void => {
    if (recovery === undefined) return
    setRecovery(undefined)
    setRecovered(true)
    setState(current => ({ ...current, snapshot: recovery.snapshot }))
  }, [recovery])

  const discardDraft = useCallback((): void => {
    void clearDraftRecord(new AbortController().signal)
  }, [clearDraftRecord])

  const discardRecovery = useCallback((): void => {
    setRecovery(undefined)
    discardDraft()
  }, [discardDraft])

  const dismissDraftError = useCallback((): void => { setDraftError(undefined) }, [])

  const reload = useCallback((): void => {
    setExternalChange(false)
    setGeneration(value => value + 1)
  }, [])

  const dismissExternalChange = useCallback((): void => { setExternalChange(false) }, [])

  const canSave = state.snapshot !== undefined
    && state.unsupported.length === 0
    && !saving
    && !externalChange
    && state.status !== 'conflict'
    && state.status !== 'error'
    && state.status !== 'loading'

  return {
    status: state.status,
    snapshot: state.snapshot,
    version: state.version,
    unsupported: state.unsupported,
    error: state.error,
    canSave,
    saving,
    externalChange,
    recovery,
    draftError,
    recovered,
    scheduleDraft,
    acceptRecovery,
    discardRecovery,
    dismissDraftError,
    discardDraft,
    save,
    reload,
    dismissExternalChange,
  }
}
