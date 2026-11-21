/**
 * The open document's Remote-backed state.
 *
 * Reads run through the injected `officeDocument` Remote — never the Host
 * service — and a save sends the live snapshot with the version the read
 * observed. A refused version is reported as a conflict. Preserved content the
 * editor cannot change never blocks a save; the writer refuses only an edit
 * that intersects a coupled construct. An external version change reloads a
 * clean unit and prompts instead of reloading a dirty one.
 *
 * The content version is one value shared by the Host, agent commands, and local
 * saves. When the Host's own save advances the `document` projection's version,
 * the hook adopts it and treats the matching file reload as the Host's write
 * rather than an external edit, so a later command still reports the version the
 * Host recorded. A local save publishes its new file version and numeric content
 * version up front; the matching file-feed frame and projection rise are adopted
 * in place, so the mounted unit survives its own save.
 *
 * An agent command this tab applied already put the saved content in the live
 * unit. The save that follows it publishes the version the command reported, so
 * the hook adopts the new file version in place instead of re-reading: the unit
 * stays mounted and the editor's native undo history keeps the agent edit. A
 * file change no applied command accounts for still reloads or prompts.
 *
 * Edits are captured as recovery drafts through the same Remote: the first
 * edit after a clean state is written immediately and later edits are
 * idle-debounced. A tab close and a page unload flush the pending draft, so an
 * edit already in the window is on disk before the tab or page goes away. A
 * draft found after a successful read is offered only as a recovery choice: it
 * is never applied automatically, and a draft captured from an older file
 * version is labelled stale so the user chooses it explicitly. A successful
 * save and a discard both clear the stored draft.
 *
 * A legacy `.doc`, `.xls`, or `.ppt` address is never read: it is converted
 * through the same Remote, and the converted copy opens in a new tab in this
 * one's place. The original is never modified, and this body reports
 * conversion state until the new tab takes over. The read's detected format is
 * carried with the content, so a macro-enabled `.docm` or `.xlsm` save keeps
 * the container it was opened from instead of inferring one from the snapshot
 * kind.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { SessionId } from '@forwork/session/types'
import type { RemoteFailure, RemoteResult } from '@forwork/typert-protocol'
import type {
  DocumentSaveAttribution,
  OfficeDocumentDraft,
  OfficeDocumentDraftWriteResult,
  OfficeDocumentReadResult,
  OfficeDocumentSaveResult,
  OfficeDocumentSnapshot,
  OfficeFormat,
  PreservedConstructKind,
} from '@forwork/office-exchange/types'
import { externalChangeVerdict } from './changes.ts'
import { createDraftScheduler, draftRecoveryDecision } from './drafts.ts'
import { convertedAddress, legacyFormatOf } from './target.ts'

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
   * @param attribution - who caused the save; a user save names the user.
   * @param signal - cancels the call.
   * @returns the version the write published, or the declared failure.
   */
  save(
    sessionId: SessionId,
    path: string,
    snapshot: OfficeDocumentSnapshot,
    format: OfficeFormat,
    expectedVersion: string,
    attribution: DocumentSaveAttribution,
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
  /** `offer` when the draft matches the file version, `stale` when it does not. */
  readonly decision: 'offer' | 'stale'
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
export type OfficeDocumentStatus = 'loading' | 'converting' | 'ready' | 'conflict' | 'error'

/** One open document's observable state. */
export interface OfficeDocumentState {
  /** Load, edit-ready, refusal, save-conflict, or failed state. */
  readonly status: OfficeDocumentStatus
  /** Format the read detected; a save writes this exact container. */
  readonly format: OfficeFormat | undefined
  /** The read snapshot; absent while loading or after a failed read. */
  readonly snapshot: OfficeDocumentSnapshot | undefined
  /** Version the snapshot was read from; absent until a read succeeds. */
  readonly version: string | undefined
  /** Construct classes the file carries verbatim; presence never blocks a save. */
  readonly preserved: readonly PreservedConstructKind[]
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
  /** Workspace path of the open file; a rebind moves it to the file's new location. */
  readonly path: string
  /**
   * Session-scoped identity of the open document. It stays fixed when a rebind
   * changes the path, so drafts and command targets keep naming one document.
   */
  readonly documentId: string
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
  /**
   * Host-reported content version for this document, from the `document`
   * projection. It rises on a Host save, which is how the hook distinguishes
   * the Host's own write from an external edit.
   */
  readonly contentVersion?: number | undefined
}

/** The hook's state plus the actions the panel calls. */
export interface UseOfficeDocumentResult {
  /** Current visible state. */
  readonly status: OfficeDocumentStatus
  /** Read snapshot. */
  readonly snapshot: OfficeDocumentSnapshot | undefined
  /** Version of the displayed snapshot. */
  readonly version: string | undefined
  /**
   * Content version the live unit holds. Host saves, agent commands, and local
   * saves advance this one value, so every writer reports the same version.
   */
  readonly contentVersion: number
  /** Construct classes the file carries verbatim; shown without disabling save. */
  readonly preserved: readonly PreservedConstructKind[]
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
   * Capture the live unit as a recovery draft. The first edit after a clean
   * state is written immediately; later edits wait for the idle period.
   * @param snapshot - the live unit snapshot to capture.
   */
  scheduleDraft(snapshot: OfficeDocumentSnapshot): void
  /**
   * Write the pending recovery draft now, for a tab close or page unload.
   * @returns settles after the pending and in-flight writes.
   */
  flushDraft(): Promise<void>
  /** Replace the live unit with the offered recovery draft. */
  acceptRecovery(): void
  /** Dismiss the offered draft and remove it from the draft store. */
  discardRecovery(): void
  /** Hide the last draft failure message. */
  dismissDraftError(): void
  /** Drop any pending draft and remove the stored draft without closing. */
  discardDraft(): void
  /**
   * Read the content version without waiting for a render.
   * @returns the live content version.
   */
  currentContentVersion(): number
  /**
   * Advance the content version after the adapter applied an agent command and
   * expect the Host's save of that command's snapshot to publish it.
   */
  markApplied(): void
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
  format: undefined,
  snapshot: undefined,
  version: undefined,
  preserved: [],
  error: undefined,
}

/**
 * Preserved classes the read reported.
 *
 * The Exchange read reports preserved classes alongside the snapshot. A Host
 * that does not report them yields an empty list, which shows no marker rather
 * than inventing preserved content.
 * @param result - the Remote read result.
 * @returns reported preserved classes, or an empty list.
 */
function preservedOf(result: OfficeDocumentReadResult): readonly PreservedConstructKind[] {
  const reported = result as OfficeDocumentReadResult
    & { readonly preserved?: readonly PreservedConstructKind[] }
  return reported.preserved ?? []
}

/** Render a non-Error rejection as a diagnostic message. */
function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

/**
 * State reported when an effect's promise rejects without a Remote failure.
 * @param error - the rejection value.
 * @returns error state carrying the diagnostic message.
 */
function unknownFailureState(error: unknown): OfficeDocumentState {
  return {
    ...INITIAL,
    status: 'error',
    error: { code: 'unknown', message: messageOf(error) },
  }
}

/** One effect-scoped lifetime shared by a Remote call and its teardown. */
interface EffectLifetime {
  /** Lifetime passed to the Remote call. */
  readonly signal: AbortSignal
  /** Whether the effect still owns its result. */
  isLive(): boolean
  /** Mark the effect unmounted and abort its lifetime. */
  end(): void
}

/**
 * Start one effect's cancellable lifetime.
 * @param signal - tab lifetime the effect must also honor.
 * @returns the combined lifetime and its teardown controls.
 */
function effectLifetime(signal: AbortSignal): EffectLifetime {
  const controller = new AbortController()
  let live = true
  return {
    signal: AbortSignal.any([signal, controller.signal]),
    isLive: () => live,
    end: () => {
      live = false
      controller.abort()
    },
  }
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
  const { remote, sessionId, path, documentId, signal, observedVersion, workspaceRoot, openConverted, dirty } = input
  const legacy = legacyFormatOf(path)
  const [state, setState] = useState<OfficeDocumentState>(INITIAL)
  const [saving, setSaving] = useState(false)
  const [externalChange, setExternalChange] = useState(false)
  const [generation, setGeneration] = useState(0)
  const [recovery, setRecovery] = useState<OfficeDraftRecovery | undefined>(undefined)
  const [draftError, setDraftError] = useState<OfficeDocumentFailure | undefined>(undefined)
  const [recovered, setRecovered] = useState(false)
  const [contentVersionState, setContentVersionState] = useState(0)
  const savingRef = useRef(false)
  const observedRef = useRef(observedVersion)
  const observedAtRead = useRef<string | undefined>(undefined)
  const rootRef = useRef(workspaceRoot)
  const pathRef = useRef(path)
  const draftCheckedFor = useRef<string | undefined>(undefined)
  const statusRef = useRef(state.status)
  const fileVersionRef = useRef(state.version)
  const contentVersionRef = useRef(0)
  const hostVersionRef = useRef(input.contentVersion)
  const lastHostVersion = useRef(input.contentVersion)
  const hostSaved = useRef(false)
  /** Version an accepted agent command reported, whose Host save this tab expects. */
  const selfSaveVersion = useRef<number | undefined>(undefined)
  /**
   * Identity of a local save whose own projection rise and file transition are
   * not external changes. `from` is the file version the save replaced; `to` is
   * the version it published; `contentVersion` is the numeric version this tab
   * advanced. The flags record which of the two matching events arrived.
   */
  const localSave = useRef<{
    readonly from: string
    readonly to: string
    readonly contentVersion: number
    projectionSeen: boolean
    fileSeen: boolean
  } | undefined>(undefined)
  observedRef.current = observedVersion
  rootRef.current = workspaceRoot
  pathRef.current = path
  statusRef.current = state.status
  fileVersionRef.current = state.version
  hostVersionRef.current = input.contentVersion
  const contentVersion = Math.max(contentVersionState, input.contentVersion ?? 0)
  const currentContentVersion = useCallback(
    () => Math.max(contentVersionRef.current, hostVersionRef.current ?? 0),
    [],
  )
  const advanceContentVersion = useCallback((): number => {
    const next = Math.max(contentVersionRef.current, hostVersionRef.current ?? 0) + 1
    contentVersionRef.current = next
    setContentVersionState(next)
    return next
  }, [])
  const markApplied = useCallback((): void => {
    selfSaveVersion.current = advanceContentVersion()
  }, [advanceContentVersion])

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
    // Both events run while the page can still start a write: `beforeunload`
    // for a reload or navigation and `pagehide` for a tab close or an app quit.
    const flush = (): void => { void scheduler.flush() }
    window.addEventListener('beforeunload', flush)
    window.addEventListener('pagehide', flush)
    return () => {
      window.removeEventListener('beforeunload', flush)
      window.removeEventListener('pagehide', flush)
    }
  }, [scheduler])

  useEffect(() => {
    if (legacy !== undefined || signal.aborted) return
    const lifetime = effectLifetime(signal)
    // A fresh read is the new base for the unit, so any save a command staged
    // before it no longer describes this content.
    selfSaveVersion.current = undefined
    observedAtRead.current = observedRef.current
    setState(current => ({ ...current, status: 'loading', error: undefined }))
    // A rebind moves `path` without re-reading: the live unit keeps its edits,
    // and the read path is taken at read time so an explicit reload uses the
    // file's current location.
    void remote.read(sessionId, pathRef.current, lifetime.signal).then((result) => {
      if (!lifetime.isLive()) return
      if (!result.ok) {
        setState({ ...INITIAL, status: 'error', error: mapOfficeFailure(result.error) })
        return
      }
      setState({
        status: 'ready',
        format: result.value.format,
        snapshot: result.value.snapshot,
        version: result.value.version,
        preserved: preservedOf(result.value),
        error: undefined,
      })
      setExternalChange(false)
      setRecovery(undefined)
      setRecovered(false)
      if (draftCheckedFor.current === documentId) return
      draftCheckedFor.current = documentId
      void remote.readDraft(sessionId, documentId, lifetime.signal).then((draft) => {
        if (!lifetime.isLive()) return
        if (!draft.ok) {
          setDraftError(mapOfficeFailure(draft.error))
          return
        }
        const stored = draft.value
        if (stored === null) return
        const decision = draftRecoveryDecision(stored, result.value.version)
        /* v8 ignore next -- a non-null stored draft always decides offer or stale. */
        if (decision === 'none') return
        setRecovery({
          savedAt: stored.savedAt,
          decision,
          snapshot: stored.snapshot,
        })
      }, (error: unknown) => {
        if (!lifetime.isLive()) return
        setDraftError({ code: 'unknown', message: messageOf(error) })
      })
    }, (error: unknown) => {
      if (!lifetime.isLive()) return
      setState(unknownFailureState(error))
    })
    return () => { lifetime.end() }
  }, [remote, sessionId, signal, generation, legacy, documentId])

  useEffect(() => {
    // A rebind can change the suffix. Converting then would replace the tab and
    // drop live edits, so conversion waits until the unit is clean.
    if (legacy === undefined || signal.aborted || dirty) return
    const lifetime = effectLifetime(signal)
    setState({ ...INITIAL, status: 'converting' })
    void remote.convertLegacy(sessionId, path, lifetime.signal).then((result) => {
      if (!lifetime.isLive()) return
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
      if (!lifetime.isLive()) return
      setState(unknownFailureState(error))
    })
    return () => { lifetime.end() }
  }, [remote, sessionId, path, signal, legacy, openConverted, dirty])

  useEffect(() => {
    const reported = input.contentVersion
    const local = localSave.current
    const rose = reported !== undefined
      && (lastHostVersion.current === undefined || reported > lastHostVersion.current)
    // A projection rise this tab's own save advanced is its own fact. Any other
    // rise is a Host save this tab did not make.
    if (rose && reported !== local?.contentVersion) hostSaved.current = true
    lastHostVersion.current = reported
    if (state.status === 'loading') return
    // A save in flight may be the write the resource is reporting; its result
    // decides whether the new version is this tab's own.
    if (saving) return
    if (local !== undefined) {
      if (reported !== undefined && reported === local.contentVersion) {
        local.projectionSeen = true
        hostSaved.current = false
      }
      if (observedVersion === local.to) {
        // The watched file moved to the version this tab just published: the
        // live unit already holds that content, so adopt it in place.
        local.fileSeen = true
        observedAtRead.current = observedVersion
        setExternalChange(false)
      } else if (observedVersion === local.from) {
        if (reported !== undefined && reported > local.contentVersion) {
          // Another save advanced the projection past this tab's, so this
          // transition is not the one the local save published.
          localSave.current = undefined
        } else {
          // A frame from before this tab's save; it reports no external change.
          return
        }
      } else if (observedVersion !== undefined) {
        // A version the save neither replaced nor published is a new change.
        localSave.current = undefined
      }
      if (local.projectionSeen && local.fileSeen) localSave.current = undefined
    }
    const verdict = externalChangeVerdict({
      observedVersion,
      version: state.version,
      observedAtRead: observedAtRead.current,
      dirty,
    })
    if (verdict === 'none') {
      if (observedVersion !== undefined && observedVersion === state.version) {
        // The resource now reports the displayed content, so a later frame can
        // be classified against this version instead of the read's first one.
        observedAtRead.current = observedVersion
      }
      return
    }
    const expected = selfSaveVersion.current
    if (expected !== undefined) {
      // The expected version comes from the same `document` projection the Host
      // reads, so it is the evidence that this tab's save landed. While the
      // projection is absent or behind, the change is held rather than
      // reloaded: the version guard still refuses a save at the stale version.
      if (dirty) {
        selfSaveVersion.current = undefined
      } else if (reported === expected) {
        // The Host saved the snapshot this tab's agent command returned. The
        // live unit already holds that content: adopt the published version in
        // place so the mounted unit and its native undo history survive.
        selfSaveVersion.current = undefined
        hostSaved.current = false
        observedAtRead.current = observedVersion
        setExternalChange(false)
        setState(current => ({ ...current, version: observedVersion }))
        return
      } else if (reported !== undefined && reported > expected) {
        // Another save advanced past this command, so this change is not it.
        selfSaveVersion.current = undefined
      } else {
        return
      }
    }
    // A version the Host advanced itself wrote the file: reload its content
    // without presenting the change as somebody else's edit.
    if (hostSaved.current && !dirty) {
      hostSaved.current = false
      setExternalChange(false)
      setGeneration(value => value + 1)
      return
    }
    hostSaved.current = false
    if (verdict === 'prompt') setExternalChange(true)
    else setGeneration(value => value + 1)
  }, [observedVersion, state.status, state.version, dirty, input.contentVersion, saving])

  const save = useCallback(async (edited?: OfficeDocumentSnapshot): Promise<OfficeSaveOutcome> => {
    const snapshot = state.snapshot
    const version = state.version
    const format = state.format
    if (snapshot === undefined || version === undefined || format === undefined) return 'refused'
    if (externalChange || savingRef.current) return 'refused'
    savingRef.current = true
    setSaving(true)
    try {
      const result = await remote.save(
        sessionId, path, edited ?? snapshot, format, version, { actor: 'user' }, signal,
      )
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
      // This write published the live unit itself; any staged command save it
      // supersedes no longer describes the file. The published version and the
      // numeric version this tab advances arrive back through the file feed and
      // the `document` projection, so record them as this tab's own change.
      selfSaveVersion.current = undefined
      localSave.current = {
        from: version,
        to: result.value.version,
        contentVersion: advanceContentVersion(),
        projectionSeen: false,
        fileSeen: false,
      }
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
    remote, sessionId, path, signal, state.snapshot, state.version, state.format, externalChange,
    clearDraftRecord, advanceContentVersion,
  ])

  const scheduleDraft = useCallback((snapshot: OfficeDocumentSnapshot): void => {
    if (statusRef.current !== 'ready' || fileVersionRef.current === undefined) return
    scheduler.schedule(snapshot, fileVersionRef.current)
  }, [scheduler])

  const flushDraft = useCallback((): Promise<void> => scheduler.flush(), [scheduler])

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
    && !saving
    && !externalChange
    && state.status !== 'conflict'
    && state.status !== 'error'
    && state.status !== 'loading'

  return {
    status: state.status,
    snapshot: state.snapshot,
    version: state.version,
    contentVersion,
    preserved: state.preserved,
    error: state.error,
    canSave,
    saving,
    externalChange,
    recovery,
    draftError,
    recovered,
    scheduleDraft,
    flushDraft,
    acceptRecovery,
    discardRecovery,
    dismissDraftError,
    discardDraft,
    currentContentVersion,
    markApplied,
    save,
    reload,
    dismissExternalChange,
  }
}
