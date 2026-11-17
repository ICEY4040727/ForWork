/**
 * Recovery-draft capture for one open document.
 *
 * The first edit after a clean state — a fresh document, a save, or a discard —
 * writes its candidate through the injected `officeDocument.writeDraft`
 * immediately, so a crash cannot lose the first keystroke. Later edits only
 * replace the candidate, which is written once the idle period elapses without
 * another edit. Drafts never touch the document's own file, a successful save
 * clears the stored draft, and a write failure is forwarded to the owning hook
 * instead of disappearing. Each panel creates one scheduler; there is no
 * module-level instance.
 */
import { RemoteError } from '@forwork/typert-protocol'
import type { RemoteFailure, RemoteResult } from '@forwork/typert-protocol'
import type {
  OfficeDocumentDraft,
  OfficeDocumentDraftWriteResult,
  OfficeDocumentSnapshot,
} from '@forwork/office-exchange/types'

/** Idle period after the last edit before its draft is written. */
export const DRAFT_IDLE_MS = 2000

/** What the editor may do with one stored draft after a successful read. */
export type DraftRecoveryDecision = 'none' | 'offer' | 'stale'

/**
 * Decide how a stored draft relates to the file a successful read just returned.
 *
 * This mirrors the Host rule in `@forwork/session-document`'s
 * `document-recovery` module: the Host and this client cannot share a value
 * module across the planes, so both sides keep the same three-way decision and
 * the panel exposes only its result.
 * @param draft - the stored draft, or `null` when the document has none.
 * @param currentVersion - the version the successful read observed.
 * @returns `none` without a draft, `offer` when the versions match, `stale` otherwise.
 */
export function draftRecoveryDecision(
  draft: OfficeDocumentDraft | null,
  currentVersion: string,
): DraftRecoveryDecision {
  if (draft === null) return 'none'
  return draft.contentVersion === currentVersion ? 'offer' : 'stale'
}

/** One draft write bound to its document identity and caller lifetime. */
export type OfficeDraftWrite = (
  snapshot: OfficeDocumentSnapshot,
  contentVersion: string,
  signal: AbortSignal,
) => Promise<RemoteResult<OfficeDocumentDraftWriteResult>>

/** Everything one scheduler needs from its owning hook. */
export interface OfficeDraftSchedulerInput {
  /** Writes the captured snapshot at the version the read observed. */
  readonly write: OfficeDraftWrite
  /** Caller lifetime; an aborted lifetime stops new writes. */
  readonly signal: AbortSignal
  /** Receives every refused or rejected write. */
  readonly onError: (failure: RemoteFailure) => void
  /** Idle period in milliseconds; defaults to {@link DRAFT_IDLE_MS}. */
  readonly idleMs?: number
}

/** The capture surface the hook schedules edits through. */
export interface OfficeDraftScheduler {
  /**
   * Capture one edit. The first edit after a clean state writes immediately;
   * every later edit replaces the pending candidate and restarts the idle
   * period.
   * @param snapshot - the live unit snapshot after the edit.
   * @param contentVersion - the file version the edit started from.
   */
  schedule(snapshot: OfficeDocumentSnapshot, contentVersion: string): void
  /**
   * Write the pending candidate now, stop the idle period, and wait for any
   * write already running.
   * @returns settles after the write attempt, including a forwarded failure.
   */
  flush(): Promise<void>
  /**
   * Drop the pending candidate, restore the leading-edge first edit, then wait
   * for an in-flight write.
   * @returns settles after the in-flight write, so a later clear cannot race it.
   */
  cancel(): Promise<void>
  /** Whether a candidate is waiting for the idle period to elapse. */
  isPending(): boolean
}

/** One candidate waiting for the idle period. */
interface DraftCandidate {
  readonly snapshot: OfficeDocumentSnapshot
  readonly contentVersion: string
}

/** Report a non-Error rejection the same way the hook reports every unknown failure. */
function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

/**
 * Create a scheduler over one document identity.
 * @param input - the write call, caller lifetime, failure sink, and idle period.
 * @returns the debounced scheduler the hook owns.
 */
export function createDraftScheduler(input: OfficeDraftSchedulerInput): OfficeDraftScheduler {
  const idleMs = input.idleMs ?? DRAFT_IDLE_MS
  let timer: ReturnType<typeof setTimeout> | undefined
  let candidate: DraftCandidate | undefined
  let inFlight: Promise<void> | undefined
  /** The next edit is the first after a clean state and writes without waiting. */
  let leading = true

  const stopTimer = (): void => {
    if (timer === undefined) return
    clearTimeout(timer)
    timer = undefined
  }

  const write = (): Promise<void> => {
    const capture = candidate
    if (capture === undefined) return Promise.resolve()
    candidate = undefined
    if (input.signal.aborted) return Promise.resolve()
    const started = (async (): Promise<void> => {
      try {
        const result = await input.write(capture.snapshot, capture.contentVersion, input.signal)
        if (!result.ok) input.onError(result.error)
      } catch (error: unknown) {
        input.onError(new RemoteError('gateway/internal', messageOf(error), {}))
      }
    })()
    inFlight = started
    void started.finally(() => {
      if (inFlight === started) inFlight = undefined
    })
    return started
  }

  return {
    schedule(snapshot, contentVersion) {
      if (input.signal.aborted) return
      candidate = { snapshot, contentVersion }
      stopTimer()
      if (leading) {
        leading = false
        void write()
        return
      }
      timer = setTimeout(() => {
        timer = undefined
        void write()
      }, idleMs)
    },
    async flush() {
      stopTimer()
      await write()
      if (inFlight !== undefined) await inFlight
    },
    async cancel() {
      stopTimer()
      candidate = undefined
      leading = true
      if (inFlight !== undefined) await inFlight
    },
    isPending: () => candidate !== undefined,
  }
}
