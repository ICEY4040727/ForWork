/**
 * Client half of the Host-to-renderer document command bridge.
 *
 * The Host tool appends `document/command` facts to the Session; this module
 * translates the session's event feed into correlated replies. One bridge
 * targets one live tab: it ignores commands for another document, refuses a
 * stale edit version, refuses unsaved user edits, applies the operation through
 * the adapter handle, and returns the post-command snapshot exactly once per
 * request. The transport and the Remote callback are injected, so the bridge
 * owns no Context and no module state.
 */
import type { DocumentEditOperation } from '@deepseek-ai/dsh-session/src/document-events.ts'
import type {
  OfficeCommandOperation,
  OfficeCommandResult,
  OfficeDocumentSnapshot,
  UnsupportedFeature,
} from '@deepseek-ai/dsh-office-exchange/types'

/** One `document/command` payload as the Host appends it to the Session. */
export interface DocumentCommandRequest {
  /** Correlation identity the Host resolves the pending tool call with. */
  readonly requestId: string
  /** Stable identity of the document the command targets. */
  readonly documentId: string
  /** Read request or edit steps to apply to the live unit. */
  readonly operation: OfficeCommandOperation
  /** Live document version the Host validated the command against. */
  readonly fromVersion: number
}

/** The live editor's answer to one command, in the Remote's wire vocabulary. */
export type DocumentCommandResult = OfficeCommandResult

/** What one adapter apply did to the live unit. */
export type OfficeApplyOutcome =
  | { readonly status: 'applied' }
  | { readonly status: 'unsupported-feature'; readonly features: readonly UnsupportedFeature[] }

/** The live tab state a bridge acts on; every member reads the current unit. */
export interface OfficeCommandTarget {
  /** Stable identity this tab has for its document. */
  readonly documentId: string
  /** The live numeric content version, or absence while no unit is mounted. */
  version(): number | undefined
  /** Whether an editor unit is mounted in this tab. */
  isOpen(): boolean
  /** Whether the live unit has unsaved user edits. */
  isDirty(): boolean
  /** Features a save would drop, as the read reported them. */
  unsupported(): readonly UnsupportedFeature[]
  /**
   * Apply one structured operation to the live unit.
   * @param operation - ordered steps the Host command carries.
   * @returns whether the unit applied every step, or the features that stop it.
   */
  apply(operation: DocumentEditOperation): OfficeApplyOutcome
  /** The post-command wire snapshot, or absence when the unit is gone. */
  snapshot(): OfficeDocumentSnapshot | undefined
}

/** One entry of the Session event window this bridge reads. */
export interface CommandEventEntry {
  /** Durable event versus a client-only transient row. */
  readonly type: string
  /** The event envelope. */
  readonly event: {
    /** Event discriminant. */
    readonly type: string
    /** Event payload; validated at this feed boundary. */
    readonly data?: unknown
  }
}

/** The window delta the bridge reacts to. */
export interface CommandEventChange {
  /** Delta discriminant; only `append` carries commands. */
  readonly kind: string
  /** Entries the delta added, when it added any. */
  readonly entries?: readonly CommandEventEntry[]
}

/**
 * The subset of the Session event source the bridge needs. The Session
 * controller's `eventSource` satisfies it structurally.
 */
export interface CommandEventFeed {
  /** @returns the current window and the delta that produced it. */
  getSnapshot(): {
    /** Latest window delta. */
    readonly change: CommandEventChange
  }
  /**
   * Observe window publication.
   * @param listener - invalidation callback.
   * @returns removes the listener.
   */
  subscribe(listener: () => void): () => void
}

/** The live command feed a bridge subscribes to. */
export interface DocumentCommandSource {
  /**
   * Observe newly appended commands.
   * @param listener - receives each validated request once.
   * @returns removes the subscription.
   */
  subscribe(listener: (request: DocumentCommandRequest) => void): () => void
}

/** The transport and Remote seam one bridge closes over. */
export interface CommandBridgeOptions {
  /** Newly appended commands for the Session. */
  readonly source: DocumentCommandSource
  /**
   * Return one correlated result to the Host. Called exactly once per request.
   * @param requestId - the command's identity.
   * @param result - the outcome the Host resolves the pending call with.
   */
  readonly complete: (requestId: string, result: DocumentCommandResult) => void
}

/** A bridge over one Session event feed. */
export interface CommandBridge {
  /**
   * Act on commands for one live tab until the returned disposer runs.
   * @param target - the live document state.
   * @returns stops listening for commands.
   */
  attach(target: OfficeCommandTarget): () => void
}

/**
 * The `officeDocument` Remote members this bridge uses.
 *
 * `completeCommand` is optional because the Host half may not have published it
 * yet; an absent method leaves the pending Host call to its timeout.
 */
export interface OfficeCommandRemote {
  /**
   * Existing `officeDocument` namespace member used only to identify the Remote
   * face; the bridge never calls it.
   */
  readonly read: (...args: never[]) => unknown
  /**
   * Resolve one pending Host command with the renderer's result.
   * @param sessionId - the Session whose command is being settled.
   * @param requestId - the correlation identity the Host minted.
   * @param result - the renderer's outcome.
   * @param signal - caller cancellation; an aborted call drops the answer.
   * @returns completion of the settlement.
   */
  completeCommand?(
    sessionId: string,
    requestId: string,
    result: DocumentCommandResult,
    signal?: AbortSignal,
  ): Promise<unknown>
}

/** The Session binding subset the bridge subscribes through. */
export interface SessionCommandBinding {
  /** The Session's live event window. */
  readonly eventSource: CommandEventFeed
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

/**
 * Whether one payload is a step kind the live adapter understands.
 *
 * Cell and paragraph steps are checked because the adapter reads their fields.
 * The remaining known spreadsheet steps are refused by the adapter before it
 * reads any field, so their payload is only required to name a known kind.
 */
function isEditStep(value: unknown): boolean {
  if (!isRecord(value)) return false
  switch (value['kind']) {
    case 'cell':
      return typeof value['sheetId'] === 'string'
        && typeof value['row'] === 'number'
        && typeof value['column'] === 'number'
        && 'value' in value
    case 'paragraph':
      return typeof value['index'] === 'number' && typeof value['text'] === 'string'
    case 'row':
    case 'column':
    case 'formula':
    case 'style':
      return true
    default:
      return false
  }
}

function isSteps(value: unknown): value is DocumentEditOperation {
  if (!Array.isArray(value)) return false
  return value.every((step: unknown) => isEditStep(step))
}

/**
 * Validate one operation against the command contract.
 * @param value - the raw operation payload.
 * @returns the operation, or `undefined` when the payload does not match.
 */
function commandOperation(value: unknown): OfficeCommandOperation | undefined {
  if (!isRecord(value)) return undefined
  if (value['kind'] === 'read') return { kind: 'read' }
  if (value['kind'] === 'edit' && isSteps(value['steps'])) return { kind: 'edit', steps: value['steps'] }
  return undefined
}

/**
 * Validate one event payload against the command contract.
 * @param data - the raw `document/command` payload.
 * @returns the request, or `undefined` when the payload does not match.
 */
function commandPayload(data: unknown): DocumentCommandRequest | undefined {
  if (!isRecord(data)) return undefined
  const { requestId, documentId, operation, fromVersion } = data
  if (typeof requestId !== 'string' || typeof documentId !== 'string') return undefined
  if (typeof fromVersion !== 'number') return undefined
  const parsed = commandOperation(operation)
  if (parsed === undefined) return undefined
  return { requestId, documentId, operation: parsed, fromVersion }
}

/**
 * Translate a Session event feed into newly appended command payloads.
 * @param feed - the Session's live event window.
 * @returns a source that emits each appended command once.
 */
export function sessionCommandSource(feed: CommandEventFeed): DocumentCommandSource {
  return {
    subscribe(listener) {
      return feed.subscribe(() => {
        const { change } = feed.getSnapshot()
        if (change.kind !== 'append') return
        for (const entry of change.entries ?? []) {
          if (entry.type !== 'event' || entry.event.type !== 'document/command') continue
          const request = commandPayload(entry.event.data)
          if (request !== undefined) listener(request)
        }
      })
    },
  }
}

/**
 * Attach one tab to a retained Session's command feed.
 * @param binding - the retained Session binding, or absence when it is not open.
 * @param sessionId - the Session whose commands are settled.
 * @param target - the live tab state and adapter the bridge acts on.
 * @param remote - the `officeDocument` Remote that settles the Host call.
 * @returns stops listening for commands; a no-op without a binding.
 */
export function attachSessionCommandBridge(
  binding: SessionCommandBinding | undefined,
  sessionId: string,
  target: OfficeCommandTarget,
  remote: OfficeCommandRemote,
): () => void {
  if (binding === undefined) return () => {}
  return createCommandBridge({
    source: sessionCommandSource(binding.eventSource),
    complete: (requestId, result) => {
      void remote.completeCommand?.(sessionId, requestId, result)
    },
  }).attach(target)
}

/**
 * Create a bridge over one Session's command feed.
 * @param options - the command source and the Remote settlement callback.
 * @returns the bridge the panel attaches its live target to.
 */
export function createCommandBridge(options: CommandBridgeOptions): CommandBridge {
  return {
    attach(target) {
      const handled = new Set<string>()
      const handle = (request: DocumentCommandRequest): void => {
        if (request.documentId !== target.documentId) return
        if (handled.has(request.requestId)) return
        handled.add(request.requestId)
        if (!target.isOpen()) {
          options.complete(request.requestId, { status: 'document-not-open' })
          return
        }
        const version = target.version()
        if (version === undefined) {
          options.complete(request.requestId, { status: 'document-not-open' })
          return
        }
        if (target.isDirty()) {
          options.complete(request.requestId, { status: 'pending-edits' })
          return
        }
        if (request.operation.kind === 'read') {
          const snapshot = target.snapshot()
          if (snapshot === undefined) {
            options.complete(request.requestId, { status: 'document-not-open' })
            return
          }
          options.complete(request.requestId, { status: 'read', snapshot, contentVersion: version })
          return
        }
        if (version !== request.fromVersion) {
          options.complete(request.requestId, { status: 'conflict', currentVersion: version })
          return
        }
        const unsupported = target.unsupported()
        if (unsupported.length > 0) {
          options.complete(request.requestId, { status: 'unsupported-feature', features: unsupported })
          return
        }
        const outcome = target.apply(request.operation.steps)
        if (outcome.status === 'unsupported-feature') {
          options.complete(request.requestId, outcome)
          return
        }
        const snapshot = target.snapshot()
        const contentVersion = target.version()
        if (snapshot === undefined || contentVersion === undefined) {
          options.complete(request.requestId, { status: 'document-not-open' })
          return
        }
        options.complete(request.requestId, { status: 'applied', snapshot, contentVersion })
      }
      return options.source.subscribe(handle)
    },
  }
}
