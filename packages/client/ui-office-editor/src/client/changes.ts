/**
 * Unsaved-edit, external-change, and moved-file state for open Office tabs.
 *
 * The panel and the tab close handler observe one plugin-owned snapshot store:
 * the panel marks a live unit dirty when the editor runtime reports a mutation,
 * and the close handler reads the same entry before the tab is removed. A close
 * attempt on a dirty tab is vetoed through the Sidebar's cleanup-failure path
 * and flags the entry so the panel can show the localized save-or-discard
 * prompt. The same plugin-owned surface follows the workspace file change feed:
 * when a tracked document's path disappears and its content version reappears
 * at another path, the document is rebound to that path so the Session records
 * `document/rebind` instead of losing the open document. The store is created
 * in `apply` and never at module scope.
 */
import { createSnapshotStore, type SnapshotStore } from '@forwork/client-store'
import type { TabId } from '@forwork/client-ui-dockkit'
import type { SessionId } from '@forwork/session/types'

/** One tab's unsaved state. */
export interface OfficeTabChangeState {
  /** The live unit holds edits the file does not. */
  readonly dirty: boolean
  /** A close attempt was vetoed while dirty; the panel shows its prompt. */
  readonly warned: boolean
}

/** Every tracked tab, keyed by tab id. */
export type OfficeTabChangeSnapshot = Readonly<Record<string, OfficeTabChangeState>>

/** Record one open document's new location after its file moved. */
export type OfficeDocumentRebind = (
  sessionId: SessionId,
  documentId: string,
  toPath: string,
) => void

/** One document the move tracker follows for one plugin instance. */
interface TrackedDocument {
  readonly sessionId: SessionId
  readonly documentId: string
  path: string
  /** File version last observed at `path`, which identifies a move. */
  version: string | undefined
  /** The path was reported absent and the document is waiting for its new location. */
  missing: boolean
}

/** One path a file feed reported present, keyed by Session. */
type ObservedPaths = Map<string, string>

/**
 * Normalize a path for comparison.
 * @param path - a workspace path from a tab address or file feed.
 * @returns the path with `\\` folded to `/`.
 */
function pathKey(path: string): string {
  return path.replace(/\\/g, '/')
}

/**
 * The open documents a plugin instance follows across file moves.
 *
 * A document becomes a rebind candidate only after the feed reports its path
 * absent. A later observation of the same file version at another path
 * resolves it, whether the disappearance or the reappearance arrives first.
 */
class DocumentMoves {
  private readonly documents = new Map<string, TrackedDocument>()
  private readonly observed = new Map<SessionId, ObservedPaths>()

  /**
   * @param rebind - records the move for one document; absent leaves the
   * tracker as an observation log for callers that do not record.
   */
  constructor(private readonly rebind: OfficeDocumentRebind | undefined) {}

  /**
   * Follow one open document at its current path.
   * @param sessionId - Session whose workspace owns the file.
   * @param documentId - Session-scoped document identity.
   * @param path - the path the tab currently addresses.
   */
  track(sessionId: SessionId, documentId: string, path: string): void {
    const current = this.documents.get(documentId)
    if (current !== undefined && current.sessionId === sessionId) {
      if (current.path !== pathKey(path)) {
        current.path = pathKey(path)
        current.missing = false
      }
      return
    }
    this.documents.set(documentId, {
      sessionId,
      documentId,
      path: pathKey(path),
      version: undefined,
      missing: false,
    })
  }

  /**
   * Stop following one document after its tab is removed.
   * @param documentId - the closed document.
   */
  forget(documentId: string): void {
    this.documents.delete(documentId)
  }

  /**
   * Apply one workspace file feed observation.
   * @param sessionId - Session whose workspace owns the file.
   * @param documentId - the tab's document identity.
   * @param path - the path the observation names.
   * @param version - the observed file version, or `undefined` when the path is gone.
   */
  observe(sessionId: SessionId, documentId: string, path: string, version: string | undefined): void {
    const document = this.documents.get(documentId)
    if (document === undefined || document.sessionId !== sessionId) return
    const key = pathKey(path)
    const seen = this.observedOf(sessionId)
    if (version === undefined) {
      seen.delete(key)
      if (document.path !== key) return
      document.missing = true
      const candidate = document.version === undefined
        ? undefined
        : firstObserved(seen, document.version, key)
      if (candidate !== undefined) this.rebindTo(document, candidate[0], candidate[1])
      return
    }
    seen.set(key, version)
    if (document.path === key) {
      document.version = version
      document.missing = false
    } else if (document.missing && document.version === version) {
      this.rebindTo(document, key, version)
      return
    } else if (document.missing) {
      return
    }
    const moved = this.missingDocument(sessionId, documentId, version, key)
    if (moved !== undefined) this.rebindTo(moved, key, version)
  }

  private observedOf(sessionId: SessionId): ObservedPaths {
    let seen = this.observed.get(sessionId)
    if (seen === undefined) {
      seen = new Map()
      this.observed.set(sessionId, seen)
    }
    return seen
  }

  private missingDocument(
    sessionId: SessionId,
    excludedDocumentId: string,
    version: string,
    path: string,
  ): TrackedDocument | undefined {
    for (const document of this.documents.values()) {
      if (document.sessionId !== sessionId || document.documentId === excludedDocumentId) continue
      if (document.missing && document.version === version && document.path !== path) return document
    }
    return undefined
  }

  private rebindTo(document: TrackedDocument, path: string, version: string): void {
    document.path = path
    document.version = version
    document.missing = false
    this.observedOf(document.sessionId).set(path, version)
    this.rebind?.(document.sessionId, document.documentId, path)
  }
}

/**
 * The first observed path other than `except` holding `version`.
 * @param seen - the Session's observed paths.
 * @param version - content version the moved document holds.
 * @param except - the path that disappeared.
 * @returns the path and version, or `undefined` when no other path matches.
 */
function firstObserved(
  seen: ObservedPaths,
  version: string,
  except: string,
): readonly [string, string] | undefined {
  for (const [path, observed] of seen) {
    if (path !== except && observed === version) return [path, observed]
  }
  return undefined
}

/** The plugin-owned change surface shared by the panel and the close handler. */
export interface OfficeTabChanges {
  /** Observable state the panel reads through the framework-bound hook. */
  readonly source: SnapshotStore<OfficeTabChangeSnapshot>
  /**
   * Record whether one tab's live unit has unsaved edits.
   * @param tabId - the tab whose unit changed.
   * @param dirty - the new state; clearing also clears a pending close warning.
   */
  markDirty(tabId: TabId, dirty: boolean): void
  /**
   * Flag a refused close while the tab is dirty.
   * @param tabId - the tab whose close was refused.
   */
  warnClose(tabId: TabId): void
  /**
   * Dismiss the close prompt without changing the dirty state.
   * @param tabId - the warned tab.
   */
  dismissWarning(tabId: TabId): void
  /**
   * Drop one tab's entry once its record is gone.
   * @param tabId - the closed tab.
   */
  forget(tabId: TabId): void
  /**
   * Whether one tab currently has unsaved edits.
   * @param tabId - the tab to inspect.
   * @returns `true` when the close handler must refuse removal.
   */
  isDirty(tabId: TabId): boolean
  /**
   * Follow one open document so a file move keeps its identity.
   * @param sessionId - Session whose workspace owns the file.
   * @param documentId - Session-scoped document identity.
   * @param path - the path the tab currently addresses.
   */
  trackDocument(sessionId: SessionId, documentId: string, path: string): void
  /**
   * Stop following one document after its tab is removed.
   * @param documentId - the closed document.
   */
  forgetDocument(documentId: string): void
  /**
   * Report one observation from the workspace file change feed.
   * @param sessionId - Session whose workspace owns the file.
   * @param documentId - the tab's document identity.
   * @param path - the observed path.
   * @param version - the observed file version, or `undefined` when the path is gone.
   */
  observeFile(sessionId: SessionId, documentId: string, path: string, version: string | undefined): void
}

/**
 * Create the change surface for one plugin instance.
 * @param rebind - records an observed move; production supplies `officeDocument.rebindDocument`.
 * @returns the observable source and its write set.
 */
export function createOfficeTabChanges(rebind?: OfficeDocumentRebind): OfficeTabChanges {
  const source = createSnapshotStore<OfficeTabChangeSnapshot>({})
  const moves = new DocumentMoves(rebind)

  const write = (tabId: TabId, next: OfficeTabChangeState): void => {
    source.set({ ...source.getSnapshot(), [tabId]: next })
  }
  return {
    source,
    markDirty(tabId, dirty) {
      const current = source.getSnapshot()[tabId]
      write(tabId, { dirty, warned: dirty ? current?.warned ?? false : false })
    },
    warnClose(tabId) {
      const current = source.getSnapshot()[tabId]
      if (current?.dirty !== true || current.warned) return
      write(tabId, { dirty: true, warned: true })
    },
    dismissWarning(tabId) {
      const current = source.getSnapshot()[tabId]
      if (current?.warned !== true) return
      write(tabId, { dirty: current.dirty, warned: false })
    },
    forget(tabId) {
      const current = source.getSnapshot()
      if (current[tabId] === undefined) return
      const next: Record<string, OfficeTabChangeState> = {}
      for (const [id, state] of Object.entries(current)) {
        if (id !== tabId) next[id] = state
      }
      source.set(next)
    },
    isDirty(tabId) {
      return source.getSnapshot()[tabId]?.dirty === true
    },
    trackDocument(sessionId, documentId, path) {
      moves.track(sessionId, documentId, path)
    },
    forgetDocument(documentId) {
      moves.forget(documentId)
    },
    observeFile(sessionId, documentId, path, version) {
      moves.observe(sessionId, documentId, path, version)
    },
  }
}

/** Inputs that decide how one external file version change is handled. */
export interface ExternalChangeInput {
  /** File version the resource provider currently reports; `undefined` before its first frame. */
  readonly observedVersion: string | undefined
  /** Version the snapshot currently displayed was read from. */
  readonly version: string | undefined
  /** Resource version observed when that read started, which tells a stale metadata frame from a real change. */
  readonly observedAtRead: string | undefined
  /** Whether the live unit has unsaved edits the reload would discard. */
  readonly dirty: boolean
}

/** What the hook must do about a reported external change. */
export type ExternalChangeVerdict = 'none' | 'reload' | 'prompt'

/**
 * Classify an external file change without discarding unsaved edits silently.
 * @param input - the read version, the resource's version, and the dirty state.
 * @returns `none`, an automatic `reload` for a clean unit, or a `prompt` while dirty.
 */
export function externalChangeVerdict(input: ExternalChangeInput): ExternalChangeVerdict {
  if (input.version === undefined || input.observedVersion === undefined) return 'none'
  if (input.observedVersion === input.version || input.observedVersion === input.observedAtRead) return 'none'
  return input.dirty ? 'prompt' : 'reload'
}
