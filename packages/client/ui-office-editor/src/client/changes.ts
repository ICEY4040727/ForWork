/**
 * Unsaved-edit and external-change state for open Office tabs.
 *
 * The panel and the tab close handler observe one plugin-owned snapshot store:
 * the panel marks a live unit dirty when the editor runtime reports a mutation,
 * and the close handler reads the same entry before the tab is removed. A close
 * attempt on a dirty tab is vetoed through the Sidebar's cleanup-failure path
 * and flags the entry so the panel can show the localized save-or-discard
 * prompt. The store is created in `apply` and never at module scope.
 */
import { createSnapshotStore, type SnapshotStore } from '@deepseek-ai/dsh-client-store'
import type { TabId } from '@deepseek-ai/dsh-client-ui-dockkit'

/** One tab's unsaved state. */
export interface OfficeTabChangeState {
  /** The live unit holds edits the file does not. */
  readonly dirty: boolean
  /** A close attempt was vetoed while dirty; the panel shows its prompt. */
  readonly warned: boolean
}

/** Every tracked tab, keyed by tab id. */
export type OfficeTabChangeSnapshot = Readonly<Record<string, OfficeTabChangeState>>

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
}

/**
 * Create the change surface for one plugin instance.
 * @returns the observable source and its write set.
 */
export function createOfficeTabChanges(): OfficeTabChanges {
  const source = createSnapshotStore<OfficeTabChangeSnapshot>({})
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
