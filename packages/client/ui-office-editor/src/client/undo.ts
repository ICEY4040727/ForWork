/**
 * The editor's native history, as the adapter handle exposes it.
 *
 * Undo and redo are the preset's own stack: the embedded preset header draws
 * the standard controls, and this binding reaches the same stack for the
 * adapter handle and, later, the Host command bridge. No second inverse stack
 * exists.
 */

/** The facade history operations the presets expose. */
export interface NativeHistoryApi {
  /** Undo the last edit on the focused unit. */
  undo(): Promise<boolean>
  /** Redo the last undone edit on the focused unit. */
  redo(): Promise<boolean>
}

/** Synchronous history controls the adapter handle forwards. */
export interface NativeHistory {
  /** Undo the last edit. */
  undo(): void
  /** Redo the last undone edit. */
  redo(): void
}

/**
 * Bind the facade's asynchronous history calls to the handle's synchronous surface.
 * @param api - the preset facade carrying `undo` and `redo`.
 * @returns the handle's history controls.
 */
export function nativeHistory(api: NativeHistoryApi): NativeHistory {
  return {
    undo: () => { void api.undo() },
    redo: () => { void api.redo() },
  }
}
