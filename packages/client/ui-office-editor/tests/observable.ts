/**
 * Driven fixture: bind a plugin-owned observable the way the slot renderer does.
 *
 * Component specs feed the framework-bound selector hooks directly. Subscribing
 * to the real source keeps a store write (a dismissal, a newly opened tab)
 * observable to the rendered component without building render machinery.
 */
import { useSyncExternalStore } from 'react'
import type { SnapshotSelectorHook } from '@deepseek-ai/dsh-client-store'
import type { HostObservable } from '@deepseek-ai/dsh-client-ui-slots'

/**
 * Bind one observable as a selector hook.
 * @param source - the observable the renderer would bind.
 * @returns a hook over the source's live snapshot.
 */
export function bindObservable<T>(source: HostObservable<T>): SnapshotSelectorHook<T> {
  return function useBound<Selected>(selector: (value: T) => Selected): Selected {
    return useSyncExternalStore(listener => source.subscribe(listener), () => selector(source.getSnapshot()))
  }
}
