/**
 * Live editor lifecycle shared by the docs and sheets mounts.
 *
 * Both runtimes subscribe to the same facade events to report mutations and
 * release the same resources on teardown. The runtime-specific snapshot and
 * apply behavior stays in each mount.
 * @module @forwork/client-ui-office-editor/client/editor/lifecycle
 */

import { CommandType } from '@univerjs/core'
import type { Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/facade'
import type { OfficeRuntime } from '../adapter.ts'
import { nativeHistory } from '../undo.ts'

/** The adapter methods both Univer mounts share. */
export type EditorLifecycle = Pick<OfficeRuntime, 'undo' | 'redo' | 'onMutation' | 'dispose'>

/**
 * Subscribe to one Univer facade's mutation events and build the adapter lifecycle.
 * @param univer - Univer instance whose root the adapter owns.
 * @param api - preset facade that fires command, undo, and redo events.
 * @returns undo and redo controls, the mutation subscription, and ordered teardown.
 */
export function editorLifecycle(univer: Univer, api: FUniver): EditorLifecycle {
  const history = nativeHistory(api)
  const listeners = new Set<() => void>()
  const notify = (): void => {
    for (const listener of [...listeners]) listener()
  }
  const events = [
    api.addEvent(api.Event.CommandExecuted, (event) => {
      if (event.type === CommandType.MUTATION) notify()
    }),
    api.addEvent(api.Event.Undo, notify),
    api.addEvent(api.Event.Redo, notify),
  ]
  return {
    undo: () => { history.undo() },
    redo: () => { history.redo() },
    onMutation(listener) {
      listeners.add(listener)
      return () => { listeners.delete(listener) }
    },
    dispose() {
      for (const event of events) event.dispose()
      listeners.clear()
      // Univer's UI unmounts its own nested root; removing its children here
      // could race that unmount when React defers it.
      univer.dispose()
    },
  }
}
