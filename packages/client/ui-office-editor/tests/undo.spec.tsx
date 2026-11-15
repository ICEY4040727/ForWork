// @vitest-environment jsdom
/**
 * Native undo and redo through the adapter handle (SC-012).
 *
 * The fake runtime models a linear history so the handle's snapshot, dirty
 * transition, and history controls are observable without loading Univer or
 * constructing the Host command bridge that US2 adds.
 */
import { describe, expect, it, vi } from 'vitest'
import { createAdapter } from '../src/client/adapter.ts'
import type { OfficeRuntime, OfficeSnapshot } from '../src/client/adapter.ts'

const snapshot: OfficeSnapshot = { kind: 'workbook', data: { id: 'book' } }
const firstEdit: OfficeSnapshot = { kind: 'workbook', data: { id: 'book', a1: 'first' } }
const secondEdit: OfficeSnapshot = { kind: 'workbook', data: { id: 'book', a1: 'second' } }

/** A runtime whose history is a linear snapshot list the test can extend. */
function historyRuntime(): {
  runtime: OfficeRuntime
  edit(snapshot: OfficeSnapshot): void
  disposed: () => boolean
} {
  const states: OfficeSnapshot[] = [snapshot]
  let index = 0
  const listeners = new Set<() => void>()
  const notify = (): void => {
    for (const listener of [...listeners]) listener()
  }
  const runtime: OfficeRuntime = {
    getSnapshot: () => states[index] ?? snapshot,
    apply() { return { status: 'applied' } },
    undo() {
      if (index === 0) return
      index -= 1
      notify()
    },
    redo() {
      if (index >= states.length - 1) return
      index += 1
      notify()
    },
    onMutation(listener) {
      listeners.add(listener)
      return () => { listeners.delete(listener) }
    },
    dispose() { listeners.clear() },
  }
  return {
    runtime,
    edit(next) {
      states.length = index + 1
      states.push(next)
      index = states.length - 1
      notify()
    },
    disposed: () => listeners.size === 0,
  }
}

describe('adapter native history', () => {
  it('applies an edit, then undo and redo move the handle through the history', async () => {
    const history = historyRuntime()
    const adapter = createAdapter(async () => history.runtime)
    const onChange = vi.fn()
    const handle = await adapter.mount(document.createElement('div'), {
      kind: 'workbook',
      snapshot,
      onChange,
    })
    expect(handle.isDirty()).toBe(false)
    expect(handle.getSnapshot()).toBe(snapshot)

    history.edit(firstEdit)
    expect(handle.isDirty()).toBe(true)
    expect(handle.getSnapshot()).toBe(firstEdit)
    expect(onChange).toHaveBeenCalledOnce()

    handle.undo()
    expect(handle.getSnapshot()).toBe(snapshot)
    expect(handle.isDirty()).toBe(true)

    handle.redo()
    expect(handle.getSnapshot()).toBe(firstEdit)
    expect(handle.isDirty()).toBe(true)

    handle.clearDirty()
    expect(handle.isDirty()).toBe(false)
    history.edit(secondEdit)
    expect(handle.isDirty()).toBe(true)
    expect(onChange).toHaveBeenCalledTimes(2)
    expect(handle.getSnapshot()).toBe(secondEdit)
  })

  it('keeps the mounted snapshot when there is nothing left to undo', async () => {
    const history = historyRuntime()
    const adapter = createAdapter(async () => history.runtime)
    const handle = await adapter.mount(document.createElement('div'), {
      kind: 'workbook',
      snapshot,
      onChange: vi.fn(),
    })
    handle.undo()
    expect(handle.getSnapshot()).toBe(snapshot)
    handle.redo()
    expect(handle.getSnapshot()).toBe(snapshot)
    expect(handle.isDirty()).toBe(false)
  })

  it('stops reporting mutations after disposal', async () => {
    const history = historyRuntime()
    const adapter = createAdapter(async () => history.runtime)
    const onChange = vi.fn()
    const handle = await adapter.mount(document.createElement('div'), {
      kind: 'workbook',
      snapshot,
      onChange,
    })
    handle.dispose()
    history.edit(firstEdit)
    expect(onChange).not.toHaveBeenCalled()
    expect(history.disposed()).toBe(true)
  })
})
