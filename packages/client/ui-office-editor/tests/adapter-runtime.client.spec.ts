// @vitest-environment jsdom
/**
 * The adapter seam over an injected runtime factory.
 *
 * The fake factory stands in for the lazily imported Univer modules, so this
 * suite verifies mount, snapshot, history, dirty transitions, and teardown
 * without loading the editor runtime.
 */
import { describe, expect, it, vi } from 'vitest'
import {
  createAdapter,
  createMockAdapter,
  type OfficeRuntime,
  type OfficeRuntimeMount,
  type OfficeSnapshot,
} from '../src/client/adapter.ts'
import { nativeHistory } from '../src/client/undo.ts'

const snapshot: OfficeSnapshot = { kind: 'workbook', data: { id: 'book' } }
const edited: OfficeSnapshot = { kind: 'workbook', data: { id: 'book', name: 'edited' } }

function fakeRuntime(): {
  runtime: OfficeRuntime
  mounts: OfficeRuntimeMount[]
  undo: ReturnType<typeof vi.fn>
  redo: ReturnType<typeof vi.fn>
  dispose: ReturnType<typeof vi.fn>
  emit(): void
} {
  const mounts: OfficeRuntimeMount[] = []
  const listeners = new Set<() => void>()
  const undo = vi.fn()
  const redo = vi.fn()
  const dispose = vi.fn()
  const runtime: OfficeRuntime = {
    getSnapshot: () => edited,
    undo,
    redo,
    onMutation(listener) {
      listeners.add(listener)
      return () => { listeners.delete(listener) }
    },
    dispose,
  }
  return {
    runtime,
    mounts,
    undo,
    redo,
    dispose,
    emit() {
      for (const listener of [...listeners]) listener()
    },
  }
}

describe('createAdapter', () => {
  it('mounts through the factory and delegates snapshot, history, and dirty state', async () => {
    const fake = fakeRuntime()
    const factory = vi.fn(async (mount: OfficeRuntimeMount) => {
      fake.mounts.push(mount)
      mount.container.appendChild(mount.container.ownerDocument.createElement('div'))
      return fake.runtime
    })
    const adapter = createAdapter(factory)
    const container = document.createElement('div')
    const onChange = vi.fn()
    const onEdit = vi.fn()
    const handle = await adapter.mount(container, { kind: 'workbook', snapshot, onChange, onEdit })
    expect(factory).toHaveBeenCalledOnce()
    expect(fake.mounts[0]?.kind).toBe('workbook')
    expect(fake.mounts[0]?.snapshot).toBe(snapshot)
    expect(fake.mounts[0]?.container).toBe(container)
    expect(handle.getSnapshot()).toBe(edited)
    expect(handle.isDirty()).toBe(false)
    fake.emit()
    expect(handle.isDirty()).toBe(true)
    expect(onChange).toHaveBeenCalledOnce()
    expect(onEdit).toHaveBeenCalledOnce()
    fake.emit()
    expect(onChange).toHaveBeenCalledOnce()
    expect(onEdit).toHaveBeenCalledTimes(2)
    handle.clearDirty()
    expect(handle.isDirty()).toBe(false)
    fake.emit()
    expect(onChange).toHaveBeenCalledTimes(2)
    expect(onEdit).toHaveBeenCalledTimes(3)
    handle.undo()
    handle.redo()
    expect(fake.undo).toHaveBeenCalledOnce()
    expect(fake.redo).toHaveBeenCalledOnce()
  })

  it('disposes the runtime, unsubscribes mutations, and clears the container', async () => {
    const fake = fakeRuntime()
    const adapter = createAdapter(async () => fake.runtime)
    const container = document.createElement('div')
    container.appendChild(document.createElement('span'))
    const handle = await adapter.mount(container, { kind: 'document', snapshot, onChange: vi.fn() })
    handle.dispose()
    expect(fake.dispose).toHaveBeenCalledOnce()
    expect(container.childElementCount).toBe(0)
    fake.emit()
    expect(handle.isDirty()).toBe(false)
    handle.dispose()
    expect(fake.dispose).toHaveBeenCalledOnce()
  })

  it('propagates a runtime that cannot start', async () => {
    const adapter = createAdapter(async () => { throw new Error('no canvas') })
    await expect(adapter.mount(document.createElement('div'), {
      kind: 'document', snapshot, onChange: vi.fn(),
    })).rejects.toThrow('no canvas')
  })
})

describe('createMockAdapter', () => {
  it('records mounts and drives dirty state through the latest handle', async () => {
    const adapter = createMockAdapter()
    adapter.setDirty(true)
    expect(adapter.mounts).toEqual([])
    const container = document.createElement('div')
    const onChange = vi.fn()
    const handle = await adapter.mount(container, { kind: 'document', snapshot, onChange })
    expect(adapter.mounts).toHaveLength(1)
    expect(container.dataset.officeMount).toBe('document')
    adapter.setDirty(true)
    expect(handle.isDirty()).toBe(true)
    expect(onChange).toHaveBeenCalledOnce()
    handle.clearDirty()
    expect(handle.isDirty()).toBe(false)
    adapter.setDirty(false)
    expect(onChange).toHaveBeenCalledTimes(2)
    handle.undo()
    handle.redo()
    expect(handle.getSnapshot()).toBe(snapshot)
    handle.dispose()
    expect(container.dataset.officeMount).toBeUndefined()
  })
})

describe('nativeHistory', () => {
  it('forwards undo and redo to the preset facade', async () => {
    const undo = vi.fn(async () => true)
    const redo = vi.fn(async () => false)
    const history = nativeHistory({ undo, redo })
    history.undo()
    history.redo()
    expect(undo).toHaveBeenCalledOnce()
    expect(redo).toHaveBeenCalledOnce()
    await Promise.resolve()
  })
})
