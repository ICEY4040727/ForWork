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
  const emit = (): void => {
    for (const listener of [...listeners]) listener()
  }
  const runtime: OfficeRuntime = {
    getSnapshot: () => edited,
    apply: () => {
      emit()
      return { status: 'applied' }
    },
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
    emit,
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
    // The runtime gets its own host inside the caller's container, so its
    // teardown never removes a node another renderer owns.
    expect(fake.mounts[0]?.container.parentElement).toBe(container)
    expect(fake.mounts[0]?.container).not.toBe(container)
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
    handle.clearDirty()
    expect(handle.apply([{ kind: 'cell', sheetId: 'sheet-1', row: 0, column: 0, value: 1 }]))
      .toEqual({ status: 'applied' })
    expect(handle.isDirty()).toBe(false)
    expect(onChange).toHaveBeenCalledTimes(2)
    expect(onEdit).toHaveBeenCalledTimes(3)
  })

  it('disposes the runtime and its own host without touching the caller-owned container', async () => {
    const fake = fakeRuntime()
    const adapter = createAdapter(async (mount) => {
      fake.mounts.push(mount)
      return fake.runtime
    })
    const container = document.createElement('div')
    const owned = document.createElement('span')
    container.appendChild(owned)
    const handle = await adapter.mount(container, { kind: 'document', snapshot, onChange: vi.fn() })
    const host = fake.mounts[0]?.container
    if (host === undefined) throw new Error('the runtime mount was not recorded')
    host.appendChild(document.createElement('div'))
    handle.dispose()
    expect(fake.dispose).toHaveBeenCalledOnce()
    expect(container.contains(host)).toBe(false)
    // The caller's own child (a React-managed sibling, in the panel) survives:
    // the runtime owns only the host it was given.
    expect(container.firstElementChild).toBe(owned)
    fake.emit()
    expect(handle.isDirty()).toBe(false)
    handle.dispose()
    expect(fake.dispose).toHaveBeenCalledOnce()
  })

  it('removes its host when the runtime cannot start', async () => {
    const adapter = createAdapter(async () => { throw new Error('no canvas') })
    const container = document.createElement('div')
    await expect(adapter.mount(container, {
      kind: 'document', snapshot, onChange: vi.fn(),
    })).rejects.toThrow('no canvas')
    expect(container.childElementCount).toBe(0)
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
    const operation = [{ kind: 'paragraph' as const, index: 0, text: 'Hello' }]
    expect(handle.apply(operation)).toEqual({ status: 'applied' })
    expect(adapter.applied).toEqual([operation])
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
