// @vitest-environment jsdom
/**
 * The panel supplies the command bridge with the live tab target: identity,
 * numeric version, dirty state, adapter apply, and post-command snapshot.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, waitFor } from '@testing-library/react'
import type { OfficeDocumentReadResult, OfficeDocumentSnapshot } from '@deepseek-ai/dsh-office-exchange/types'
import type { RemoteResult } from '@deepseek-ai/dsh-typert-protocol'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import type { OfficeCommandTarget } from '../src/client/command-bridge.ts'
import type { OfficeDocumentRemote } from '../src/client/useOfficeDocument.ts'
import { renderPanel, SESSION_ID } from './panel.tsx'

const PATH = 'work/budget.xlsx'
const snapshot: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book' } }

function success<T>(value: T): RemoteResult<T> {
  return { ok: true, value }
}

function remote(): OfficeDocumentRemote {
  return {
    read: vi.fn(async (): Promise<RemoteResult<OfficeDocumentReadResult>> => (
      success({ format: 'xlsx', snapshot, unsupported: [], version: 'v1' })
    )),
    convertLegacy: vi.fn(async () => success('/work/report.docx')),
    save: vi.fn(async () => success({ version: 'v2' })),
    writeDraft: vi.fn(async () => success({ savedAt: 't' })),
    readDraft: vi.fn(async () => success(null)),
    clearDraft: vi.fn(async () => success(undefined)),
  }
}

afterEach(cleanup)

describe('OfficePanel command target', () => {
  it('exposes the mounted adapter and advances the content version on an applied command', async () => {
    const dispose = vi.fn()
    let captured: OfficeCommandTarget | undefined
    const attachCommandBridge = vi.fn((_sessionId: SessionId, target: OfficeCommandTarget) => {
      captured = target
      return dispose
    })
    const view = renderPanel({
      meta: { absolutePath: `/${PATH}`, version: 'v1' },
      remote: remote(),
      attachCommandBridge,
    })
    await waitFor(() => { expect(view.fixture.adapter.mounts).toHaveLength(1) })
    await waitFor(() => { expect(attachCommandBridge).toHaveBeenCalledOnce() })
    if (captured === undefined) throw new Error('command bridge was not attached')
    const target = captured
    expect(target.documentId).toBe(`${SESSION_ID}:${PATH}`)
    expect(target.version()).toBe(0)
    expect(target.isOpen()).toBe(true)
    expect(target.isDirty()).toBe(false)
    expect(target.unsupported()).toEqual([])
    expect(target.snapshot()).toBe(snapshot)
    const operation = [{ kind: 'paragraph' as const, index: 0, text: 'Hello' }]
    expect(target.apply(operation)).toEqual({ status: 'applied' })
    expect(view.fixture.adapter.applied).toEqual([operation])
    expect(target.version()).toBe(1)
    view.unmount()
    expect(dispose).toHaveBeenCalledOnce()
    expect(target.isOpen()).toBe(false)
    expect(target.isDirty()).toBe(false)
    expect(target.snapshot()).toBeUndefined()
    expect(target.apply(operation)).toEqual({
      status: 'unsupported-feature',
      features: [{ kind: 'other', detail: 'editor is not mounted' }],
    })
    expect(target.version()).toBe(1)
  })

  it('keeps the content version when the adapter refuses an operation', async () => {
    let captured: OfficeCommandTarget | undefined
    const attachCommandBridge = vi.fn((_sessionId: SessionId, target: OfficeCommandTarget) => {
      captured = target
      return () => {}
    })
    const features = [{ kind: 'other' as const, detail: 'row' }]
    renderPanel({
      meta: { absolutePath: `/${PATH}`, version: 'v1' },
      remote: remote(),
      attachCommandBridge,
      mount: async (_container, input) => ({
        dispose: vi.fn(),
        getSnapshot: () => input.snapshot,
        apply: () => ({ status: 'unsupported-feature', features }),
        undo: vi.fn(),
        redo: vi.fn(),
        isDirty: () => false,
        clearDirty: vi.fn(),
      }),
    })
    await waitFor(() => { expect(attachCommandBridge).toHaveBeenCalledOnce() })
    if (captured === undefined) throw new Error('command bridge was not attached')
    const target = captured
    await waitFor(() => { expect(target.isOpen()).toBe(true) })
    const operation = [{ kind: 'paragraph' as const, index: 0, text: 'Hello' }]
    expect(target.apply(operation)).toEqual({ status: 'unsupported-feature', features })
    expect(target.version()).toBe(0)
  })
})
