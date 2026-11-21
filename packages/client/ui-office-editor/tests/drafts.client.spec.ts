// @vitest-environment jsdom
/**
 * Recovery drafts: the idle-debounced scheduler and the hook's post-read
 * recovery offer. The Remote is a fake face, so no Host service and no Univer
 * runtime enters this suite.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { Mock } from 'vitest'
import { act, cleanup, renderHook, waitFor } from '@testing-library/react'
import { RemoteError } from '@forwork/typert-protocol'
import type { RemoteResult } from '@forwork/typert-protocol'
import type { OfficeDocumentDraft } from '@forwork/office-exchange/types'
import { createDraftScheduler, draftRecoveryDecision } from '../src/client/drafts.ts'
import type { OfficeDraftSchedulerInput } from '../src/client/drafts.ts'
import { useOfficeDocument } from '../src/client/useOfficeDocument.ts'
import type { OfficeDocumentRemote, UseOfficeDocumentInput } from '../src/client/useOfficeDocument.ts'
import {
  DOCUMENT_ID,
  PATH,
  SESSION_ID,
  SIGNAL,
  draftAt,
  edited,
  makeDraftRemote,
  readResult,
  snapshot,
  success,
} from './draft-fixtures.client.ts'

function schedulerHarness(overrides: Partial<OfficeDraftSchedulerInput> = {}): {
  scheduler: ReturnType<typeof createDraftScheduler>
  write: Mock<OfficeDraftSchedulerInput['write']>
  onError: Mock<OfficeDraftSchedulerInput['onError']>
  controller: AbortController
} {
  const write = vi.fn<OfficeDraftSchedulerInput['write']>(
    async () => success({ savedAt: '2026-11-14T00:00:00.000Z' }),
  )
  const onError = vi.fn<OfficeDraftSchedulerInput['onError']>()
  const controller = new AbortController()
  const scheduler = createDraftScheduler({
    write,
    onError,
    signal: controller.signal,
    idleMs: 20,
    ...overrides,
  })
  return { scheduler, write, onError, controller }
}

function input(remote: OfficeDocumentRemote, overrides: Partial<UseOfficeDocumentInput> = {}): UseOfficeDocumentInput {
  return {
    remote,
    sessionId: SESSION_ID,
    path: PATH,
    documentId: DOCUMENT_ID,
    signal: SIGNAL,
    observedVersion: 'v1',
    workspaceRoot: '/work',
    openConverted: vi.fn(),
    dirty: false,
    ...overrides,
  }
}

function renderDocument(remote: OfficeDocumentRemote, overrides: Partial<UseOfficeDocumentInput> = {}) {
  return renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
    initialProps: input(remote, overrides),
  })
}

afterEach(() => {
  cleanup()
  vi.useRealTimers()
})

describe('draft scheduler', () => {
  it('decides no recovery when the document has no stored draft', () => {
    expect(draftRecoveryDecision(null, 'v1')).toBe('none')
  })

  it('writes the first edit immediately and debounces later edits', async () => {
    const { scheduler, write } = schedulerHarness()
    scheduler.schedule(snapshot, 'v1')
    expect(write).toHaveBeenCalledOnce()
    expect(write.mock.calls[0]?.[0]).toBe(snapshot)
    expect(write.mock.calls[0]?.[1]).toBe('v1')
    expect(scheduler.isPending()).toBe(false)

    vi.useFakeTimers()
    scheduler.schedule(edited, 'v1')
    await vi.advanceTimersByTimeAsync(19)
    expect(write).toHaveBeenCalledOnce()
    expect(scheduler.isPending()).toBe(true)
    await vi.advanceTimersByTimeAsync(1)
    expect(write).toHaveBeenCalledTimes(2)
    expect(write.mock.calls[1]?.[0]).toBe(edited)
    expect(scheduler.isPending()).toBe(false)
  })

  it('flushes a pending candidate without waiting for the idle period', async () => {
    vi.useFakeTimers()
    const { scheduler, write } = schedulerHarness()
    scheduler.schedule(snapshot, 'v1')
    scheduler.schedule(edited, 'v1')
    await scheduler.flush()
    expect(write).toHaveBeenCalledTimes(2)
    expect(write.mock.calls[1]?.[0]).toBe(edited)
    await vi.advanceTimersByTimeAsync(20)
    expect(write).toHaveBeenCalledTimes(2)
  })

  it('waits for an in-flight write when flushing without a candidate', async () => {
    let settle: () => void = () => {}
    const write = vi.fn<OfficeDraftSchedulerInput['write']>(() => new Promise<RemoteResult<{ savedAt: string }>>(
      (done) => { settle = () => { done(success({ savedAt: 't' })) } },
    ))
    const { scheduler } = schedulerHarness({ write })
    scheduler.schedule(edited, 'v1')
    expect(scheduler.isPending()).toBe(false)
    let settled = false
    const flushed = scheduler.flush().then(() => { settled = true })
    await Promise.resolve()
    await Promise.resolve()
    expect(settled).toBe(false)
    settle()
    await flushed
    expect(settled).toBe(true)
  })

  it('flushes with no candidate as a no-op', async () => {
    const { scheduler, write } = schedulerHarness()
    await scheduler.flush()
    expect(write).not.toHaveBeenCalled()
  })

  it('cancels a pending candidate', async () => {
    vi.useFakeTimers()
    const { scheduler, write } = schedulerHarness()
    scheduler.schedule(snapshot, 'v1')
    scheduler.schedule(edited, 'v1')
    await scheduler.cancel()
    await vi.advanceTimersByTimeAsync(20)
    expect(write).toHaveBeenCalledOnce()
    expect(scheduler.isPending()).toBe(false)
  })

  it('writes the first edit after a cancel immediately again', async () => {
    const { scheduler, write } = schedulerHarness()
    scheduler.schedule(snapshot, 'v1')
    await scheduler.cancel()
    scheduler.schedule(edited, 'v1')
    expect(write).toHaveBeenCalledTimes(2)
    expect(write.mock.calls[1]?.[0]).toBe(edited)
  })

  it('waits for an in-flight write before cancelling', async () => {
    let settle: () => void = () => {}
    const write = vi.fn<OfficeDraftSchedulerInput['write']>(() => new Promise<RemoteResult<{ savedAt: string }>>(
      (done) => { settle = () => { done(success({ savedAt: 't' })) } },
    ))
    const { scheduler } = schedulerHarness({ write })
    scheduler.schedule(edited, 'v1')
    expect(write).toHaveBeenCalledOnce()
    let settled = false
    const cancelled = scheduler.cancel().then(() => { settled = true })
    await Promise.resolve()
    await Promise.resolve()
    expect(settled).toBe(false)
    settle()
    await cancelled
    expect(settled).toBe(true)
  })

  it('forwards a refused write', async () => {
    const write = vi.fn<OfficeDraftSchedulerInput['write']>(
      async () => ({
        ok: false,
        error: new RemoteError('office-document/out-of-scope', 'outside', { path: PATH }),
      }),
    )
    const { scheduler, onError } = schedulerHarness({ write })
    scheduler.schedule(edited, 'v1')
    await scheduler.flush()
    expect(onError).toHaveBeenCalledOnce()
    expect(onError.mock.calls[0]?.[0].code).toBe('office-document/out-of-scope')
  })

  it('forwards a rejected write as an internal failure', async () => {
    const write = vi.fn<OfficeDraftSchedulerInput['write']>(
      async () => { throw new Error('offline') },
    )
    const { scheduler, onError } = schedulerHarness({ write })
    scheduler.schedule(edited, 'v1')
    await scheduler.flush()
    expect(onError).toHaveBeenCalledOnce()
    expect(onError.mock.calls[0]?.[0].code).toBe('gateway/internal')
    expect(onError.mock.calls[0]?.[0].message).toBe('offline')
  })

  it('forwards a non-Error rejection', async () => {
    const write = vi.fn<OfficeDraftSchedulerInput['write']>(
      async () => { throw 'offline' },
    )
    const { scheduler, onError } = schedulerHarness({ write })
    scheduler.schedule(edited, 'v1')
    await scheduler.flush()
    expect(onError.mock.calls[0]?.[0].message).toBe('offline')
  })

  it('forgets a settled write only when no later write replaced it', async () => {
    vi.useFakeTimers()
    const pendings: Array<(value: RemoteResult<{ savedAt: string }>) => void> = []
    const write = vi.fn<OfficeDraftSchedulerInput['write']>(() => new Promise((resolve) => {
      pendings.push(resolve)
    }))
    const { scheduler } = schedulerHarness({ write })
    scheduler.schedule(snapshot, 'v1')
    await vi.advanceTimersByTimeAsync(20)
    scheduler.schedule(edited, 'v1')
    await vi.advanceTimersByTimeAsync(20)
    expect(write).toHaveBeenCalledTimes(2)
    pendings[0]?.(success({ savedAt: 'first' }))
    await vi.advanceTimersByTimeAsync(0)
    pendings[1]?.(success({ savedAt: 'second' }))
    await vi.advanceTimersByTimeAsync(0)
  })

  it('drops a candidate scheduled after the lifetime ended', async () => {
    vi.useFakeTimers()
    const controller = new AbortController()
    controller.abort()
    const { scheduler, write } = schedulerHarness({ signal: controller.signal })
    scheduler.schedule(edited, 'v1')
    expect(scheduler.isPending()).toBe(false)
    await vi.advanceTimersByTimeAsync(20)
    expect(write).not.toHaveBeenCalled()
  })

  it('drops a candidate whose lifetime ends before its write', async () => {
    vi.useFakeTimers()
    const { scheduler, write, controller } = schedulerHarness()
    scheduler.schedule(snapshot, 'v1')
    scheduler.schedule(edited, 'v1')
    controller.abort()
    await vi.advanceTimersByTimeAsync(20)
    expect(write).toHaveBeenCalledOnce()
    expect(write.mock.calls[0]?.[0]).toBe(snapshot)
  })
})

describe('useOfficeDocument recovery', () => {
  it('offers a draft captured at the file version and applies it on request', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockResolvedValueOnce(success(draftAt('v1')))
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.recovery).toBeDefined() })
    expect(result.current.recovery?.decision).toBe('offer')
    expect(remote.readDraft.mock.calls[0]?.[1]).toBe(DOCUMENT_ID)
    expect(result.current.snapshot).toBe(snapshot)
    act(() => { result.current.acceptRecovery() })
    expect(result.current.snapshot).toBe(edited)
    expect(result.current.recovered).toBe(true)
    expect(result.current.recovery).toBeUndefined()
  })

  it('labels a draft from an earlier file version', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockResolvedValueOnce(success(draftAt('v0')))
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.recovery).toBeDefined() })
    expect(result.current.recovery?.decision).toBe('stale')
  })

  it('discards an offered draft and removes it from the store', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockResolvedValueOnce(success(draftAt('v1')))
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.recovery).toBeDefined() })
    act(() => { result.current.discardRecovery() })
    await waitFor(() => { expect(remote.clearDraft).toHaveBeenCalledOnce() })
    expect(result.current.recovery).toBeUndefined()
  })

  it('ignores a recovery request when no draft is offered', async () => {
    const remote = makeDraftRemote()
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    act(() => { result.current.acceptRecovery() })
    expect(result.current.snapshot).toBe(snapshot)
    expect(result.current.recovered).toBe(false)
  })

  it('offers no recovery when the store has no draft', async () => {
    const remote = makeDraftRemote()
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    expect(remote.readDraft).toHaveBeenCalledOnce()
    expect(result.current.recovery).toBeUndefined()
  })

  it('reports a refused draft read', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockResolvedValueOnce({
      ok: false,
      error: new RemoteError('office-document/out-of-scope', 'outside', { path: PATH }),
    })
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.draftError?.code).toBe('out-of-scope') })
  })

  it('reports a rejected draft read', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockRejectedValueOnce(new Error('offline'))
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.draftError).toEqual({ code: 'unknown', message: 'offline' }) })
  })

  it('offers a draft for a file whose read reported preserved features', async () => {
    const remote = makeDraftRemote()
    remote.read.mockResolvedValueOnce(success({ ...readResult(), unsupported: [{ kind: 'macro' }] }))
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    await waitFor(() => { expect(remote.readDraft).toHaveBeenCalledOnce() })
  })

  it('checks the draft store once per opened document', async () => {
    const remote = makeDraftRemote()
    remote.read.mockResolvedValue(success(readResult()))
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    await waitFor(() => { expect(remote.readDraft).toHaveBeenCalledOnce() })
    act(() => { result.current.reload() })
    await waitFor(() => { expect(remote.read).toHaveBeenCalledTimes(2) })
    expect(remote.readDraft).toHaveBeenCalledOnce()
  })

  it('ignores a draft that settles after the body unmounts', async () => {
    const remote = makeDraftRemote()
    let resolveDraft: (value: RemoteResult<OfficeDocumentDraft | null>) => void = () => {}
    remote.readDraft.mockImplementationOnce(() => new Promise((done) => { resolveDraft = done }))
    const { unmount } = renderDocument(remote.face)
    await waitFor(() => { expect(remote.readDraft).toHaveBeenCalledOnce() })
    unmount()
    resolveDraft(success(draftAt('v1')))
    await Promise.resolve()
    await Promise.resolve()
  })

  it('ignores a draft rejection that arrives after the body unmounts', async () => {
    const remote = makeDraftRemote()
    let rejectDraft: (reason: unknown) => void = () => {}
    remote.readDraft.mockImplementationOnce(() => new Promise((_done, fail) => { rejectDraft = fail }))
    const { unmount } = renderDocument(remote.face)
    await waitFor(() => { expect(remote.readDraft).toHaveBeenCalledOnce() })
    unmount()
    rejectDraft(new Error('late'))
    await Promise.resolve()
    await Promise.resolve()
  })
})

describe('useOfficeDocument draft capture', () => {
  it('writes the first scheduled snapshot immediately at the read version', async () => {
    const remote = makeDraftRemote()
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    act(() => { result.current.scheduleDraft(edited) })
    expect(remote.writeDraft).toHaveBeenCalledOnce()
    expect(remote.writeDraft.mock.calls[0]?.[1]).toBe(DOCUMENT_ID)
    expect(remote.writeDraft.mock.calls[0]?.[2]).toBe(PATH)
    expect(remote.writeDraft.mock.calls[0]?.[3]).toBe('v1')
    expect(remote.writeDraft.mock.calls[0]?.[4]).toBe(edited)
  })

  it('debounces a later edit until the idle period', async () => {
    vi.useFakeTimers()
    const remote = makeDraftRemote()
    const { result } = renderDocument(remote.face)
    await act(async () => { await vi.advanceTimersByTimeAsync(0) })
    act(() => { result.current.scheduleDraft(snapshot) })
    act(() => { result.current.scheduleDraft(edited) })
    expect(remote.writeDraft).toHaveBeenCalledOnce()
    await act(async () => { await vi.advanceTimersByTimeAsync(1999) })
    expect(remote.writeDraft).toHaveBeenCalledOnce()
    await act(async () => { await vi.advanceTimersByTimeAsync(1) })
    expect(remote.writeDraft).toHaveBeenCalledTimes(2)
    expect(remote.writeDraft.mock.calls[1]?.[4]).toBe(edited)
  })

  it('flushes the pending draft on pagehide and beforeunload', async () => {
    const remote = makeDraftRemote()
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    act(() => { result.current.scheduleDraft(snapshot) })
    act(() => { result.current.scheduleDraft(edited) })
    expect(remote.writeDraft).toHaveBeenCalledOnce()

    await act(async () => { window.dispatchEvent(new Event('pagehide')) })
    expect(remote.writeDraft).toHaveBeenCalledTimes(2)
    expect(remote.writeDraft.mock.calls[1]?.[4]).toBe(edited)

    act(() => { result.current.scheduleDraft(snapshot) })
    await act(async () => { window.dispatchEvent(new Event('beforeunload')) })
    expect(remote.writeDraft).toHaveBeenCalledTimes(3)
    expect(remote.writeDraft.mock.calls[2]?.[4]).toBe(snapshot)
  })

  it('does not capture an edit before the read succeeds', async () => {
    vi.useFakeTimers()
    const remote = makeDraftRemote()
    remote.read.mockImplementationOnce(() => new Promise(() => {}))
    const { result } = renderDocument(remote.face)
    act(() => { result.current.scheduleDraft(edited) })
    await act(async () => { await vi.advanceTimersByTimeAsync(2000) })
    expect(remote.writeDraft).not.toHaveBeenCalled()
  })

  it('reports a refused draft write into the failure state', async () => {
    vi.useFakeTimers()
    const remote = makeDraftRemote()
    remote.writeDraft.mockResolvedValueOnce({
      ok: false,
      error: new RemoteError('office-document/bound-exceeded', 'too large', { path: PATH }),
    })
    const { result } = renderDocument(remote.face)
    await act(async () => { await vi.advanceTimersByTimeAsync(0) })
    act(() => { result.current.scheduleDraft(edited) })
    await act(async () => { await vi.advanceTimersByTimeAsync(2000) })
    expect(result.current.draftError?.code).toBe('bound-exceeded')
  })

  it('captures an edit for a file whose read reported preserved features', async () => {
    vi.useFakeTimers()
    const remote = makeDraftRemote()
    remote.read.mockResolvedValueOnce(success({ ...readResult(), unsupported: [{ kind: 'macro' }] }))
    const { result } = renderDocument(remote.face)
    await act(async () => { await vi.advanceTimersByTimeAsync(0) })
    expect(result.current.status).toBe('ready')
    act(() => { result.current.scheduleDraft(edited) })
    await act(async () => { await vi.advanceTimersByTimeAsync(2000) })
    expect(remote.writeDraft).toHaveBeenCalledOnce()
  })

  it('clears the stored draft after a successful save', async () => {
    const remote = makeDraftRemote()
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    let outcome: string | undefined
    await act(async () => { outcome = await result.current.save(edited) })
    expect(outcome).toBe('saved')
    expect(remote.clearDraft).toHaveBeenCalledOnce()
  })

  it('keeps the saved outcome when the draft store cannot be cleared', async () => {
    const remote = makeDraftRemote()
    remote.clearDraft.mockRejectedValueOnce(new Error('offline'))
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    let outcome: string | undefined
    await act(async () => { outcome = await result.current.save(edited) })
    expect(outcome).toBe('saved')
    expect(result.current.draftError).toEqual({ code: 'unknown', message: 'offline' })
  })

  it('clears the draft after the tab lifetime ended', async () => {
    const remote = makeDraftRemote()
    const controller = new AbortController()
    const { result } = renderDocument(remote.face, { signal: controller.signal })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    controller.abort()
    act(() => { result.current.discardDraft() })
    await waitFor(() => { expect(remote.clearDraft).toHaveBeenCalledOnce() })
    expect(remote.clearDraft.mock.calls[0]?.[2]?.aborted).toBe(false)
  })

  it('reports a refused draft removal', async () => {
    const remote = makeDraftRemote()
    remote.clearDraft.mockResolvedValueOnce({
      ok: false,
      error: new RemoteError('office-document/out-of-scope', 'outside', { path: PATH }),
    })
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    act(() => { result.current.discardDraft() })
    await waitFor(() => { expect(result.current.draftError?.code).toBe('out-of-scope') })
  })

  it('discards the draft on request without saving', async () => {
    const remote = makeDraftRemote()
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    act(() => { result.current.discardDraft() })
    await waitFor(() => { expect(remote.clearDraft).toHaveBeenCalledOnce() })
    expect(remote.save).not.toHaveBeenCalled()
  })

  it('hides the draft failure message on request', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockResolvedValueOnce({
      ok: false,
      error: new RemoteError('office-document/out-of-scope', 'outside', { path: PATH }),
    })
    const { result } = renderDocument(remote.face)
    await waitFor(() => { expect(result.current.draftError).toBeDefined() })
    act(() => { result.current.dismissDraftError() })
    expect(result.current.draftError).toBeUndefined()
  })
})
