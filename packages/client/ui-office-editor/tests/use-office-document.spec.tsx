// @vitest-environment jsdom
/**
 * `useOfficeDocument`: reads, version-guarded saves, refusal, and external
 * changes. The Remote is a fake face; no Host service and no Univer runtime
 * enters this suite.
 */
import { describe, expect, it, vi } from 'vitest'
import type { Mock } from 'vitest'
import { act, renderHook, waitFor } from '@testing-library/react'
import { brandString } from '@forwork/brand'
import type { SessionId } from '@forwork/session/types'
import { RemoteError } from '@forwork/typert-protocol'
import type { RemoteResult } from '@forwork/typert-protocol'
import type {
  DocumentSaveAttribution,
  OfficeDocumentReadResult,
  OfficeDocumentSaveResult,
  OfficeDocumentSnapshot,
  OfficeFormat,
} from '@forwork/office-exchange/types'
import { mapOfficeFailure, useOfficeDocument } from '../src/client/useOfficeDocument.ts'
import type { OfficeDocumentRemote, UseOfficeDocumentInput } from '../src/client/useOfficeDocument.ts'

const SESSION_ID = brandString<SessionId>('s-1')
const PATH = '/work/budget.xlsx'
const SIGNAL = new AbortController().signal
const snapshot: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book' } }
const edited: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book', name: 'edited' } }

function readResult(version: string, overrides: Partial<OfficeDocumentReadResult> = {}): OfficeDocumentReadResult {
  return { format: 'xlsx', snapshot, unsupported: [], version, ...overrides }
}

function success<T>(value: T): RemoteResult<T> {
  return { ok: true, value }
}

type ReadSignature = (
  sessionId: SessionId,
  path: string,
  signal?: AbortSignal,
) => Promise<RemoteResult<OfficeDocumentReadResult>>

type ConvertSignature = (
  sessionId: SessionId,
  path: string,
  signal?: AbortSignal,
) => Promise<RemoteResult<string>>

type SaveSignature = (
  sessionId: SessionId,
  path: string,
  snapshot: OfficeDocumentSnapshot,
  format: OfficeFormat,
  expectedVersion: string,
  attribution: DocumentSaveAttribution,
  signal?: AbortSignal,
) => Promise<RemoteResult<OfficeDocumentSaveResult>>

type WriteDraftSignature = OfficeDocumentRemote['writeDraft']
type ReadDraftSignature = OfficeDocumentRemote['readDraft']
type ClearDraftSignature = OfficeDocumentRemote['clearDraft']

function makeRemote(): {
  read: Mock<ReadSignature>
  convert: Mock<ConvertSignature>
  save: Mock<SaveSignature>
  writeDraft: Mock<WriteDraftSignature>
  readDraft: Mock<ReadDraftSignature>
  clearDraft: Mock<ClearDraftSignature>
  face: OfficeDocumentRemote
} {
  const read = vi.fn<ReadSignature>().mockResolvedValue(success(readResult('v1')))
  const convert = vi.fn<ConvertSignature>().mockResolvedValue(success('/work/report.docx'))
  const save = vi.fn<SaveSignature>().mockResolvedValue(success({ version: 'v2' }))
  const writeDraft = vi.fn<WriteDraftSignature>().mockResolvedValue(success({ savedAt: 't' }))
  const readDraft = vi.fn<ReadDraftSignature>().mockResolvedValue(success(null))
  const clearDraft = vi.fn<ClearDraftSignature>().mockResolvedValue(success(undefined))
  return {
    read,
    convert,
    save,
    writeDraft,
    readDraft,
    clearDraft,
    face: { read, convertLegacy: convert, save, writeDraft, readDraft, clearDraft },
  }
}

function input(remote: OfficeDocumentRemote, overrides: Partial<UseOfficeDocumentInput> = {}): UseOfficeDocumentInput {
  return {
    remote,
    sessionId: SESSION_ID,
    path: PATH,
    documentId: `${SESSION_ID}:${PATH}`,
    signal: SIGNAL,
    observedVersion: 'v1',
    workspaceRoot: '/work',
    openConverted: vi.fn(),
    dirty: false,
    ...overrides,
  }
}

describe('useOfficeDocument read', () => {
  it('reads through the Remote and exposes the snapshot and version', async () => {
    const remote = makeRemote()
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    expect(result.current.snapshot).toBe(snapshot)
    expect(result.current.version).toBe('v1')
    expect(result.current.canSave).toBe(true)
    expect(remote.read.mock.calls[0]?.[0]).toBe(SESSION_ID)
    expect(remote.read.mock.calls[0]?.[1]).toBe(PATH)
  })

  it('maps a read failure onto its visible class', async () => {
    const remote = makeRemote()
    remote.read.mockResolvedValueOnce({
      ok: false,
      error: new RemoteError('office-document/out-of-scope', 'outside', { path: PATH }),
    })
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('error') })
    expect(result.current.error?.code).toBe('out-of-scope')
    expect(result.current.canSave).toBe(false)
  })

  it('reports a rejected read as an unknown failure', async () => {
    const remote = makeRemote()
    remote.read.mockRejectedValueOnce(new Error('offline'))
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('error') })
    expect(result.current.error).toEqual({ code: 'unknown', message: 'offline' })
  })

  it('stringifies a non-Error read rejection', async () => {
    const remote = makeRemote()
    remote.read.mockRejectedValueOnce('offline')
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('error') })
    expect(result.current.error).toEqual({ code: 'unknown', message: 'offline' })
  })

  it('does not read after the tab lifetime ended', () => {
    const remote = makeRemote()
    const controller = new AbortController()
    controller.abort()
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face, { signal: controller.signal }),
    })
    expect(remote.read).not.toHaveBeenCalled()
    expect(result.current.status).toBe('loading')
  })

  it('ignores a read that settles after unmount', async () => {
    const remote = makeRemote()
    let resolve: (value: RemoteResult<OfficeDocumentReadResult>) => void = () => {}
    remote.read.mockImplementationOnce(() => new Promise((done) => { resolve = done }))
    const { unmount } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    unmount()
    resolve(success(readResult('v1')))
    await Promise.resolve()
  })

  it('ignores a read rejection that arrives after unmount', async () => {
    const remote = makeRemote()
    let reject: (reason: unknown) => void = () => {}
    remote.read.mockImplementationOnce(() => new Promise((_done, fail) => { reject = fail }))
    const { unmount } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    unmount()
    reject(new Error('late'))
    await Promise.resolve()
    await Promise.resolve()
  })
})

describe('useOfficeDocument save', () => {
  it('saves the live snapshot at the read version and advances the version', async () => {
    const remote = makeRemote()
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    let outcome: string | undefined
    await act(async () => { outcome = await result.current.save(edited) })
    expect(outcome).toBe('saved')
    expect(remote.save.mock.calls[0]?.[2]).toBe(edited)
    expect(remote.save.mock.calls[0]?.[3]).toBe('xlsx')
    expect(remote.save.mock.calls[0]?.[4]).toBe('v1')
    expect(result.current.version).toBe('v2')
  })

  it('saves a document snapshot with its own format and the read snapshot when no edit is passed', async () => {
    const remote = makeRemote()
    const documentSnapshot: OfficeDocumentSnapshot = { kind: 'document', data: { id: 'doc' } }
    remote.read.mockResolvedValueOnce(success(readResult('v1', { format: 'docx', snapshot: documentSnapshot })))
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    let outcome: string | undefined
    await act(async () => { outcome = await result.current.save() })
    expect(outcome).toBe('saved')
    expect(remote.save.mock.calls[0]?.[2]).toBe(documentSnapshot)
    expect(remote.save.mock.calls[0]?.[3]).toBe('docx')
  })

  it('refuses save when the read reported an unsupported feature', async () => {
    const remote = makeRemote()
    remote.read.mockResolvedValueOnce(success(readResult('v1', { unsupported: [{ kind: 'macro' }] })))
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('unsupported') })
    expect(result.current.canSave).toBe(false)
    let outcome: string | undefined
    await act(async () => { outcome = await result.current.save(edited) })
    expect(outcome).toBe('refused')
    expect(remote.save).not.toHaveBeenCalled()
  })

  it('disables save for a document that declares explicit run formatting', async () => {
    const remote = makeRemote()
    const documentSnapshot: OfficeDocumentSnapshot = { kind: 'document', data: { id: 'doc' } }
    remote.read.mockResolvedValueOnce(success(readResult('v1', {
      format: 'docx',
      snapshot: documentSnapshot,
      unsupported: [{ kind: 'run-formatting' }],
    })))
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('unsupported') })
    expect(result.current.unsupported).toEqual([{ kind: 'run-formatting' }])
    expect(result.current.canSave).toBe(false)
  })

  it('refuses save before a snapshot has been read', async () => {
    const remote = makeRemote()
    remote.read.mockImplementationOnce(() => new Promise(() => {}))
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    let outcome: string | undefined
    await act(async () => { outcome = await result.current.save(edited) })
    expect(outcome).toBe('refused')
    expect(remote.save).not.toHaveBeenCalled()
  })

  it('maps a stale version onto the conflict state', async () => {
    const remote = makeRemote()
    remote.save.mockResolvedValueOnce({
      ok: false,
      error: new RemoteError('office-document/conflict', 'changed', { path: PATH }),
    })
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    let outcome: string | undefined
    await act(async () => { outcome = await result.current.save(edited) })
    expect(outcome).toBe('conflict')
    expect(result.current.status).toBe('conflict')
    expect(result.current.error?.code).toBe('conflict')
    expect(result.current.version).toBe('v1')
    expect(result.current.canSave).toBe(false)
  })

  it('reports a non-conflict save failure without losing the snapshot', async () => {
    const remote = makeRemote()
    remote.save.mockResolvedValueOnce({
      ok: false,
      error: new RemoteError('office-document/conversion-failed', 'io', { path: PATH }),
    })
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    let outcome: string | undefined
    await act(async () => { outcome = await result.current.save(edited) })
    expect(outcome).toBe('error')
    expect(result.current.error?.code).toBe('conversion-failed')
    expect(result.current.snapshot).toBe(snapshot)
  })

  it('refuses a second save while one is in flight', async () => {
    const remote = makeRemote()
    let resolve: (value: RemoteResult<OfficeDocumentSaveResult>) => void = () => {}
    remote.save.mockImplementationOnce(() => new Promise((done) => { resolve = done }))
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    let first: string | undefined
    let second: string | undefined
    await act(async () => {
      const pending = result.current.save(edited)
      second = await result.current.save(edited)
      resolve(success({ version: 'v2' }))
      first = await pending
    })
    expect(second).toBe('refused')
    expect(first).toBe('saved')
    expect(remote.save).toHaveBeenCalledOnce()
  })

  it('refuses a save that settles after the tab lifetime ended', async () => {
    const remote = makeRemote()
    const controller = new AbortController()
    let resolve: (value: RemoteResult<OfficeDocumentSaveResult>) => void = () => {}
    remote.save.mockImplementationOnce(() => new Promise((done) => { resolve = done }))
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face, { signal: controller.signal }),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    let outcome: string | undefined
    await act(async () => {
      const pending = result.current.save(edited)
      controller.abort()
      resolve(success({ version: 'v2' }))
      outcome = await pending
    })
    expect(outcome).toBe('refused')
  })
})

describe('useOfficeDocument external change', () => {
  it('prompts instead of reloading a dirty unit and dismisses on request', async () => {
    const remote = makeRemote()
    const { result, rerender } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    rerender(input(remote.face, { observedVersion: 'v2', dirty: true }))
    await waitFor(() => { expect(result.current.externalChange).toBe(true) })
    expect(remote.read).toHaveBeenCalledOnce()
    act(() => { result.current.dismissExternalChange() })
    expect(result.current.externalChange).toBe(false)
  })

  it('reloads a clean unit when the file version changes', async () => {
    const remote = makeRemote()
    remote.read.mockResolvedValueOnce(success(readResult('v1')))
    remote.read.mockResolvedValueOnce(success(readResult('v2')))
    const { result, rerender } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    rerender(input(remote.face, { observedVersion: 'v2' }))
    await waitFor(() => { expect(result.current.version).toBe('v2') })
    expect(remote.read).toHaveBeenCalledTimes(2)
  })

  it('reloads on demand, ignoring an ignored external-change frame', async () => {
    const remote = makeRemote()
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face, { observedVersion: undefined }),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    act(() => { result.current.reload() })
    await waitFor(() => { expect(remote.read).toHaveBeenCalledTimes(2) })
  })
})

describe('useOfficeDocument agent command save', () => {
  it('adopts the published version in place when the Host save matches the applied command', async () => {
    const remote = makeRemote()
    remote.read.mockResolvedValueOnce(success(readResult('v1')))
    const { result, rerender } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face, { contentVersion: 0 }),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    act(() => { result.current.markApplied() })
    expect(result.current.contentVersion).toBe(1)

    rerender(input(remote.face, { contentVersion: 0, observedVersion: 'v2' }))
    expect(result.current.version).toBe('v1')
    expect(remote.read).toHaveBeenCalledOnce()

    rerender(input(remote.face, { contentVersion: 1, observedVersion: 'v2' }))
    expect(result.current.version).toBe('v2')
    expect(result.current.snapshot).toBe(snapshot)
    expect(result.current.externalChange).toBe(false)
    expect(remote.read).toHaveBeenCalledOnce()
  })

  it('prompts for a file change while the unit gained user edits after the command', async () => {
    const remote = makeRemote()
    const { result, rerender } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face, { contentVersion: 0 }),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    act(() => { result.current.markApplied() })

    rerender(input(remote.face, { contentVersion: 1, observedVersion: 'v2', dirty: true }))
    await waitFor(() => { expect(result.current.externalChange).toBe(true) })
    expect(remote.read).toHaveBeenCalledOnce()
  })

  it('holds a file change until the projection reports the applied command version', async () => {
    const remote = makeRemote()
    const { result, rerender } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    act(() => { result.current.markApplied() })

    rerender(input(remote.face, { observedVersion: 'v2' }))
    expect(result.current.version).toBe('v1')
    expect(result.current.externalChange).toBe(false)
    expect(remote.read).toHaveBeenCalledOnce()

    rerender(input(remote.face, { contentVersion: 1, observedVersion: 'v2' }))
    expect(result.current.version).toBe('v2')
    expect(remote.read).toHaveBeenCalledOnce()
  })

  it('reloads when another save advanced the projection past the applied command version', async () => {
    const remote = makeRemote()
    remote.read.mockResolvedValueOnce(success(readResult('v1')))
    remote.read.mockResolvedValueOnce(success(readResult('v2')))
    const { result, rerender } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face, { contentVersion: 0 }),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    act(() => { result.current.markApplied() })

    rerender(input(remote.face, { contentVersion: 2, observedVersion: 'v2' }))
    await waitFor(() => { expect(result.current.version).toBe('v2') })
    expect(remote.read).toHaveBeenCalledTimes(2)
  })

  it('drops a staged command save when a fresh read starts', async () => {
    const remote = makeRemote()
    remote.read.mockResolvedValueOnce(success(readResult('v1')))
    remote.read.mockResolvedValueOnce(success(readResult('v2')))
    remote.read.mockResolvedValueOnce(success(readResult('v3')))
    const { result, rerender } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face, { contentVersion: 0 }),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    act(() => { result.current.markApplied() })
    act(() => { result.current.reload() })
    await waitFor(() => { expect(result.current.version).toBe('v2') })

    rerender(input(remote.face, { contentVersion: 1, observedVersion: 'v3' }))
    await waitFor(() => { expect(result.current.version).toBe('v3') })
    expect(remote.read).toHaveBeenCalledTimes(3)
  })
})

describe('mapOfficeFailure', () => {
  it('maps every declared Remote code', () => {
    expect(mapOfficeFailure(new RemoteError('office-document/conflict', 'm', { path: PATH })).code).toBe('conflict')
    expect(mapOfficeFailure(new RemoteError('office-document/unsupported-format', 'm', { path: PATH })).code)
      .toBe('unsupported-format')
    expect(mapOfficeFailure(new RemoteError('office-document/unsupported-feature', 'm', { path: PATH })).code)
      .toBe('unsupported-feature')
    expect(mapOfficeFailure(new RemoteError('office-document/out-of-scope', 'm', { path: PATH })).code).toBe('out-of-scope')
    expect(mapOfficeFailure(new RemoteError('office-document/bound-exceeded', 'm', { path: PATH })).code).toBe('bound-exceeded')
    expect(mapOfficeFailure(new RemoteError('office-document/conversion-failed', 'm', { path: PATH })).code)
      .toBe('conversion-failed')
    expect(mapOfficeFailure(new RemoteError('gateway/cancelled', 'm', {})).code).toBe('cancelled')
    expect(mapOfficeFailure(new RemoteError('gateway/internal', 'm', {})).code).toBe('unknown')
  })
})

describe('useOfficeDocument Host content version', () => {
  it('adopts a Host save before the file reload instead of treating it as external', async () => {
    const remote = makeRemote()
    remote.read.mockResolvedValueOnce(success(readResult('v1')))
    remote.read.mockResolvedValueOnce(success(readResult('v2')))
    const { result, rerender } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face, { contentVersion: 0 }),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })

    rerender(input(remote.face, { contentVersion: 1 }))
    expect(result.current.contentVersion).toBe(1)
    expect(remote.read).toHaveBeenCalledOnce()

    rerender(input(remote.face, { contentVersion: 1, observedVersion: 'v2' }))
    await waitFor(() => { expect(result.current.version).toBe('v2') })
    expect(result.current.externalChange).toBe(false)
    expect(remote.read).toHaveBeenCalledTimes(2)
  })

  it('adopts a Host save that arrives after the file version changed', async () => {
    const remote = makeRemote()
    remote.read.mockResolvedValueOnce(success(readResult('v1')))
    remote.read.mockResolvedValueOnce(success(readResult('v2')))
    const { result, rerender } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face, { contentVersion: 0 }),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })

    rerender(input(remote.face, { contentVersion: 0, observedVersion: 'v2' }))
    await waitFor(() => { expect(result.current.version).toBe('v2') })

    rerender(input(remote.face, { contentVersion: 1, observedVersion: 'v2' }))
    expect(result.current.contentVersion).toBe(1)
    expect(result.current.externalChange).toBe(false)
    expect(remote.read).toHaveBeenCalledTimes(2)
  })

  it('keeps one content version across Host saves, commands, and local saves', async () => {
    const remote = makeRemote()
    const { result } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face, { contentVersion: 4 }),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })
    expect(result.current.contentVersion).toBe(4)
    expect(result.current.currentContentVersion()).toBe(4)

    act(() => { result.current.markApplied() })
    expect(result.current.contentVersion).toBe(5)
    expect(result.current.currentContentVersion()).toBe(5)

    let outcome: string | undefined
    await act(async () => { outcome = await result.current.save(edited) })
    expect(outcome).toBe('saved')
    expect(result.current.contentVersion).toBe(6)
    expect(result.current.currentContentVersion()).toBe(6)
  })

  it('still prompts for an external change while dirty', async () => {
    const remote = makeRemote()
    const { result, rerender } = renderHook((props: UseOfficeDocumentInput) => useOfficeDocument(props), {
      initialProps: input(remote.face, { contentVersion: 0 }),
    })
    await waitFor(() => { expect(result.current.status).toBe('ready') })

    rerender(input(remote.face, { contentVersion: 1, observedVersion: 'v2', dirty: true }))
    await waitFor(() => { expect(result.current.externalChange).toBe(true) })
    expect(remote.read).toHaveBeenCalledOnce()
  })
})
