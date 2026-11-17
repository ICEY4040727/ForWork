// @vitest-environment jsdom
/**
 * The panel's recovery-draft surfaces: the localized offer, the stale label,
 * discard, draft-store failures, and the discard-and-close path.
 */
import { act, cleanup, fireEvent, waitFor } from '@testing-library/react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { RemoteError } from '@forwork/typert-protocol'
import { en } from '../src/client/locales.ts'
import { TAB_ID, SESSION_ID, renderPanel } from './panel.tsx'
import { PATH, draftAt, edited, makeDraftRemote, snapshot, success } from './draft-fixtures.client.ts'

afterEach(() => {
  cleanup()
  vi.useRealTimers()
})

function fixture(remote: ReturnType<typeof makeDraftRemote>, version = 'v1') {
  return renderPanel({ meta: { absolutePath: PATH, version }, remote: remote.face })
}

describe('OfficePanel recovery drafts', () => {
  it('offers a current draft, recovers it, and marks the tab dirty', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockResolvedValueOnce(success(draftAt('v1')))
    const view = fixture(remote)
    expect(await view.findByText(en.draftFound)).toBeDefined()
    await waitFor(() => { expect(view.fixture.adapter.mounts).toHaveLength(1) })
    expect(view.fixture.adapter.mounts[0]?.snapshot).toBe(snapshot)
    fireEvent.click(view.getByText(en.recover))
    expect(view.queryByText(en.draftFound)).toBeNull()
    await waitFor(() => {
      expect(view.fixture.adapter.mounts.at(-1)?.snapshot).toBe(edited)
    })
    expect(view.fixture.props.markTabDirty).toHaveBeenCalledWith(TAB_ID, true)
  })

  it('labels a stale draft and offers to recover it anyway', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockResolvedValueOnce(success(draftAt('v0')))
    const view = fixture(remote)
    expect(await view.findByText(en.draftStale)).toBeDefined()
    fireEvent.click(view.getByText(en.recoverAnyway))
    await waitFor(() => {
      expect(view.fixture.adapter.mounts.at(-1)?.snapshot).toBe(edited)
    })
  })

  it('discards the offered draft from the panel', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockResolvedValueOnce(success(draftAt('v1')))
    const view = fixture(remote)
    expect(await view.findByText(en.draftFound)).toBeDefined()
    fireEvent.click(view.getByText(en.discardDraft))
    await waitFor(() => { expect(remote.clearDraft).toHaveBeenCalledOnce() })
    expect(view.queryByText(en.draftFound)).toBeNull()
  })

  it('reports and dismisses a draft store failure', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockResolvedValueOnce({
      ok: false,
      error: new RemoteError('office-document/out-of-scope', 'outside', { path: PATH }),
    })
    const view = fixture(remote)
    expect(await view.findByText(en.draftError)).toBeDefined()
    fireEvent.click(view.getByText(en.dismiss))
    expect(view.queryByText(en.draftError)).toBeNull()
  })

  it('writes the first edit immediately and debounces later edits', async () => {
    vi.useFakeTimers()
    const remote = makeDraftRemote()
    const view = fixture(remote)
    await act(async () => { await vi.advanceTimersByTimeAsync(0) })
    expect(view.fixture.adapter.mounts).toHaveLength(1)
    act(() => { view.fixture.adapter.setDirty(true) })
    expect(remote.writeDraft).toHaveBeenCalledOnce()
    expect(remote.writeDraft.mock.calls[0]?.[3]).toBe('v1')
    expect(remote.writeDraft.mock.calls[0]?.[4]).toBe(snapshot)

    await act(async () => { await vi.advanceTimersByTimeAsync(1000) })
    act(() => { view.fixture.adapter.setDirty(true) })
    await act(async () => { await vi.advanceTimersByTimeAsync(1999) })
    expect(remote.writeDraft).toHaveBeenCalledOnce()
    await act(async () => { await vi.advanceTimersByTimeAsync(1) })
    expect(remote.writeDraft).toHaveBeenCalledTimes(2)
    expect(remote.writeDraft.mock.calls[1]?.[4]).toBe(snapshot)
  })

  it('flushes the pending draft when the page unloads', async () => {
    const remote = makeDraftRemote()
    const view = fixture(remote)
    await waitFor(() => { expect(view.fixture.adapter.mounts).toHaveLength(1) })
    act(() => { view.fixture.adapter.setDirty(true) })
    act(() => { view.fixture.adapter.setDirty(true) })
    expect(remote.writeDraft).toHaveBeenCalledOnce()

    await act(async () => { window.dispatchEvent(new Event('pagehide')) })
    expect(remote.writeDraft).toHaveBeenCalledTimes(2)
  })

  it('flushes the pending draft on request without waiting for the idle period', async () => {
    const remote = makeDraftRemote()
    const view = fixture(remote)
    await waitFor(() => { expect(view.fixture.adapter.mounts).toHaveLength(1) })
    act(() => { view.fixture.adapter.setDirty(true) })
    act(() => { view.fixture.adapter.setDirty(true) })
    expect(remote.writeDraft).toHaveBeenCalledOnce()

    act(() => { view.fixture.changes.flushDraft(TAB_ID) })
    await waitFor(() => { expect(remote.writeDraft).toHaveBeenCalledTimes(2) })
  })

  it('recovers a flushed draft after the process restarts', async () => {
    const remote = makeDraftRemote()
    const view = fixture(remote)
    await waitFor(() => { expect(view.fixture.adapter.mounts).toHaveLength(1) })
    act(() => { view.fixture.adapter.setDirty(true) })
    const written = remote.writeDraft.mock.calls[0]
    expect(written).toBeDefined()
    view.unmount()

    const restarted = makeDraftRemote()
    restarted.readDraft.mockResolvedValueOnce(success(draftAt('v1', written?.[4] ?? snapshot)))
    const again = fixture(restarted)
    expect(await again.findByText(en.draftFound)).toBeDefined()
  })

  it('ignores a change reported before the handle is recorded', async () => {
    vi.useFakeTimers()
    const remote = makeDraftRemote()
    const view = renderPanel({
      meta: { absolutePath: PATH, version: 'v1' },
      remote: remote.face,
      mount: (container, input) => {
        input.onEdit?.()
        input.onChange()
        return Promise.resolve({
          dispose() { container.replaceChildren() },
          getSnapshot: () => input.snapshot,
          apply() { return { status: 'applied' } },
          undo() {},
          redo() {},
          isDirty: () => false,
          clearDirty() {},
        })
      },
    })
    await act(async () => { await vi.advanceTimersByTimeAsync(2000) })
    expect(view.fixture.props.markTabDirty).toHaveBeenCalledWith(TAB_ID, true)
    expect(remote.writeDraft).not.toHaveBeenCalled()
  })

  it('clears the draft when the user discards and closes', async () => {
    const remote = makeDraftRemote()
    const view = fixture(remote)
    await waitFor(() => { expect(view.fixture.adapter.mounts).toHaveLength(1) })
    act(() => {
      view.fixture.changes.markDirty(TAB_ID, true)
      view.fixture.changes.warnClose(TAB_ID)
    })
    fireEvent.click(await view.findByText(en.discardAndClose))
    await waitFor(() => { expect(remote.clearDraft).toHaveBeenCalledOnce() })
  })

  it('offers the draft an interrupted process left when the document reopens', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockResolvedValueOnce(success(draftAt('v1')))
    const view = fixture(remote)
    expect(await view.findByText(en.draftFound)).toBeDefined()
    fireEvent.click(view.getByText(en.recover))
    await waitFor(() => {
      expect(view.fixture.adapter.mounts.at(-1)?.snapshot).toBe(edited)
    })
  })

  it('warns instead of closing a dirty tab and keeps the edits', async () => {
    const remote = makeDraftRemote()
    const view = fixture(remote)
    await waitFor(() => { expect(view.fixture.adapter.mounts).toHaveLength(1) })

    act(() => {
      view.fixture.changes.markDirty(TAB_ID, true)
      view.fixture.changes.warnClose(TAB_ID)
    })

    expect(await view.findByText(en.closeWarning)).toBeDefined()
    expect(view.fixture.changes.isDirty(TAB_ID)).toBe(true)
    fireEvent.click(view.getByText(en.cancel))
    await waitFor(() => { expect(view.queryByText(en.closeWarning)).toBeNull() })
    expect(view.fixture.changes.isDirty(TAB_ID)).toBe(true)
  })

  it('follows the open document in the move tracker', async () => {
    const remote = makeDraftRemote()
    const view = renderPanel({ meta: { absolutePath: PATH, version: 'v1' }, remote: remote.face })

    await waitFor(() => {
      expect(view.fixture.trackDocument).toHaveBeenCalledWith(SESSION_ID, expect.any(String), 'work/budget.xlsx')
      expect(view.fixture.observeFile).toHaveBeenCalledWith(SESSION_ID, expect.any(String), 'work/budget.xlsx', 'v1')
    })
    view.unmount()
    expect(view.fixture.forgetDocument).toHaveBeenCalledOnce()
  })

  it('reports a file resource that observed the document missing', async () => {
    const remote = makeDraftRemote()
    const view = renderPanel({
      meta: { absolutePath: PATH, version: 'v1' },
      failure: new RemoteError('workspace-file/not-found', 'gone', { path: PATH }),
      remote: remote.face,
    })

    await waitFor(() => {
      expect(view.fixture.observeFile).toHaveBeenCalledWith(SESSION_ID, expect.any(String), 'work/budget.xlsx', undefined)
    })
  })

  it('ignores a failed file resource that did not report the document missing', () => {
    const remote = makeDraftRemote()
    const view = renderPanel({
      meta: { absolutePath: PATH, version: 'v1' },
      failure: new RemoteError('workspace-file/not-regular-file', 'denied', { path: PATH, kind: 'directory' }),
      remote: remote.face,
    })

    expect(view.fixture.observeFile).not.toHaveBeenCalled()
  })
})
