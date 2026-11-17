// @vitest-environment jsdom
/**
 * A rebind moves one live document to a new path: the mounted tab keeps its
 * identity, saves the live unit to the new path, and restore reopens the
 * projection's identity instead of deriving a second one from the path.
 */
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { brandString } from '@forwork/brand'
import type { DocumentId } from '@forwork/session/src/document-events.ts'
import type { DocumentProjectionView } from '@forwork/session-document/client'
import { createOfficeTabChanges } from '../src/client/changes.ts'
import { createOfficeRestore, restoreDocuments } from '../src/client/restore.ts'
import { OfficePanel } from '../src/client/OfficePanel.tsx'
import { en } from '../src/client/locales.ts'
import { SESSION_ID, panelFixture } from './panel.tsx'
import { makeDraftRemote } from './draft-fixtures.client.ts'

afterEach(() => {
  cleanup()
})

const OLD = 'work/budget.xlsx'
const NEW = 'work/renamed.xlsx'
const DOCUMENT_ID = `${SESSION_ID}:${OLD}`

function addressOf(path: string): string {
  return `dsh-resource://file/session/s-1/${path}`
}

function meta(path: string, version = 'v1') {
  return { absolutePath: `/ws/${path}`, version }
}

function projectionOf(documentId: DocumentId, path: string): DocumentProjectionView {
  return {
    order: [documentId],
    documents: {
      [documentId]: {
        documentId,
        path,
        format: 'xlsx',
        contentVersion: 3,
        status: 'open',
        unsupported: [],
      },
    },
  }
}

describe('a rebound live document', () => {
  it('saves the mounted unit to the new path and never writes the old one', async () => {
    const remote = makeDraftRemote()
    const rebind = vi.fn()
    const changes = createOfficeTabChanges(rebind)
    const fixture = panelFixture({
      address: addressOf(OLD),
      changes,
      meta: meta(OLD),
      remote: remote.face,
    })
    const view = render(<OfficePanel {...fixture.props} />)
    await waitFor(() => { expect(fixture.attachCommandBridge).toHaveBeenCalled() })
    expect(fixture.attachCommandBridge.mock.calls.at(-1)?.[1].documentId).toBe(DOCUMENT_ID)
    act(() => { fixture.adapter.setDirty(true) })

    act(() => {
      changes.observeFile(SESSION_ID, DOCUMENT_ID, OLD, 'v1')
      changes.observeFile(SESSION_ID, DOCUMENT_ID, OLD, undefined)
      changes.observeFile(SESSION_ID, DOCUMENT_ID, NEW, 'v1')
    })

    expect(rebind).toHaveBeenCalledExactlyOnceWith(SESSION_ID, DOCUMENT_ID, NEW)
    expect(changes.locations.getSnapshot()).toEqual({ [DOCUMENT_ID]: NEW })
    expect(changes.source.getSnapshot()[fixture.props.useTabInfo().tab.id]).toEqual({ dirty: true, warned: false })
    expect(fixture.adapter.mounts).toHaveLength(1)
    expect(fixture.attachCommandBridge.mock.calls.at(-1)?.[1].documentId).toBe(DOCUMENT_ID)
    await waitFor(() => { expect(fixture.resourceAddresses.at(-1)).toBe(addressOf(NEW)) })

    await waitFor(() => { expect((view.getByText(en.save) as HTMLButtonElement).disabled).toBe(false) })
    fireEvent.click(view.getByText(en.save))
    await waitFor(() => { expect(remote.save).toHaveBeenCalledOnce() })
    expect(remote.save.mock.calls[0]?.[1]).toBe(NEW)
    expect(remote.save.mock.calls.some(call => call[1] === OLD)).toBe(false)
  })

  it('ignores a different file that now occupies the old path', async () => {
    const remote = makeDraftRemote()
    const changes = createOfficeTabChanges(vi.fn())
    const fixture = panelFixture({
      address: addressOf(OLD),
      changes,
      meta: meta(OLD),
      remote: remote.face,
    })
    const view = render(<OfficePanel {...fixture.props} />)
    await waitFor(() => { expect((view.getByText(en.save) as HTMLButtonElement).disabled).toBe(false) })

    act(() => {
      changes.observeFile(SESSION_ID, DOCUMENT_ID, OLD, 'v1')
      changes.observeFile(SESSION_ID, DOCUMENT_ID, OLD, undefined)
      changes.observeFile(SESSION_ID, DOCUMENT_ID, NEW, 'v1')
      changes.observeFile(SESSION_ID, DOCUMENT_ID, OLD, 'v2')
    })

    fireEvent.click(view.getByText(en.save))
    await waitFor(() => { expect(remote.save).toHaveBeenCalledOnce() })
    expect(fixture.adapter.mounts).toHaveLength(1)
    expect(remote.save.mock.calls[0]?.[1]).toBe(NEW)
    expect(remote.save.mock.calls.some(call => call[1] === OLD)).toBe(false)
  })
})

describe('restore after a rebind and restart', () => {
  it('opens one tab under the projection identity and reuses it for reads and drafts', async () => {
    const remote = makeDraftRemote()
    const changes = createOfficeTabChanges()
    const projected = brandString<DocumentId>(DOCUMENT_ID)
    const projection = projectionOf(projected, NEW)
    const restore = createOfficeRestore()
    const opened: string[] = []

    restoreDocuments({
      sessionId: SESSION_ID,
      projection,
      openTabs: [],
      bindDocument: (address, documentId) => { changes.bindDocument(address, documentId) },
      openResource: (address) => { opened.push(address) },
      restore,
    })

    expect(opened).toEqual([addressOf(NEW)])
    expect(changes.documentIdAt(addressOf(NEW))).toBe(projected)

    const fixture = panelFixture({
      address: opened[0] ?? '',
      changes,
      projection,
      meta: meta(NEW),
      remote: remote.face,
    })
    render(<OfficePanel {...fixture.props} />)

    await waitFor(() => { expect(remote.readDraft).toHaveBeenCalledOnce() })
    expect(remote.read).toHaveBeenCalledExactlyOnceWith(SESSION_ID, NEW, expect.any(AbortSignal))
    expect(remote.readDraft.mock.calls[0]?.[1]).toBe(projected)
    await waitFor(() => { expect(fixture.attachCommandBridge).toHaveBeenCalled() })
    expect(fixture.attachCommandBridge.mock.calls.at(-1)?.[1].documentId).toBe(projected)
  })
})
