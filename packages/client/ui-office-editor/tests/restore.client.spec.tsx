// @vitest-environment jsdom
/**
 * Restore of a Session's durable open-document set: one pass per Session,
 * projection order, missing documents reported instead of opened, no duplicate
 * tabs, and the existing draft-recovery offer for a restored document.
 */
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { brandString } from '@forwork/brand'
import type { DocumentId } from '@forwork/session/src/document-events.ts'
import type { SessionId } from '@forwork/session/types'
import type { TabId } from '@forwork/client-ui-dockkit'
import type { SidebarRightOpenTab } from '@forwork/client-ui-sidebar-right/client'
import type { DocumentProjectionEntry, DocumentProjectionView } from '@forwork/session-document/client'
import type { SessionListState, SessionSummary } from '@forwork/api-session-controller/client'
import { createSnapshotStore } from '@forwork/client-store'
import { en } from '../src/client/locales.ts'
import { OfficePanel } from '../src/client/OfficePanel.tsx'
import type { OfficePanelProps } from '../src/client/OfficePanel.tsx'
import { createOfficeRestore, restoreOnSessionActivation, restorePlan } from '../src/client/restore.ts'
import { SESSION_ID, panelFixture, renderPanel, translate } from './panel.tsx'
import { draftAt, makeDraftRemote, success } from './draft-fixtures.client.ts'

afterEach(() => {
  cleanup()
})

const BUDGET = 'work/budget.xlsx'
const REPORT = 'work/report.docx'
const NOTES = 'work/notes.docx'
const GONE = 'work/gone.docx'

function addressOf(path: string): string {
  return `dsh-resource://file/session/s-1/${path}`
}

function entryOf(path: string, status: DocumentProjectionEntry['status']): DocumentProjectionEntry {
  const documentId = brandString<DocumentId>(`${SESSION_ID}:${path}`)
  return {
    documentId,
    path,
    format: path.endsWith('.xlsx') ? 'xlsx' : 'docx',
    contentVersion: 0,
    status,
    unsupported: [],
  }
}

function projectionOf(paths: readonly string[], missing: readonly string[] = []): DocumentProjectionView {
  const entries = paths.map(path => entryOf(path, missing.includes(path) ? 'missing' : 'open'))
  const documents: Record<string, DocumentProjectionEntry> = {}
  for (const entry of entries) documents[entry.documentId] = entry
  return { order: entries.map(entry => entry.documentId), documents }
}

function openedTabs(...paths: readonly string[]): readonly SidebarRightOpenTab[] {
  return paths.map((path, index) => ({
    sessionId: SESSION_ID,
    tabId: brandString<TabId>(`tab-${index}`),
    kind: 'office',
    contentId: addressOf(path),
  }))
}

describe('restorePlan', () => {
  it('splits open paths from missing ones, keeping projection order', () => {
    const view = projectionOf([REPORT, GONE, NOTES], [GONE])

    expect(restorePlan(view)).toEqual({ open: [REPORT, NOTES], missing: [GONE] })
  })

  it('reports nothing for an empty open set', () => {
    expect(restorePlan({ order: [], documents: {} })).toEqual({ open: [], missing: [] })
  })

  it('skips an order entry with no document record', () => {
    const view = projectionOf([REPORT])
    const danglingId = brandString<DocumentId>(`${SESSION_ID}:${GONE}`)
    const dangling: DocumentProjectionView = { ...view, order: [danglingId, ...view.order] }

    expect(restorePlan(dangling)).toEqual({ open: [REPORT], missing: [] })
  })
})

describe('createOfficeRestore', () => {
  it('claims one pass per Session and keeps a notice until it is dismissed', () => {
    const restore = createOfficeRestore()

    expect(restore.claim(SESSION_ID)).toBe(true)
    expect(restore.claim(SESSION_ID)).toBe(false)
    expect(restore.source.getSnapshot().restored).toEqual([SESSION_ID])

    restore.reportMissing(SESSION_ID, [])
    expect(restore.source.getSnapshot().missing).toEqual({})

    restore.reportMissing(SESSION_ID, [GONE])
    expect(restore.source.getSnapshot().missing[SESSION_ID]).toEqual([GONE])

    restore.dismissMissing(SESSION_ID)
    expect(restore.source.getSnapshot().missing[SESSION_ID]).toBeUndefined()
    const settled = restore.source.getSnapshot()
    restore.dismissMissing(SESSION_ID)
    expect(restore.source.getSnapshot()).toBe(settled)
  })

  it('keeps another Session notice when one is dismissed', () => {
    const restore = createOfficeRestore()
    const other = brandString<SessionId>('s-2')

    restore.reportMissing(SESSION_ID, [GONE])
    restore.reportMissing(other, [NOTES])
    restore.dismissMissing(SESSION_ID)

    expect(restore.source.getSnapshot().missing[SESSION_ID]).toBeUndefined()
    expect(restore.source.getSnapshot().missing[other]).toEqual([NOTES])
  })
})

describe('useDocumentRestore', () => {
  it('opens the missing documents once in projection order and skips an already-open tab', () => {
    const remote = makeDraftRemote()
    const view = renderPanel({
      projection: projectionOf([REPORT, BUDGET, NOTES]),
      openTabs: openedTabs(BUDGET),
      remote: remote.face,
    })

    expect(view.fixture.openResource.mock.calls.map(call => call[0]))
      .toEqual([addressOf(REPORT), addressOf(NOTES)])
  })

  it('runs one pass per Session across mounted bodies and later projection updates', () => {
    const remote = makeDraftRemote()
    const restore = createOfficeRestore()
    const first = panelFixture({ projection: projectionOf([REPORT]), remote: remote.face, restore })
    const second = panelFixture({ projection: projectionOf([REPORT]), remote: remote.face, restore })
    const view = render(<><OfficePanel {...first.props} /><OfficePanel {...second.props} /></>)

    expect(first.openResource).toHaveBeenCalledOnce()
    expect(second.openResource).not.toHaveBeenCalled()

    const later = {
      ...first.props,
      useProjection: (key: string) => key === 'document' ? projectionOf([REPORT, NOTES]) : undefined,
    } as OfficePanelProps
    view.rerender(<><OfficePanel {...later} /><OfficePanel {...second.props} /></>)
    expect(first.openResource).toHaveBeenCalledOnce()
  })

  it('reports a missing document instead of opening it and hides the notice on dismissal', async () => {
    const remote = makeDraftRemote()
    const view = renderPanel({
      projection: projectionOf([REPORT, GONE], [GONE]),
      remote: remote.face,
    })

    expect(view.fixture.openResource.mock.calls.map(call => call[0])).toEqual([addressOf(REPORT)])
    expect(await view.findByText(translate('restoreMissing', { paths: GONE }))).toBeDefined()
    fireEvent.click(view.getByText(en.dismiss))
    await waitFor(() => { expect(view.queryByText(translate('restoreMissing', { paths: GONE }))).toBeNull() })
  })

  it('opens nothing and still reports the missing set when every document is missing', async () => {
    const remote = makeDraftRemote()
    const view = renderPanel({ projection: projectionOf([GONE], [GONE]), remote: remote.face })

    expect(view.fixture.openResource).not.toHaveBeenCalled()
    expect(await view.findByText(translate('restoreMissing', { paths: GONE }))).toBeDefined()
  })

  it('offers a restored document a draft captured from the current file version', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockResolvedValueOnce(success(draftAt('v1')))
    const view = renderPanel({
      projection: projectionOf([BUDGET]),
      openTabs: openedTabs(BUDGET),
      remote: remote.face,
      meta: { absolutePath: '/work/budget.xlsx', version: 'v1' },
    })

    expect(await view.findByText(en.draftFound)).toBeDefined()
    expect(view.queryByText(en.draftStale)).toBeNull()
  })

  it('labels a restored document draft from an earlier version stale', async () => {
    const remote = makeDraftRemote()
    remote.readDraft.mockResolvedValueOnce(success(draftAt('v0')))
    const view = renderPanel({
      projection: projectionOf([BUDGET]),
      openTabs: openedTabs(BUDGET),
      remote: remote.face,
      meta: { absolutePath: '/work/budget.xlsx', version: 'v1' },
    })

    expect(await view.findByText(en.draftStale)).toBeDefined()
    expect(view.queryByText(en.draftFound)).toBeNull()
  })

  it('offers nothing to a restored document without a draft', async () => {
    const remote = makeDraftRemote()
    const view = renderPanel({
      projection: projectionOf([BUDGET]),
      openTabs: openedTabs(BUDGET),
      remote: remote.face,
      meta: { absolutePath: '/work/budget.xlsx', version: 'v1' },
    })

    await waitFor(() => { expect(remote.readDraft).toHaveBeenCalledOnce() })
    expect(view.queryByText(en.draftFound)).toBeNull()
    expect(view.queryByText(en.draftStale)).toBeNull()
  })
})

function sessionList(
  projection: DocumentProjectionView | undefined,
  active = true,
): SessionListState {
  const summary: SessionSummary = {
    id: SESSION_ID,
    displayTitle: 'Session',
    running: false,
    blank: false,
    updatedAt: 0,
    retainedBy: active ? { mainView: 1 } : {},
  }
  return {
    ids: [SESSION_ID],
    byId: { [SESSION_ID]: summary },
    phase: 'ready',
    projectionsBySession: projection === undefined ? {} : {
      [SESSION_ID]: { values: { document: projection }, state: 'ready', error: null },
    },
  }
}

describe('restoreOnSessionActivation', () => {
  it('restores an active Session that has no Office tab open', () => {
    const sessions = createSnapshotStore(sessionList(projectionOf([REPORT, GONE], [GONE])))
    const openTabs = createSnapshotStore<readonly SidebarRightOpenTab[]>([])
    const restore = createOfficeRestore()
    const openResourceIn = vi.fn()

    const stop = restoreOnSessionActivation({ sessions, openTabs, restore, openResourceIn })

    expect(openResourceIn).toHaveBeenCalledExactlyOnceWith(SESSION_ID, addressOf(REPORT))
    expect(restore.source.getSnapshot().missing[SESSION_ID]).toEqual([GONE])
    stop()
  })

  it('runs one pass per Session across repeated list notifications', () => {
    const sessions = createSnapshotStore(sessionList(projectionOf([REPORT])))
    const openTabs = createSnapshotStore<readonly SidebarRightOpenTab[]>([])
    const restore = createOfficeRestore()
    const openResourceIn = vi.fn()

    const stop = restoreOnSessionActivation({ sessions, openTabs, restore, openResourceIn })
    sessions.set(sessionList(projectionOf([REPORT, NOTES])))
    sessions.set(sessionList(projectionOf([REPORT, NOTES])))
    expect(openResourceIn).toHaveBeenCalledOnce()
    stop()
  })

  it('waits for the projection and ignores an inactive Session', () => {
    const sessions = createSnapshotStore(sessionList(undefined))
    const openTabs = createSnapshotStore<readonly SidebarRightOpenTab[]>([])
    const restore = createOfficeRestore()
    const openResourceIn = vi.fn()

    const stop = restoreOnSessionActivation({ sessions, openTabs, restore, openResourceIn })
    expect(openResourceIn).not.toHaveBeenCalled()

    sessions.set(sessionList({ order: [], documents: {} }))
    expect(openResourceIn).not.toHaveBeenCalled()

    sessions.set(sessionList(projectionOf([REPORT]), false))
    expect(openResourceIn).not.toHaveBeenCalled()

    sessions.set(sessionList(projectionOf([REPORT])))
    expect(openResourceIn).toHaveBeenCalledExactlyOnceWith(SESSION_ID, addressOf(REPORT))
    stop()
  })

  it('stops following activation after disposal', () => {
    const sessions = createSnapshotStore(sessionList(projectionOf([REPORT])))
    const openTabs = createSnapshotStore<readonly SidebarRightOpenTab[]>([])
    const restore = createOfficeRestore()
    const openResourceIn = vi.fn()

    const stop = restoreOnSessionActivation({ sessions, openTabs, restore, openResourceIn })
    stop()
    sessions.set(sessionList(projectionOf([REPORT, NOTES])))
    expect(openResourceIn).toHaveBeenCalledOnce()
  })
})
