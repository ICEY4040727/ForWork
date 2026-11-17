// @vitest-environment jsdom
/** The panel's injected face binds the Remote, the adapter, and the shared tab state. */
import { describe, expect, it, vi } from 'vitest'
import { brandString } from '@forwork/brand'
import type { TabId } from '@forwork/client-ui-dockkit'
import type { SessionId } from '@forwork/session/types'
import { createSnapshotStore } from '@forwork/client-store'
import type { SidebarRightOpenTab } from '@forwork/client-ui-sidebar-right/client'
import { createMockAdapter } from '../src/client/adapter.ts'
import type { OfficeDocumentSnapshot } from '@forwork/office-exchange/types'
import { createOfficeTabChanges } from '../src/client/changes.ts'
import { createOfficeLegacyNotices } from '../src/client/notices.ts'
import { createOfficeRestore } from '../src/client/restore.ts'
import { officePanelInjected } from '../src/client/face.ts'

const TAB = brandString<TabId>('tab-1')
const snapshot: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book' } }

describe('officePanelInjected', () => {
  it('mounts through the adapter and writes the shared change state', async () => {
    const adapter = createMockAdapter()
    const changes = createOfficeTabChanges()
    const notices = createOfficeLegacyNotices()
    const openTabs = createSnapshotStore<readonly SidebarRightOpenTab[]>([])
    const read = vi.fn()
    const convert = vi.fn()
    const save = vi.fn()
    const writeDraft = vi.fn()
    const readDraft = vi.fn()
    const clearDraft = vi.fn()
    const attachCommandBridge = vi.fn(() => () => {})
    const markMissing = vi.fn()
    const restore = createOfficeRestore()
    const face = officePanelInjected({
      read, convert, save, writeDraft, readDraft, clearDraft, adapter, attachCommandBridge, markMissing, changes, notices,
      restore, openTabs,
    })
    await face.remote.read(brandString<SessionId>('s-1'), 'work/budget.xlsx')
    await face.remote.convertLegacy(brandString<SessionId>('s-1'), 'work/legacy.doc')
    await face.remote.save(brandString<SessionId>('s-1'), 'work/budget.xlsx', snapshot, 'xlsx', 'v1', { actor: 'user' })
    await face.remote.writeDraft(brandString<SessionId>('s-1'), 'doc', 'work/budget.xlsx', 'v1', snapshot)
    await face.remote.readDraft(brandString<SessionId>('s-1'), 'doc')
    await face.remote.clearDraft(brandString<SessionId>('s-1'), 'doc')
    expect(read).toHaveBeenCalledOnce()
    expect(convert).toHaveBeenCalledOnce()
    expect(save).toHaveBeenCalledOnce()
    expect(writeDraft).toHaveBeenCalledOnce()
    expect(readDraft).toHaveBeenCalledOnce()
    expect(clearDraft).toHaveBeenCalledOnce()
    const container = document.createElement('div')
    const handle = await face.mountEditor(container, { kind: 'workbook', snapshot, onChange: vi.fn() })
    expect(adapter.mounts).toHaveLength(1)
    expect(handle.getSnapshot()).toBe(snapshot)
    expect(face.attachCommandBridge).toBe(attachCommandBridge)
    face.markMissing(brandString<SessionId>('s-1'), 's-1:work/budget.xlsx')
    expect(markMissing).toHaveBeenCalledWith(brandString<SessionId>('s-1'), 's-1:work/budget.xlsx')
    face.markTabDirty(TAB, true)
    expect(changes.isDirty(TAB)).toBe(true)
    face.dismissCloseWarning(TAB)
    expect(face.hooks.officeTab.getSnapshot()[TAB]).toEqual({ dirty: true, warned: false })
    const sessionId = brandString<SessionId>('s-1')
    face.trackDocument(sessionId, `${sessionId}:work/budget.xlsx`, 'work/budget.xlsx')
    face.observeFile(sessionId, `${sessionId}:work/budget.xlsx`, 'work/budget.xlsx', 'v1')
    face.observeFile(sessionId, `${sessionId}:work/budget.xlsx`, 'work/budget.xlsx', undefined)
    expect(face.hooks.documentLocations.getSnapshot()).toEqual({ [`${sessionId}:work/budget.xlsx`]: 'work/budget.xlsx' })
    face.bindDocument('dsh-resource://file/session/s-1/work/budget.xlsx', `${sessionId}:work/budget.xlsx`)
    expect(face.documentIdAt('dsh-resource://file/session/s-1/work/budget.xlsx')).toBe(`${sessionId}:work/budget.xlsx`)
    face.forgetDocument(`${sessionId}:work/budget.xlsx`)
    expect(face.hooks.documentLocations.getSnapshot()).toEqual({})
    face.rememberLegacySource('dsh-resource://file/session/s-1/work/report.docx', 'work/report.doc')
    expect(face.hooks.legacyNotice.getSnapshot()['dsh-resource://file/session/s-1/work/report.docx'])
      .toEqual({ source: 'work/report.doc', dismissed: false })
    face.dismissLegacyNotice('dsh-resource://file/session/s-1/work/report.docx')
    expect(face.hooks.legacyNotice.getSnapshot()['dsh-resource://file/session/s-1/work/report.docx'])
      .toEqual({ source: 'work/report.doc', dismissed: true })
    expect(face.restore.claim(brandString<SessionId>('s-1'))).toBe(true)
    face.restore.reportMissing(brandString<SessionId>('s-1'), ['work/gone.docx'])
    expect(restore.source.getSnapshot().missing['s-1']).toEqual(['work/gone.docx'])
    expect(face.hooks.restore).toBe(restore.source)
    face.dismissRestoreNotice(brandString<SessionId>('s-1'))
    expect(restore.source.getSnapshot().missing['s-1']).toBeUndefined()
    expect(face.hooks.sidebarTabs).toBe(openTabs)
    handle.dispose()
  })
})
