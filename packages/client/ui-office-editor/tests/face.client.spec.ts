// @vitest-environment jsdom
/** The panel's injected face binds the Remote, the adapter, and the shared tab state. */
import { describe, expect, it, vi } from 'vitest'
import { brandString } from '@deepseek-ai/dsh-brand'
import type { TabId } from '@deepseek-ai/dsh-client-ui-dockkit'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import { createSnapshotStore } from '@deepseek-ai/dsh-client-store'
import type { SidebarRightOpenTab } from '@deepseek-ai/dsh-client-ui-sidebar-right/client'
import { createMockAdapter } from '../src/client/adapter.ts'
import type { OfficeDocumentSnapshot } from '@deepseek-ai/dsh-office-exchange/types'
import { createOfficeTabChanges } from '../src/client/changes.ts'
import { createOfficeLegacyNotices } from '../src/client/notices.ts'
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
    const face = officePanelInjected({
      read, convert, save, writeDraft, readDraft, clearDraft, adapter, changes, notices, openTabs,
    })
    await face.remote.read(brandString<SessionId>('s-1'), 'work/budget.xlsx')
    await face.remote.convertLegacy(brandString<SessionId>('s-1'), 'work/legacy.doc')
    await face.remote.save(brandString<SessionId>('s-1'), 'work/budget.xlsx', snapshot, 'xlsx', 'v1')
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
    face.markTabDirty(TAB, true)
    expect(changes.isDirty(TAB)).toBe(true)
    face.dismissCloseWarning(TAB)
    expect(face.hooks.officeTab.getSnapshot()[TAB]).toEqual({ dirty: true, warned: false })
    face.rememberLegacySource('dsh-resource://file/session/s-1/work/report.docx', 'work/report.doc')
    expect(face.hooks.legacyNotice.getSnapshot()['dsh-resource://file/session/s-1/work/report.docx'])
      .toEqual({ source: 'work/report.doc', dismissed: false })
    face.dismissLegacyNotice('dsh-resource://file/session/s-1/work/report.docx')
    expect(face.hooks.legacyNotice.getSnapshot()['dsh-resource://file/session/s-1/work/report.docx'])
      .toEqual({ source: 'work/report.doc', dismissed: true })
    expect(face.hooks.sidebarTabs).toBe(openTabs)
    handle.dispose()
  })
})
