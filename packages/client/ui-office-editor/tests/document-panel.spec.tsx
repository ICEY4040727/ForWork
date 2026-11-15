// @vitest-environment jsdom
/**
 * The open-document panel.
 *
 * It lists the Session's open Office documents, marks the one this body shows,
 * and focuses a listed tab through the Sidebar's public tab actions. The list
 * is derived from the Sidebar's open-tab inventory; T037's `document` Session
 * projection will replace that source.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render } from '@testing-library/react'
import { brandString } from '@deepseek-ai/dsh-brand'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import type { TabId } from '@deepseek-ai/dsh-client-ui-dockkit'
import type { SidebarRightOpenTab } from '@deepseek-ai/dsh-client-ui-sidebar-right/client'
import { sessionFileAddress } from '@deepseek-ai/dsh-util-workspace-path'
import { RemoteError } from '@deepseek-ai/dsh-typert-protocol'
import type { RemoteResult } from '@deepseek-ai/dsh-typert-protocol'
import type { OfficeDocumentReadResult, OfficeDocumentSnapshot } from '@deepseek-ai/dsh-office-exchange/types'
import { DocumentPanel } from '../src/client/DocumentPanel.tsx'
import type { OfficeDocumentEntry } from '../src/client/DocumentPanel.tsx'
import type { OfficeDocumentRemote } from '../src/client/useOfficeDocument.ts'
import { OFFICE_KIND } from '../src/client/definition.ts'
import { en } from '../src/client/locales.ts'
import { SESSION_ID, TAB_ID, renderPanel, translate } from './panel.tsx'

afterEach(cleanup)

const SNAPSHOT: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book' } }
const BUDGET = sessionFileAddress('s-1', 'work/budget.xlsx')
const REPORT = sessionFileAddress('s-1', 'work/report.docx')
const OTHER_TAB = brandString<TabId>('tab-2')

function editableRemote(): OfficeDocumentRemote {
  return {
    read: async (): Promise<RemoteResult<OfficeDocumentReadResult>> => ({
      ok: true,
      value: { format: 'xlsx', snapshot: SNAPSHOT, unsupported: [], version: 'v1' },
    }),
    convertLegacy: vi.fn(),
    save: vi.fn(),
    writeDraft: vi.fn(),
    readDraft: async () => ({ ok: true, value: null }),
    clearDraft: vi.fn(),
  }
}

function entry(overrides: Partial<OfficeDocumentEntry> = {}): OfficeDocumentEntry {
  return { tabId: TAB_ID, contentId: BUDGET, title: 'budget.xlsx', active: false, ...overrides }
}

describe('DocumentPanel', () => {
  it('lists every open document and marks the active one', () => {
    const view = render(<DocumentPanel
      documents={[entry(), entry({ tabId: OTHER_TAB, contentId: REPORT, title: 'report.docx', active: true })]}
      onActivate={vi.fn()}
      t={translate}
    />)
    const rows = view.getAllByRole('button')
    expect(rows.map(row => row.textContent)).toEqual(['budget.xlsx', 'report.docx'])
    expect(rows[0]?.hasAttribute('data-office-document-active')).toBe(false)
    expect(rows[1]?.getAttribute('data-office-document-active')).toBe('')
    expect(rows[1]?.getAttribute('aria-current')).toBe('true')
    expect(rows[0]?.getAttribute('aria-current')).toBeNull()
    expect(view.getByLabelText(en.openDocuments)).toBeDefined()
  })

  it('activates the clicked document', () => {
    const onActivate = vi.fn<(document: OfficeDocumentEntry) => void>()
    const view = render(<DocumentPanel
      documents={[entry(), entry({ tabId: OTHER_TAB, contentId: REPORT, title: 'report.docx' })]}
      onActivate={onActivate}
      t={translate}
    />)
    fireEvent.click(view.getByText('report.docx'))
    expect(onActivate).toHaveBeenCalledOnce()
    expect(onActivate.mock.calls[0]?.[0]?.contentId).toBe(REPORT)
  })
})

describe('OfficePanel document list', () => {
  const openTabs: readonly SidebarRightOpenTab[] = [
    { sessionId: SESSION_ID, tabId: TAB_ID, kind: OFFICE_KIND, contentId: BUDGET },
    { sessionId: SESSION_ID, tabId: OTHER_TAB, kind: OFFICE_KIND, contentId: REPORT },
    { sessionId: SESSION_ID, tabId: brandString<TabId>('tab-3'), kind: 'text', contentId: sessionFileAddress('s-1', 'notes.md') },
    { sessionId: brandString<SessionId>('s-2'), tabId: brandString<TabId>('tab-4'), kind: OFFICE_KIND, contentId: sessionFileAddress('s-2', 'other.xlsx') },
  ]

  it('lists this Session office tabs and marks this body', () => {
    const view = renderPanel({ address: BUDGET, openTabs, remote: editableRemote() })
    const rows = view.getAllByRole('button').filter(row => row.hasAttribute('data-office-document'))
    expect(rows.map(row => row.textContent)).toEqual(['budget.xlsx', 'report.docx'])
    expect(rows[0]?.getAttribute('data-office-document-active')).toBe('')
    expect(rows[1]?.getAttribute('data-office-document-active')).toBeNull()
  })

  it('focuses a listed document through the tab action', () => {
    const view = renderPanel({ address: BUDGET, openTabs, remote: editableRemote() })
    fireEvent.click(view.getByText('report.docx'))
    expect(view.fixture.openResource).toHaveBeenCalledWith(REPORT)
  })

  it('reports a file resource that no longer resolves as missing', () => {
    const failure = new RemoteError('workspace-file/not-found', 'no entry', { path: 'work/budget.xlsx' })
    const view = renderPanel({ address: BUDGET, remote: editableRemote(), failure })
    expect(view.fixture.markMissing).toHaveBeenCalledWith(SESSION_ID, `${SESSION_ID}:work/budget.xlsx`)
  })
})
