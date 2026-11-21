// @vitest-environment jsdom
/**
 * The office body without its old open-document list: the pane header owns tab
 * navigation now, so the body keeps only its save state and editor surface.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, render } from '@testing-library/react'
import { brandString } from '@forwork/brand'
import type { TabId } from '@forwork/client-ui-dockkit'
import type { SidebarRightOpenTab } from '@forwork/client-ui-sidebar-right/client'
import { sessionFileAddress } from '@forwork/util-workspace-path'
import { RemoteError } from '@forwork/typert-protocol'
import type { RemoteResult } from '@forwork/typert-protocol'
import type { OfficeDocumentReadResult, OfficeDocumentSnapshot } from '@forwork/office-exchange/types'
import { OfficePanel } from '../src/client/OfficePanel.tsx'
import type { OfficeDocumentRemote } from '../src/client/useOfficeDocument.ts'
import { OFFICE_KIND } from '../src/client/definition.ts'
import { en } from '../src/client/locales.ts'
import { panelFixture, renderPanel, SESSION_ID, TAB_ID } from './panel.tsx'

afterEach(cleanup)

const SNAPSHOT: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book' } }
const BUDGET = sessionFileAddress('s-1', 'work/budget.xlsx')
const REPORT = sessionFileAddress('s-1', 'work/report.docx')

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

describe('OfficePanel chrome', () => {
  it('draws no document list or file name while dirty and save stay visible', () => {
    const openTabs: readonly SidebarRightOpenTab[] = [
      { sessionId: SESSION_ID, tabId: TAB_ID, kind: OFFICE_KIND, contentId: BUDGET },
      { sessionId: SESSION_ID, tabId: brandString<TabId>('tab-2'), kind: OFFICE_KIND, contentId: REPORT },
    ]
    const fixture = panelFixture({ address: BUDGET, openTabs, remote: editableRemote() })
    fixture.changes.markDirty(TAB_ID, true)
    const view = render(<OfficePanel {...fixture.props} />)
    expect(view.container.querySelector('[data-office-documents]')).toBeNull()
    expect(view.container.querySelector('[data-office-document]')).toBeNull()
    expect(view.container.querySelector('[data-office-file-name]')).toBeNull()
    expect(view.queryByText('report.docx')).toBeNull()
    expect(view.container.querySelector('[data-office-dirty]')?.textContent).toBe(en.unsaved)
    expect(view.container.querySelector('[data-office-save]')?.textContent).toBe(en.save)
  })

  it('reports a file resource that no longer resolves as missing', () => {
    const failure = new RemoteError('workspace-file/not-found', 'no entry', { path: 'work/budget.xlsx' })
    const view = renderPanel({ address: BUDGET, remote: editableRemote(), failure })
    expect(view.fixture.markMissing).toHaveBeenCalledWith(SESSION_ID, `${SESSION_ID}:work/budget.xlsx`)
  })
})
