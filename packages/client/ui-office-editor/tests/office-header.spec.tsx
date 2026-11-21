// @vitest-environment jsdom
/**
 * The office pane header: every tab in the pane as a chip, the selected one
 * marked, with activation and closing dispatched through the owner's actions.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render } from '@testing-library/react'
import { brandString } from '@forwork/brand'
import type { TabId } from '@forwork/client-ui-dockkit'
import type { GlobalStandardProps, SessionStandardProps } from '@forwork/client-ui-slots'
import type { SessionId } from '@forwork/session/types'
import type { SidebarRightPaneHeaderTab } from '@forwork/client-ui-sidebar-right/client'
import { OfficeHeader } from '../src/client/OfficeHeader.tsx'
import type { OfficeHeaderProps } from '../src/client/OfficeHeader.tsx'
import { en } from '../src/client/locales.ts'
import { translate } from './panel.tsx'

afterEach(cleanup)

const unused = (): never => { throw new Error('This isolated component does not consume framework hooks') }
const SESSION = brandString<SessionId>('s-office')
const standard: GlobalStandardProps & SessionStandardProps = {
  sessionId: SESSION, useSession: unused, useProjection: unused, useConversation: unused,
  useInput: unused, useChat: unused, useTrajectory: unused,
  usePanelInfo: unused, useSessions: unused, useSessionStatus: unused,
  useSessionRetainInfo: unused, useResource: unused, useWorkspaces: unused,
  useRemoteAccess: () => true, useAccess: () => true,
  useSessionAccess: () => true, useWorkspaceAccess: () => true,
  useContextMenu: unused,
  inputActions: {
    insertQuote: () => false, captureInsertion: unused, insertText: unused, setDraft: unused,
    persistDraft: unused, addAttachments: unused, removeAttachment: unused, pruneAttachments: unused, submit: unused,
  },
}

const DOCX = brandString<TabId>('tab-docx')
const XLSX = brandString<TabId>('tab-xlsx')
const FILES = brandString<TabId>('tab-files')

const TABS: readonly SidebarRightPaneHeaderTab[] = [
  { tabId: DOCX, title: 'roundtrip.docx', kind: 'office', active: true },
  { tabId: XLSX, title: 'roundtrip.xlsx', kind: 'office', active: false },
  { tabId: FILES, title: 'Files', kind: 'files', active: false },
]

function props(overrides: Partial<OfficeHeaderProps> = {}): OfficeHeaderProps {
  return {
    ...standard,
    tabs: TABS,
    activeTabId: DOCX,
    activateTab: vi.fn(),
    closeTab: vi.fn(),
    t: translate,
    ...overrides,
  }
}

describe('OfficeHeader', () => {
  it('draws every pane tab and marks the selected one', () => {
    const view = render(<OfficeHeader {...props()} />)
    const chips = [...view.container.querySelectorAll<HTMLElement>('[data-office-pane-tab]')]
    expect(chips.map(chip => chip.querySelector('[data-office-pane-tab-title]')?.textContent))
      .toEqual(['roundtrip.docx', 'roundtrip.xlsx', 'Files'])
    expect(chips.map(chip => chip.hasAttribute('data-office-pane-tab-active')))
      .toEqual([true, false, false])
    expect(view.getByLabelText(en.openDocuments)).toBe(view.container.querySelector('[data-office-pane-header]'))
  })

  it('activates the clicked tab', () => {
    const activateTab = vi.fn<OfficeHeaderProps['activateTab']>()
    const view = render(<OfficeHeader {...props({ activateTab })} />)
    fireEvent.click(view.getByText('roundtrip.xlsx'))
    expect(activateTab).toHaveBeenCalledExactlyOnceWith(XLSX)
  })

  it('closes a tab through its localized close control', () => {
    const closeTab = vi.fn<OfficeHeaderProps['closeTab']>()
    const view = render(<OfficeHeader {...props({ closeTab })} />)
    const controls = view.getAllByRole('button', { name: en.closeTab })
    expect(controls).toHaveLength(TABS.length)
    fireEvent.click(controls[2]!)
    expect(closeTab).toHaveBeenCalledExactlyOnceWith(FILES)
  })
})
