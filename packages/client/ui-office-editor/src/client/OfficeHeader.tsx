/**
 * The office tab's pane header: one chip per tab in the pane, with the selected
 * one marked, replacing the kit's tab strip while an office tab is selected.
 *
 * The header must keep every other tab reachable, including non-office ones
 * whose bodies the office-type header does not own, so each chip activates
 * through the Sidebar's own focus intent and closes through its close action.
 */
import type { ReactNode } from 'react'
import type { PropsLocale, PropsRuntime } from '@forwork/client-ui-slots'
import { FileTypeIcon, IconCloseOutlineRegular, classifyFileType } from '@forwork/client-ui-primitives'
import css from './OfficePanel.module.css'

/** The header's composed props: the pane's owner data and this namespace's copy. */
export type OfficeHeaderProps =
  & PropsRuntime<'sidebar.right.pane.header'>
  & PropsLocale<'officeEditor'>

/**
 * Draw every tab in the pane as a 28px chip in one horizontally scrolling row.
 * @param props - the pane's tabs, the active tab, the focus and close actions, and copy.
 * @returns the pane header; the selected chip is marked and each chip can close its tab.
 */
export function OfficeHeader({ tabs, activeTabId, activateTab, closeTab, t }: OfficeHeaderProps): ReactNode {
  return (
    <nav className={css.paneHeader} aria-label={t('openDocuments')} data-office-pane-header>
      {tabs.map((tab) => {
        const active = tab.tabId === activeTabId
        return (
          <div
            key={tab.tabId}
            className={css.paneHeaderTab}
            data-office-pane-tab={tab.tabId}
            data-office-pane-tab-active={active ? '' : undefined}
          >
            <button
              type="button"
              className={css.paneHeaderActivate}
              aria-current={active ? 'true' : undefined}
              data-office-pane-tab-activate
              onClick={() => { activateTab(tab.tabId) }}
            >
              <FileTypeIcon kind={classifyFileType(tab.title)} size={16} className={css.paneHeaderIcon} />
              <span className={css.paneHeaderTitle} data-office-pane-tab-title>{tab.title}</span>
            </button>
            <button
              type="button"
              className={css.paneHeaderClose}
              aria-label={t('closeTab')}
              data-office-pane-tab-close
              onClick={() => { closeTab(tab.tabId) }}
            >
              <IconCloseOutlineRegular size={12} />
            </button>
          </div>
        )
      })}
    </nav>
  )
}
