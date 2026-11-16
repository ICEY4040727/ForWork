/**
 * The Session's open Office documents, as a compact list above the editor.
 *
 * Rendering is additive: the list names every open `office` tab of this
 * Session, marks the document this body is showing, and focuses a listed tab
 * through the Sidebar's public tab actions. The entries today come from the
 * Sidebar's open-tab inventory (`sidebarRight.openTabs`) joined with this
 * body's own tab; the `document` Session projection will replace that source
 * with the projection's open-document set, so nothing here holds its own copy.
 */
import type { ReactNode } from 'react'
import type { TabId } from '@forwork/client-ui-dockkit'
import type { TranslateNS } from '@forwork/client-ui-slots'
import css from './OfficePanel.module.css'

/** One listed document. */
export interface OfficeDocumentEntry {
  /** Identity of the tab holding the document. */
  readonly tabId: TabId
  /** The document's file resource address, which re-opening focuses. */
  readonly contentId: string
  /** Display name for the row. */
  readonly title: string
  /** Whether this panel is the body of that document's tab. */
  readonly active: boolean
}

/** The list's props: the entries, the focus action, and the editor's copy. */
export interface DocumentPanelProps {
  /** Open Office documents of this Session, in Sidebar inventory order. */
  readonly documents: readonly OfficeDocumentEntry[]
  /**
   * Focus one listed document.
   * @param entry - the chosen document.
   */
  readonly onActivate: (entry: OfficeDocumentEntry) => void
  readonly t: TranslateNS<'officeEditor'>
}

/**
 * Draw the open-document list.
 * @param props - the entries, the focus action, and the copy seat.
 * @returns a navigation list of open documents, one row per tab, or `null` when the Session has none.
 */
export function DocumentPanel({ documents, onActivate, t }: DocumentPanelProps): ReactNode {
  if (documents.length === 0) return null
  return (
    <nav className={css.documentPanel} aria-label={t('openDocuments')} data-office-documents>
      <ul className={css.documentList}>
        {documents.map(document => (
          <li key={document.tabId}>
            <button
              type="button"
              className={css.document}
              aria-current={document.active ? 'true' : undefined}
              data-office-document={document.tabId}
              data-office-document-active={document.active ? '' : undefined}
              onClick={() => { onActivate(document) }}
            >
              {document.title}
            </button>
          </li>
        ))}
      </ul>
    </nav>
  )
}
