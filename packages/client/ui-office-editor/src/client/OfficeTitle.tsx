/**
 * The office type's chip title: the file type's coloured sheet before the name
 * the registry captured at open time. Registered under
 * `sidebar.right.pane.tab.title`; without it the chip would show the bare name.
 */
import type { ReactNode } from 'react'
import type { PropsRuntime } from '@deepseek-ai/dsh-client-ui-slots'
import { FileTypeIcon, classifyFileType } from '@deepseek-ai/dsh-client-ui-primitives'
import css from './OfficePanel.module.css'

/** The title's composed props: the tab information hook. */
export type OfficeTitleProps = PropsRuntime<'sidebar.right.pane.tab.title'>

/**
 * The title as the chip and a floating panel's header show it.
 * @param props - the tab information hook.
 * @returns the file type's 16px sheet followed by the tab's title text.
 */
export function OfficeTitle({ useTabInfo }: OfficeTitleProps): ReactNode {
  const { tab } = useTabInfo()
  return (
    <>
      <FileTypeIcon kind={classifyFileType(tab.title)} size={16} className={css.titleIcon} />
      {tab.title}
    </>
  )
}
