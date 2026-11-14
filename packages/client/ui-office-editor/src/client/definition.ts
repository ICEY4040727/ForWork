/**
 * Stage one of this package's registration: what the `office` tab type IS.
 *
 * The type claims Session-scoped `dsh-resource://file/session/<sessionId>/<path>`
 * addresses whose basename ends in `.docx` or `.xlsx`, and the legacy `.doc`,
 * `.xls`, and `.ppt` sources it converts on open. `.pptx` is deliberately
 * absent: presentations stay read-only in the document preview, so a converted
 * `.ppt` copy lands there too. A page-level glob narrows the candidates and
 * `canOpen` vetoes an address whose URI merely ends in one of those suffixes.
 */
import type { SidebarRightTabDefinition } from '@deepseek-ai/dsh-client-ui-sidebar-right/client'
import { parseFileAddress } from '@deepseek-ai/dsh-util-workspace-path'

/** The tab kind this package owns. */
export const OFFICE_KIND = 'office'

/** This implementation's identity in the tab system: the key its body and title register under. */
export const OFFICE_ID = '@deepseek-ai/dsh-client-ui-office-editor'

/**
 * The tab title for one `file:` address: its decoded basename.
 *
 * The whole address stays the content identity, so two files with one name in
 * different directories are two tabs; only the chip text is shortened.
 * @param address - a `file:`-shaped address.
 * @returns the decoded last path segment, or the address itself when it has none.
 */
export function basenameOf(address: string): string {
  const name = address.slice(address.lastIndexOf('/') + 1)
  if (name === '') return address
  try {
    return decodeURIComponent(name)
  } catch {
    // A malformed percent sequence is still a name; showing it raw beats refusing the address.
    return name
  }
}

/**
 * The office type's registry definition.
 * @returns the definition to register.
 */
export function officeDefinition(): SidebarRightTabDefinition {
  return {
    id: OFFICE_ID,
    kind: OFFICE_KIND,
    patterns: ['*.docx', '*.xlsx', '*.doc', '*.xls', '*.ppt'],
    priority: 'builtin',
    keepMounted: true,
    canOpen: address => parseFileAddress(address)?.scope === 'session',
    title: basenameOf,
  }
}
