/** Translation from a tab address to the Session and workspace path the Remote reads. */
import { brandString } from '@deepseek-ai/dsh-brand'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import type { LegacyOfficeFormat } from '@deepseek-ai/dsh-office-exchange/types'
import { fileAddressFor, isAbsoluteWorkspacePath, parseFileAddress } from '@deepseek-ai/dsh-util-workspace-path'

/** One open document target: the Session whose workspace owns it and the path handed to the Remote. */
export interface OfficeOpenTarget {
  /** The Session whose workspace resolves `path`. */
  readonly sessionId: SessionId
  /** The path the Remote receives. */
  readonly path: string
}

/**
 * Read the Session and path out of an `office` tab address.
 * @param address - the tab's `dsh-resource://file/session/...` address.
 * @returns the Session and path to hand the Remote.
 * @throws when the address is not a Session-scoped file address; the tab registry only routes those here.
 */
export function openTargetOf(address: string): OfficeOpenTarget {
  const parsed = parseFileAddress(address)
  if (parsed?.scope !== 'session') {
    throw new Error(`ui-office-editor: "${address}" is not a Session-scoped file address`)
  }
  return { sessionId: brandString<SessionId>(parsed.sessionId), path: parsed.path }
}

/**
 * The legacy format a path needs converted before it can be edited.
 * @param path - a workspace path, addressed or absolute.
 * @returns `doc`, `xls`, or `ppt` for a legacy binary source, `undefined` for an editable or unrelated file.
 */
export function legacyFormatOf(path: string): LegacyOfficeFormat | undefined {
  const dot = path.lastIndexOf('.')
  if (dot < 0) return undefined
  const suffix = path.slice(dot + 1).toLowerCase()
  return suffix === 'doc' || suffix === 'xls' || suffix === 'ppt' ? suffix : undefined
}

/**
 * The Session workspace root implied by one addressed file and its absolute
 * path. The converted copy the Remote returns only has to be checked against a
 * root the Client can name; a Session address that is already absolute leaves
 * the root unknown, and the Remote's own root confinement stands alone.
 * @param addressPath - the path part of the Session file address that was opened.
 * @param absolutePath - the file's absolute path as the resource provider reported it.
 * @returns the root without a trailing separator, or `undefined` when it cannot be derived.
 */
export function workspaceRootOf(addressPath: string, absolutePath: string | undefined): string | undefined {
  if (absolutePath === undefined || isAbsoluteWorkspacePath(addressPath)) return undefined
  const relative = addressPath.replace(/\\/g, '/').replace(/^(?:\.\/)+/, '')
  const absolute = absolutePath.replace(/\\/g, '/')
  if (relative === '' || !absolute.endsWith(`/${relative}`)) return undefined
  return absolute.slice(0, absolute.length - relative.length - 1) || '/'
}

/**
 * Address a converted copy, refusing one the Client can prove is outside the
 * Session workspace. The conversion Remote writes inside the workspace root
 * so a path outside the root this Session reports is a conversion
 * defect the user must see rather than a tab that fails on first read.
 * @param sessionId - the Session whose workspace owns the copy.
 * @param convertedPath - the path the conversion Remote returned.
 * @param workspaceRoot - the workspace root derived from the opened source, when known.
 * @returns the `dsh-resource://file/...` address, or `undefined` for an out-of-workspace copy.
 */
export function convertedAddress(
  sessionId: SessionId,
  convertedPath: string,
  workspaceRoot: string | undefined,
): string | undefined {
  if (workspaceRoot !== undefined) {
    const root = workspaceRoot.replace(/\\/g, '/').replace(/\/+$/, '')
    if (!convertedPath.replace(/\\/g, '/').startsWith(`${root}/`)) return undefined
  }
  return fileAddressFor(sessionId, workspaceRoot, convertedPath)
}
