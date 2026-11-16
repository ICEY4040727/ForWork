/**
 * Document identity and path resolution for the Office document tools.
 *
 * A model-visible `documentId` is the Session-scoped identity the editor
 * already uses: `<sessionId>:<path>`. The path is a locator, never identity;
 * every resolved target is confined to the Session's workspace root before a
 * command is sent or a file is touched.
 * @module @forwork/tool-office/documents
 */

import { resolve } from 'node:path'
import { assertWithinRoots } from '@forwork/office-exchange'
import { DocumentId } from '@forwork/session'
import type { DocumentId as DocumentIdValue, SessionId } from '@forwork/session'
import { documentIdFor } from '@forwork/session-document'
import { OfficeToolResultError } from './errors.ts'

export { documentIdFor }

/** One resolved document identity and its confined absolute path. */
export interface DocumentTarget {
  /** Session-scoped identity carried by the command and the edit fact. */
  readonly documentId: DocumentIdValue
  /** Path as it appeared in the identity, relative or absolute. */
  readonly path: string
  /** Absolute path inside the workspace root. */
  readonly absolutePath: string
}

/**
 * Split a model-supplied document identity and resolve its path.
 * @param documentId - identity returned by a read or the open-document set.
 * @param sessionId - Session that must own the identity.
 * @param workspaceRoot - root the path is resolved against.
 * @returns the identity, its recorded path, and the absolute path.
 * @throws {OfficeToolResultError} `document-missing` when the identity belongs to another Session or names no path.
 */
export function documentTarget(documentId: string, sessionId: SessionId, workspaceRoot: string): DocumentTarget {
  const prefix = `${sessionId}:`
  if (!documentId.startsWith(prefix)) {
    throw new OfficeToolResultError(
      'document-missing',
      `Document id "${documentId}" does not belong to this Session; use the id returned by a document read.`,
    )
  }
  const path = documentId.slice(prefix.length)
  if (path.length === 0) {
    throw new OfficeToolResultError('document-missing', `Document id "${documentId}" names no path.`)
  }
  return { documentId: DocumentId(documentId), path, absolutePath: resolve(workspaceRoot, path) }
}

/**
 * Require one absolute path to resolve inside the Session's workspace root.
 * @param absolutePath - resolved target path.
 * @param workspaceRoot - the single permitted root.
 * @returns the real absolute path.
 * @throws {OfficeExchangeError} `out-of-scope` when the path leaves the root.
 */
export async function confinePath(absolutePath: string, workspaceRoot: string): Promise<string> {
  return await assertWithinRoots(absolutePath, [workspaceRoot])
}

/**
 * Resolve a plain tool path (for conversion) inside the workspace root.
 * @param path - absolute or workspace-relative path.
 * @param workspaceRoot - root the path is resolved against.
 * @returns the confined absolute path.
 * @throws {OfficeExchangeError} `out-of-scope` when the path leaves the root.
 */
export async function confineToolPath(path: string, workspaceRoot: string): Promise<string> {
  return await confinePath(resolve(workspaceRoot, path), workspaceRoot)
}
