/**
 * `document_read`: read the live unit's structure through the command bridge.
 * A read never saves, never advances the content version, and refuses when no
 * live editor instance matches the document.
 * @module @forwork/tool-office/read
 */

import { assertNever } from '@forwork/util-values'
import { confinePath, documentTarget } from './documents.ts'
import { failure, guarded } from './errors.ts'
import type { OfficeToolResult } from './errors.ts'
import type { DocumentToolCall } from './runtime.ts'
import { summarizeDocument } from './summary.ts'
import type { OfficeDocumentSummary } from './summary.ts'

/** Accepted result of one structure read. */
export interface DocumentReadOutput {
  /** Session-scoped identity the read used. */
  readonly documentId: string
  /** Live document content version at read time. */
  readonly contentVersion: number
  /** Structure summary with a bounded cell preview. */
  readonly summary: OfficeDocumentSummary
}

/**
 * Read one open document's structure from the live editor.
 * @param call - runtime, Session, signal, and caps.
 * @param documentId - Session-scoped document identity.
 * @returns the structure summary or a structured refusal.
 */
export async function readDocument(
  call: DocumentToolCall,
  documentId: string,
): Promise<OfficeToolResult<DocumentReadOutput>> {
  return await guarded(async () => {
    const target = documentTarget(documentId, call.session.id, call.scope.workspaceRoot)
    await confinePath(target.absolutePath, call.scope.workspaceRoot)
    const outcome = await call.runtime.commands.execute({
      session: call.session,
      scope: call.scope,
      documentId: target.documentId,
      operation: { kind: 'read' },
      fromVersion: call.runtime.commands.currentVersion(target.documentId),
    }, call.signal)
    switch (outcome.status) {
      case 'read': {
        const summary = summarizeDocument(outcome.snapshot, call.config.summaryCells)
        call.runtime.commands.remember(target.documentId, outcome.contentVersion)
        return { status: 'ok', documentId: target.documentId, contentVersion: outcome.contentVersion, summary }
      }
      case 'document-not-open':
        return failure('document-not-open', `No live editor instance has "${documentId}" open.`)
      case 'timeout':
        return failure('timeout', `The editor did not answer the read of "${documentId}" in time.`)
      case 'pending-edits':
        return failure('pending-edits', `"${documentId}" has unsaved user edits; save them before the agent reads it.`)
      case 'conflict':
        return failure('conflict', `"${documentId}" changed; read it again.`, { currentVersion: outcome.currentVersion })
      case 'unsupported-feature':
        return failure('unsupported-feature', `"${documentId}" contains features the editor cannot preserve.`, {
          features: outcome.features,
        })
      case 'applied':
        return failure('conversion-failed', 'The editor returned an edit result for a read command.')
      /* v8 ignore next 2 -- the command outcome union is closed; every member is handled above. */
      default:
        return assertNever(outcome)
    }
  })
}
