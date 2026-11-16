/**
 * Shared edit path: resolve the document, read the current file snapshot,
 * compute the inverse, dispatch one command to the live editor, persist the
 * returned snapshot with one guarded save, and append one `document/edit` fact.
 *
 * A refused or unanswered command writes nothing and appends no edit fact.
 * @module @forwork/tool-office/edit
 */

import type { OfficeCommandOutcome, OfficeDocumentReadResult, OfficeDocumentSnapshot } from '@forwork/office-exchange'
import type { DocumentEditStep } from '@forwork/session'
import { assertNever } from '@forwork/util-values'
import type { JsonValue } from '@forwork/util-values'
import { confinePath, documentTarget } from './documents.ts'
import type { DocumentTarget } from './documents.ts'
import { failure, guarded, OfficeToolResultError } from './errors.ts'
import type { OfficeToolResult } from './errors.ts'
import type { DocumentToolCall } from './runtime.ts'

/** Accepted result of one edit tool call. */
export interface EditOutput<T> {
  /** Session-scoped identity of the edited document. */
  readonly documentId: string
  /** Content version after the accepted save. */
  readonly contentVersion: number
  /** Operations the editor applied, in request order. */
  readonly applied: readonly T[]
}

/** Inverse computation over the file snapshot read before the command. */
export interface EditPlan<T> {
  /** Steps dispatched to the live editor. */
  readonly steps: readonly DocumentEditStep[]
  /** Operations reported as applied. */
  readonly applied: readonly T[]
  /**
   * Build the exact inverse from the pre-edit snapshot.
   * @param snapshot - file snapshot read before the command.
   * @returns steps that restore the prior content.
   */
  inverse(snapshot: OfficeDocumentSnapshot): readonly DocumentEditStep[]
}

/**
 * Require one workbook sheet's cell data.
 * @param snapshot - pre-edit document snapshot.
 * @param sheetId - sheet identity inside the workbook snapshot.
 * @returns the sheet's cell data keyed by row and column.
 * @throws {OfficeToolResultError} `invalid-request` when the snapshot is not a workbook or the sheet is absent.
 */
export function sheetCells(
  snapshot: OfficeDocumentSnapshot,
  sheetId: string,
): Record<string, JsonValue> {
  if (snapshot.kind !== 'workbook') {
    throw new OfficeToolResultError('invalid-request', 'Cell edits require an .xlsx workbook.')
  }
  const workbook = objectOf(snapshot.data)
  const sheets = objectOf(workbook?.['sheets'])
  const sheet = objectOf(sheets?.[sheetId])
  const cells = objectOf(sheet?.['cellData'])
  if (cells === undefined) {
    throw new OfficeToolResultError('invalid-request', `The workbook has no sheet "${sheetId}".`)
  }
  return cells
}

/** One cell's prior value and formula. */
export interface PriorCell {
  /** JSON value before the edit, or `null` for an empty cell. */
  readonly value: JsonValue
  /** Formula before the edit, when the cell held one. */
  readonly formula?: string
}

/**
 * Read one cell from the pre-edit snapshot.
 * @param cells - sheet cell data.
 * @param row - zero-based row index.
 * @param column - zero-based column index.
 * @returns the prior value and formula.
 */
export function priorCell(cells: Record<string, JsonValue>, row: number, column: number): PriorCell {
  const cell = objectOf(objectOf(cells[String(row)])?.[String(column)])
  const formula = cell?.['f']
  return {
    value: cell?.['v'] ?? null,
    ...typeof formula === 'string' ? { formula } : {},
  }
}

/**
 * Read one paragraph's text from the pre-edit snapshot.
 * @param snapshot - document snapshot read before the command.
 * @param index - zero-based paragraph index.
 * @returns the text that paragraph holds.
 * @throws {OfficeToolResultError} `invalid-request` when the paragraph does not exist.
 */
export function priorParagraph(snapshot: OfficeDocumentSnapshot, index: number): string {
  if (snapshot.kind !== 'document') {
    throw new OfficeToolResultError('invalid-request', 'Paragraph edits require a .docx document.')
  }
  const body = objectOf(objectOf(snapshot.data)?.['body'])
  const stream = body?.['dataStream']
  const declared = body?.['paragraphs']
  if (typeof stream !== 'string' || !Array.isArray(declared)) {
    throw new OfficeToolResultError('invalid-request', 'The document snapshot has no paragraph structure.')
  }
  const paragraph = index < declared.length ? objectOf(declared[index]) : undefined
  const start = paragraph?.['startIndex']
  const nextEntry = index + 1 < declared.length ? objectOf(declared[index + 1]) : undefined
  const next = index + 1 < declared.length ? nextEntry?.['startIndex'] : stream.length
  if (typeof start !== 'number' || typeof next !== 'number' || start < 0 || next < start || next > stream.length) {
    throw new OfficeToolResultError('invalid-request', `The document has no paragraph ${index}.`)
  }
  return stream.slice(start, next).replace(/[\r\n]+$/u, '')
}

/**
 * Run one edit through the live editor and persist its accepted result.
 * @param call - runtime, Session, signal, and caps.
 * @param documentId - Session-scoped document identity.
 * @param fromVersion - content version the caller believes the live unit holds.
 * @param plan - dispatched steps, applied operations, and inverse computation.
 * @returns the new content version and applied operations, or a structured refusal.
 */
export async function runEdit<T>(
  call: DocumentToolCall,
  documentId: string,
  fromVersion: number,
  plan: EditPlan<T>,
): Promise<OfficeToolResult<EditOutput<T>>> {
  return await guarded(async () => {
    if (plan.steps.length === 0) {
      return failure('invalid-request', 'An edit requires at least one operation.')
    }
    if (plan.steps.length > call.config.maxEditSteps) {
      return failure('bound-exceeded', `An edit accepts at most ${call.config.maxEditSteps} operations.`)
    }
    if (!Number.isSafeInteger(fromVersion) || fromVersion < 0) {
      return failure('invalid-request', `fromVersion ${String(fromVersion)} is not a non-negative integer.`)
    }
    const target = documentTarget(documentId, call.session.id, call.scope.workspaceRoot)
    await confinePath(target.absolutePath, call.scope.workspaceRoot)
    const read = await call.runtime.documents.read(call.scope, target.absolutePath, call.signal)
    if (read.unsupported.length > 0) {
      return failure('unsupported-feature', `"${documentId}" contains features the editor cannot preserve.`, {
        features: read.unsupported,
      })
    }
    const inverse = plan.inverse(read.snapshot)
    const outcome = await call.runtime.commands.execute({
      session: call.session,
      scope: call.scope,
      documentId: target.documentId,
      operation: { kind: 'edit', steps: plan.steps },
      fromVersion,
    }, call.signal)
    return await settleEdit(call, target, read, fromVersion, plan, inverse, outcome)
  })
}

/** Persist one applied command and append its edit fact. */
async function settleEdit<T>(
  call: DocumentToolCall,
  target: DocumentTarget,
  read: OfficeDocumentReadResult,
  fromVersion: number,
  plan: EditPlan<T>,
  inverse: readonly DocumentEditStep[],
  outcome: OfficeCommandOutcome,
): Promise<OfficeToolResult<EditOutput<T>>> {
  switch (outcome.status) {
    case 'applied': {
      const toVersion = fromVersion + 1
      if (outcome.contentVersion !== toVersion) {
        return failure(
          'conversion-failed',
          `The editor applied the edit at version ${fromVersion} but reported ${outcome.contentVersion}.`,
        )
      }
      await call.runtime.documents.save(
        call.scope,
        target.absolutePath,
        outcome.snapshot,
        read.format,
        read.version,
        call.signal,
      )
      call.session.append('document/edit', {
        documentId: target.documentId,
        actor: 'agent',
        fromVersion,
        toVersion,
        operation: plan.steps,
        inverse,
      })
      call.runtime.commands.remember(target.documentId, toVersion)
      return { status: 'ok', documentId: target.documentId, contentVersion: toVersion, applied: plan.applied }
    }
    case 'conflict':
      call.runtime.commands.remember(target.documentId, outcome.currentVersion)
      return failure('conflict', `"${target.documentId}" changed; read it again.`, {
        currentVersion: outcome.currentVersion,
      })
    case 'pending-edits':
      return failure('pending-edits', `"${target.documentId}" has unsaved user edits; save them first.`)
    case 'unsupported-feature':
      return failure('unsupported-feature', `"${target.documentId}" contains features the editor cannot preserve.`, {
        features: outcome.features,
      })
    case 'document-not-open':
      return failure('document-not-open', `No live editor instance has "${target.documentId}" open.`)
    case 'timeout':
      return failure('timeout', `The editor did not answer the edit of "${target.documentId}" in time.`)
    case 'read':
      return failure('conversion-failed', 'The editor returned a read result for an edit command.')
    /* v8 ignore next 2 -- the command outcome union is closed; every member is handled above. */
    default:
      return assertNever(outcome)
  }
}

/** Narrow a JSON value to a plain object. */
function objectOf(value: JsonValue | undefined): Record<string, JsonValue> | undefined {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return undefined
  return value
}
