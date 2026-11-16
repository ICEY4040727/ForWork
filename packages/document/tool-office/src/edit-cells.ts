/**
 * `document_edit_cells`: replace spreadsheet cell values or formulas in the
 * live unit, then persist the accepted result with one guarded save.
 * @module @forwork/tool-office/edit-cells
 */

import type { DocumentEditStep } from '@forwork/session'
import type { JsonValue } from '@forwork/util-values'
import { priorCell, runEdit, sheetCells } from './edit.ts'
import type { EditOutput, EditPlan } from './edit.ts'
import { failure } from './errors.ts'
import type { OfficeToolFailure, OfficeToolResult } from './errors.ts'
import type { DocumentToolCall } from './runtime.ts'

/** One requested cell replacement. */
export interface CellEditInput {
  /** Sheet identity inside the workbook snapshot. */
  readonly sheetId: string
  /** Zero-based row index. */
  readonly row: number
  /** Zero-based column index. */
  readonly column: number
  /** New JSON cell value; omit when assigning a formula. */
  readonly value?: JsonValue
  /** New formula text; omit when assigning a value. */
  readonly formula?: string
}

/** Input of one `document_edit_cells` call. */
export interface CellEditRequest {
  /** Session-scoped document identity. */
  readonly documentId: string
  /** Content version the caller believes the live unit holds. */
  readonly fromVersion: number
  /** Cell replacements, applied in order. */
  readonly operations: readonly CellEditInput[]
}

/**
 * Replace cell values or formulas in the live document.
 * @param call - runtime, Session, signal, and caps.
 * @param request - document identity, expected version, and cell replacements.
 * @returns the new content version and applied replacements, or a refusal.
 */
export async function editCells(
  call: DocumentToolCall,
  request: CellEditRequest,
): Promise<OfficeToolResult<EditOutput<CellEditInput>>> {
  const steps: DocumentEditStep[] = []
  for (const operation of request.operations) {
    const step = stepOf(operation)
    if ('status' in step) return step
    steps.push(step)
  }
  const plan: EditPlan<CellEditInput> = {
    steps,
    applied: request.operations,
    inverse(snapshot) {
      return request.operations.flatMap((operation): DocumentEditStep[] => {
        const cells = sheetCells(snapshot, operation.sheetId)
        const prior = priorCell(cells, operation.row, operation.column)
        if (operation.formula !== undefined) {
          return [{
            kind: 'formula',
            sheetId: operation.sheetId,
            row: operation.row,
            column: operation.column,
            formula: prior.formula ?? null,
          }]
        }
        const value: DocumentEditStep = {
          kind: 'cell',
          sheetId: operation.sheetId,
          row: operation.row,
          column: operation.column,
          value: prior.value,
        }
        return prior.formula === undefined ? [value] : [value, {
          kind: 'formula',
          sheetId: operation.sheetId,
          row: operation.row,
          column: operation.column,
          formula: prior.formula,
        }]
      })
    },
  }
  return await runEdit(call, request.documentId, request.fromVersion, plan)
}

/** Validate one requested replacement and project it to a command step. */
function stepOf(operation: CellEditInput): DocumentEditStep | OfficeToolFailure {
  if (operation.sheetId.length === 0) {
    return failure('invalid-request', 'Every cell operation requires a sheetId.')
  }
  if (!index(operation.row) || !index(operation.column)) {
    return failure('invalid-request', `Cell ${String(operation.row)}:${String(operation.column)} has an invalid index.`)
  }
  if ((operation.value === undefined) === (operation.formula === undefined)) {
    return failure('invalid-request', `Cell ${operation.row}:${operation.column} must set exactly one of value or formula.`)
  }
  if (operation.formula !== undefined) {
    if (operation.formula.length === 0) {
      return failure('invalid-request', `Cell ${operation.row}:${operation.column} has an empty formula.`)
    }
    return {
      kind: 'formula',
      sheetId: operation.sheetId,
      row: operation.row,
      column: operation.column,
      formula: operation.formula,
    }
  }
  return {
    kind: 'cell',
    sheetId: operation.sheetId,
    row: operation.row,
    column: operation.column,
    value: operation.value ?? null,
  }
}

/** Whether a value is a non-negative safe index. */
function index(value: number): boolean {
  return Number.isSafeInteger(value) && value >= 0
}
