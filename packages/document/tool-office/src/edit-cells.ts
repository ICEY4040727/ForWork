/**
 * `document_edit_cells`: replace spreadsheet cell values or formulas, insert or
 * delete rows or columns, or style a range in the live unit, then persist the
 * accepted result with one guarded save.
 * @module @forwork/tool-office/edit-cells
 */

import type { OfficeDocumentSnapshot } from '@forwork/office-exchange'
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

/** One requested worksheet row insertion or deletion. */
export interface RowEditInput {
  /** Marks a row insertion or deletion. */
  readonly kind: 'row'
  /** Sheet identity inside the workbook snapshot. */
  readonly sheetId: string
  /** Zero-based index of the first affected row. */
  readonly row: number
  /** Number of rows to insert or delete. */
  readonly count: number
  /** Whether the rows are inserted before `row` or deleted from `row`. */
  readonly action: 'insert' | 'delete'
}

/** One requested worksheet column insertion or deletion. */
export interface ColumnEditInput {
  /** Marks a column insertion or deletion. */
  readonly kind: 'column'
  /** Sheet identity inside the workbook snapshot. */
  readonly sheetId: string
  /** Zero-based index of the first affected column. */
  readonly column: number
  /** Number of columns to insert or delete. */
  readonly count: number
  /** Whether the columns are inserted before `column` or deleted from `column`. */
  readonly action: 'insert' | 'delete'
}

/** One requested style assignment over a rectangular worksheet range. */
export interface StyleEditInput {
  /** Marks a range style assignment. */
  readonly kind: 'style'
  /** Sheet identity inside the workbook snapshot. */
  readonly sheetId: string
  /** Zero-based first row of the range, inclusive. */
  readonly startRow: number
  /** Zero-based first column of the range, inclusive. */
  readonly startColumn: number
  /** Zero-based last row of the range, inclusive. */
  readonly endRow: number
  /** Zero-based last column of the range, inclusive. */
  readonly endColumn: number
  /** JSON style object applied to every cell in the range; `null` clears the style. */
  readonly style: JsonValue
}

/** One requested spreadsheet edit step. */
export type SheetEditInput = CellEditInput | RowEditInput | ColumnEditInput | StyleEditInput

/** Input of one `document_edit_cells` call. */
export interface CellEditRequest {
  /** Session-scoped document identity. */
  readonly documentId: string
  /** Content version the caller believes the live unit holds. */
  readonly fromVersion: number
  /** Sheet edits, applied in order. */
  readonly operations: readonly SheetEditInput[]
}

/**
 * Replace spreadsheet content, structure, or styles in the live document.
 * @param call - runtime, Session, signal, and caps.
 * @param request - document identity, expected version, and sheet edits.
 * @returns the new content version and applied edits, or a refusal.
 */
export async function editCells(
  call: DocumentToolCall,
  request: CellEditRequest,
): Promise<OfficeToolResult<EditOutput<SheetEditInput>>> {
  const steps: DocumentEditStep[] = []
  for (const operation of request.operations) {
    const step = stepOf(operation)
    if ('status' in step) return step
    steps.push(step)
  }
  const plan: EditPlan<SheetEditInput> = {
    steps,
    applied: request.operations,
    inverse(snapshot) {
      return request.operations.flatMap((operation): DocumentEditStep[] => (
        'kind' in operation ? inverseStep(snapshot, operation) : inverseCell(snapshot, operation)
      ))
    },
  }
  return await runEdit(call, request.documentId, request.fromVersion, plan)
}

/** Build the inverse of one cell value or formula replacement. */
function inverseCell(
  snapshot: OfficeDocumentSnapshot,
  operation: CellEditInput,
): DocumentEditStep[] {
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
}

/** Build the inverse of one row, column, or style step from the pre-edit snapshot. */
function inverseStep(
  snapshot: OfficeDocumentSnapshot,
  operation: RowEditInput | ColumnEditInput | StyleEditInput,
): DocumentEditStep[] {
  switch (operation.kind) {
    case 'row':
      return inverseRow(sheetCells(snapshot, operation.sheetId), operation)
    case 'column':
      return inverseColumn(sheetCells(snapshot, operation.sheetId), operation)
    case 'style':
      return inverseStyle(sheetCells(snapshot, operation.sheetId), operation)
  }
}

/**
 * Reverse one row edit. A deletion inserts blank rows first, then restores the
 * cells that occupied the deleted rows in the pre-edit snapshot.
 */
function inverseRow(
  cells: Record<string, JsonValue>,
  operation: RowEditInput,
): DocumentEditStep[] {
  const structural: DocumentEditStep = {
    kind: 'row',
    sheetId: operation.sheetId,
    row: operation.row,
    count: operation.count,
    action: operation.action === 'insert' ? 'delete' : 'insert',
  }
  if (operation.action === 'insert') return [structural]
  return [structural, ...restoreDeletedCells(cells, operation.sheetId, 'row', {
    start: operation.row,
    count: operation.count,
  })]
}

/**
 * Reverse one column edit. A deletion inserts blank columns first, then restores
 * the cells that occupied the deleted columns in the pre-edit snapshot.
 */
function inverseColumn(
  cells: Record<string, JsonValue>,
  operation: ColumnEditInput,
): DocumentEditStep[] {
  const structural: DocumentEditStep = {
    kind: 'column',
    sheetId: operation.sheetId,
    column: operation.column,
    count: operation.count,
    action: operation.action === 'insert' ? 'delete' : 'insert',
  }
  if (operation.action === 'insert') return [structural]
  return [structural, ...restoreDeletedCells(cells, operation.sheetId, 'column', {
    start: operation.column,
    count: operation.count,
  })]
}

/**
 * Restore the cells that a deleted row or column span removed, reading the
 * pre-edit snapshot.
 * @param cells - recorded sheet cells before the deletion.
 * @param sheetId - sheet identity inside the workbook snapshot.
 * @param axis - whether the deletion removed rows or columns.
 * @param span - first deleted row or column and the number removed.
 * @returns steps that restore the value, formula, and style of every recorded cell inside the span.
 */
function restoreDeletedCells(
  cells: Record<string, JsonValue>,
  sheetId: string,
  axis: 'row' | 'column',
  span: { readonly start: number; readonly count: number },
): DocumentEditStep[] {
  const restored: DocumentEditStep[] = []
  for (const [rowKey, rowValue] of Object.entries(cells)) {
    const row = Number(rowKey)
    if (!Number.isSafeInteger(row)) continue
    const rowCells = objectOf(rowValue)
    if (rowCells === undefined) continue
    for (const [columnKey, cellValue] of Object.entries(rowCells)) {
      const column = Number(columnKey)
      if (!Number.isSafeInteger(column)) continue
      const position = axis === 'row' ? row : column
      if (position < span.start || position >= span.start + span.count) continue
      restored.push(...restoreCell(sheetId, row, column, objectOf(cellValue)))
    }
  }
  return restored
}

/**
 * Reverse one range style by clearing the range first, then restoring each
 * previously styled cell's own style from the pre-edit snapshot.
 */
function inverseStyle(
  cells: Record<string, JsonValue>,
  operation: StyleEditInput,
): DocumentEditStep[] {
  const steps: DocumentEditStep[] = [{
    kind: 'style',
    sheetId: operation.sheetId,
    startRow: operation.startRow,
    startColumn: operation.startColumn,
    endRow: operation.endRow,
    endColumn: operation.endColumn,
    style: null,
  }]
  for (const [rowKey, rowValue] of Object.entries(cells)) {
    const row = Number(rowKey)
    if (!Number.isSafeInteger(row) || row < operation.startRow || row > operation.endRow) continue
    const rowCells = objectOf(rowValue)
    if (rowCells === undefined) continue
    for (const [columnKey, cellValue] of Object.entries(rowCells)) {
      const column = Number(columnKey)
      if (!Number.isSafeInteger(column) || column < operation.startColumn || column > operation.endColumn) continue
      const style = objectOf(cellValue)?.['s']
      if (style === undefined) continue
      steps.push({
        kind: 'style',
        sheetId: operation.sheetId,
        startRow: row,
        startColumn: column,
        endRow: row,
        endColumn: column,
        style,
      })
    }
  }
  return steps
}

/** Build the steps that restore one cell's prior value, formula, and style. */
function restoreCell(
  sheetId: string,
  row: number,
  column: number,
  cell: Record<string, JsonValue> | undefined,
): DocumentEditStep[] {
  if (cell === undefined) return []
  const steps: DocumentEditStep[] = []
  if (cell['v'] !== undefined || cell['f'] !== undefined) {
    steps.push({ kind: 'cell', sheetId, row, column, value: cell['v'] ?? null })
  }
  if (typeof cell['f'] === 'string') {
    steps.push({ kind: 'formula', sheetId, row, column, formula: cell['f'] })
  }
  const style = cell['s']
  if (style !== undefined) {
    steps.push({ kind: 'style', sheetId, startRow: row, startColumn: column, endRow: row, endColumn: column, style })
  }
  return steps
}

/** Validate one requested edit and project it to a command step. */
function stepOf(operation: SheetEditInput): DocumentEditStep | OfficeToolFailure {
  if ('kind' in operation) {
    switch (operation.kind) {
      case 'row':
        return rowStepOf(operation)
      case 'column':
        return columnStepOf(operation)
      case 'style':
        return styleStepOf(operation)
    }
  }
  return cellStepOf(operation)
}

/** Validate one cell value or formula replacement. */
function cellStepOf(operation: CellEditInput): DocumentEditStep | OfficeToolFailure {
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

/** Validate one row insertion or deletion. */
function rowStepOf(operation: RowEditInput): DocumentEditStep | OfficeToolFailure {
  if (operation.sheetId.length === 0) {
    return failure('invalid-request', 'Every row operation requires a sheetId.')
  }
  if (!index(operation.row)) {
    return failure('invalid-request', `Row ${String(operation.row)} has an invalid index.`)
  }
  if (!positive(operation.count)) {
    return failure('invalid-request', `Row ${operation.row} has an invalid count ${String(operation.count)}.`)
  }
  return {
    kind: 'row',
    sheetId: operation.sheetId,
    row: operation.row,
    count: operation.count,
    action: operation.action,
  }
}

/** Validate one column insertion or deletion. */
function columnStepOf(operation: ColumnEditInput): DocumentEditStep | OfficeToolFailure {
  if (operation.sheetId.length === 0) {
    return failure('invalid-request', 'Every column operation requires a sheetId.')
  }
  if (!index(operation.column)) {
    return failure('invalid-request', `Column ${String(operation.column)} has an invalid index.`)
  }
  if (!positive(operation.count)) {
    return failure('invalid-request', `Column ${operation.column} has an invalid count ${String(operation.count)}.`)
  }
  return {
    kind: 'column',
    sheetId: operation.sheetId,
    column: operation.column,
    count: operation.count,
    action: operation.action,
  }
}

/** Validate one rectangular range style assignment. */
function styleStepOf(operation: StyleEditInput): DocumentEditStep | OfficeToolFailure {
  const range = `${String(operation.startRow)}:${String(operation.startColumn)}-${String(operation.endRow)}:${String(operation.endColumn)}`
  if (operation.sheetId.length === 0) {
    return failure('invalid-request', 'Every style operation requires a sheetId.')
  }
  if (
    !index(operation.startRow)
    || !index(operation.startColumn)
    || !index(operation.endRow)
    || !index(operation.endColumn)
  ) {
    return failure('invalid-request', `Style range ${range} has an invalid index.`)
  }
  if (operation.startRow > operation.endRow || operation.startColumn > operation.endColumn) {
    return failure('invalid-request', `Style range ${range} is reversed.`)
  }
  if (operation.style !== null && (typeof operation.style !== 'object' || Array.isArray(operation.style))) {
    return failure('invalid-request', `Style range ${range} requires a JSON object or null style.`)
  }
  return {
    kind: 'style',
    sheetId: operation.sheetId,
    startRow: operation.startRow,
    startColumn: operation.startColumn,
    endRow: operation.endRow,
    endColumn: operation.endColumn,
    style: operation.style,
  }
}

/** Whether a value is a non-negative safe index. */
function index(value: number): boolean {
  return Number.isSafeInteger(value) && value >= 0
}

/** Whether a value is a positive safe count. */
function positive(value: number): boolean {
  return Number.isSafeInteger(value) && value > 0
}

/** Narrow a JSON value to a plain object. */
function objectOf(value: JsonValue | undefined): Record<string, JsonValue> | undefined {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return undefined
  return value
}
