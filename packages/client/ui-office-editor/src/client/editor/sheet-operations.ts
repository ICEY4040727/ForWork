/**
 * Sheet edit-step translation for the Univer sheets runtime.
 *
 * Every step is validated and translated into one facade command before the
 * first command runs. A step that cannot be expressed refuses the whole
 * operation with the workbook still untouched.
 */
import type { ICellData } from '@univerjs/core'
import type { DocumentEditOperation } from '@forwork/session/src/document-events.ts'
import { assertNever } from '@forwork/util-values'
import type { JsonValue } from '@forwork/util-values'
import type { OfficeApplyOutcome } from '../command-bridge.ts'

/** The style fields an Univer cell-data update accepts. */
type CellStyle = Extract<NonNullable<ICellData['s']>, object>

/** One rectangular sheet range the executor writes. */
export interface SheetRange {
  /**
   * Replace cell content or merge style properties across the range.
   * @param value - a scalar, a content replacement, or a style update.
   */
  setValue(value: string | number | boolean | ICellData): void
  /**
   * Assign one formula across the range.
   * @param formula - formula text in A1 notation.
   */
  setFormula(formula: string): void
}

/** One worksheet the executor can resize and index. */
export interface SheetTarget {
  /** @returns the worksheet's row capacity. */
  getMaxRows(): number
  /** @returns the worksheet's column capacity. */
  getMaxColumns(): number
  /**
   * Insert blank rows before `row`.
   * @param row - zero-based first new row index.
   * @param count - positive number of rows to insert.
   */
  insertRows(row: number, count: number): void
  /**
   * Delete rows starting at `row`.
   * @param row - zero-based first deleted row index.
   * @param count - positive number of rows to delete.
   */
  deleteRows(row: number, count: number): void
  /**
   * Insert blank columns before `column`.
   * @param column - zero-based first new column index.
   * @param count - positive number of columns to insert.
   */
  insertColumns(column: number, count: number): void
  /**
   * Delete columns starting at `column`.
   * @param column - zero-based first deleted column index.
   * @param count - positive number of columns to delete.
   */
  deleteColumns(column: number, count: number): void
  /**
   * Resolve one inclusive rectangular range.
   * @param startRow - zero-based first row.
   * @param startColumn - zero-based first column.
   * @param endRow - zero-based last row.
   * @param endColumn - zero-based last column.
   * @returns the range the executor writes.
   */
  getRange(startRow: number, startColumn: number, endRow: number, endColumn: number): SheetRange
}

/** The workbook surface the executor reads. */
export interface SheetWorkbook {
  /**
   * Resolve one worksheet by identity.
   * @param sheetId - sheet identity inside the workbook snapshot.
   * @returns the worksheet, or `null` when the workbook has no such sheet.
   */
  getSheetBySheetId(sheetId: string): SheetTarget | null
}

/** One validated command in the order the operation declared it. */
type SheetCommand =
  | { readonly kind: 'value'; readonly range: SheetRange; readonly value: string | number | boolean | null }
  | { readonly kind: 'formula'; readonly range: SheetRange; readonly formula: string | null }
  | {
    readonly kind: 'row'
    readonly sheet: SheetTarget
    readonly row: number
    readonly count: number
    readonly action: 'insert' | 'delete'
  }
  | {
    readonly kind: 'column'
    readonly sheet: SheetTarget
    readonly column: number
    readonly count: number
    readonly action: 'insert' | 'delete'
  }
  | { readonly kind: 'style'; readonly range: SheetRange; readonly style: CellStyle | null }

/** Validation of one step: its command, or the detail naming what is unsupported. */
type Planned = { readonly command: SheetCommand } | { readonly detail: string }

/**
 * Apply one document edit to a live workbook.
 *
 * All steps are validated and translated first; an unsupported step returns
 * `unsupported-feature` without running any command.
 * @param workbook - the live workbook facade.
 * @param operation - ordered steps from the Host command.
 * @returns whether every step applied, or the content that stops it.
 */
export function applySheetOperation(
  workbook: SheetWorkbook,
  operation: DocumentEditOperation,
): OfficeApplyOutcome {
  const commands: SheetCommand[] = []
  for (const step of operation) {
    const planned = plan(workbook, step)
    if ('detail' in planned) {
      return { status: 'unsupported-feature', features: [{ kind: 'other', detail: planned.detail }] }
    }
    commands.push(planned.command)
  }
  for (const command of commands) execute(command)
  return { status: 'applied' }
}

/** Run one validated command. */
function execute(command: SheetCommand): void {
  switch (command.kind) {
    case 'value':
      if (command.value === null) command.range.setValue({ v: null, f: null, p: null })
      else command.range.setValue(command.value)
      return
    case 'formula':
      if (command.formula === null) command.range.setValue({ v: null, f: null, p: null })
      else command.range.setFormula(command.formula)
      return
    case 'row':
      if (command.action === 'insert') command.sheet.insertRows(command.row, command.count)
      else command.sheet.deleteRows(command.row, command.count)
      return
    case 'column':
      if (command.action === 'insert') command.sheet.insertColumns(command.column, command.count)
      else command.sheet.deleteColumns(command.column, command.count)
      return
    case 'style':
      command.range.setValue({ s: command.style })
      return
    default:
      return assertNever(command)
  }
}

/**
 * Validate one step and translate it into its command.
 * @param workbook - the live workbook facade.
 * @param step - one declared edit step.
 * @returns the command, or the detail naming what is unsupported.
 */
function plan(workbook: SheetWorkbook, step: DocumentEditOperation[number]): Planned {
  if (step.kind === 'paragraph') return { detail: 'paragraph' }
  const sheet = workbook.getSheetBySheetId(step.sheetId)
  if (sheet === null) return { detail: `sheet ${step.sheetId}` }
  switch (step.kind) {
    case 'cell':
      return planCell(sheet, step)
    case 'formula':
      return planFormula(sheet, step)
    case 'row':
      return planRow(sheet, step)
    case 'column':
      return planColumn(sheet, step)
    case 'style':
      return planStyle(sheet, step)
    default:
      return assertNever(step)
  }
}

/** Validate one cell replacement. */
function planCell(
  sheet: SheetTarget,
  step: Extract<DocumentEditOperation[number], { readonly kind: 'cell' }>,
): Planned {
  const { value } = step
  if (value !== null && typeof value !== 'string' && typeof value !== 'number' && typeof value !== 'boolean') {
    return { detail: `cell value ${Array.isArray(value) ? 'array' : typeof value}` }
  }
  if (!contains(sheet, step.row, step.column)) return { detail: `cell ${step.row}:${step.column}` }
  return { command: { kind: 'value', range: cellRange(sheet, step.row, step.column), value } }
}

/** Validate one formula assignment. */
function planFormula(
  sheet: SheetTarget,
  step: Extract<DocumentEditOperation[number], { readonly kind: 'formula' }>,
): Planned {
  if (step.formula !== null && typeof step.formula !== 'string') {
    return { detail: `formula ${typeof step.formula}` }
  }
  if (!contains(sheet, step.row, step.column)) return { detail: `cell ${step.row}:${step.column}` }
  return {
    command: { kind: 'formula', range: cellRange(sheet, step.row, step.column), formula: step.formula },
  }
}

/** Validate one row insertion or deletion. */
function planRow(
  sheet: SheetTarget,
  step: Extract<DocumentEditOperation[number], { readonly kind: 'row' }>,
): Planned {
  if (!positive(step.count) || !index(step.row)) return { detail: `row ${step.row} count ${step.count}` }
  const maxRows = sheet.getMaxRows()
  const fits = step.action === 'insert' ? step.row < maxRows : step.row + step.count <= maxRows
  if (!fits) return { detail: `row ${step.row} count ${step.count} of ${maxRows}` }
  return { command: { kind: 'row', sheet, row: step.row, count: step.count, action: step.action } }
}

/** Validate one column insertion or deletion. */
function planColumn(
  sheet: SheetTarget,
  step: Extract<DocumentEditOperation[number], { readonly kind: 'column' }>,
): Planned {
  if (!positive(step.count) || !index(step.column)) {
    return { detail: `column ${step.column} count ${step.count}` }
  }
  const maxColumns = sheet.getMaxColumns()
  const fits = step.action === 'insert' ? step.column < maxColumns : step.column + step.count <= maxColumns
  if (!fits) return { detail: `column ${step.column} count ${step.count} of ${maxColumns}` }
  return { command: { kind: 'column', sheet, column: step.column, count: step.count, action: step.action } }
}

/** Validate one range style assignment. */
function planStyle(
  sheet: SheetTarget,
  step: Extract<DocumentEditOperation[number], { readonly kind: 'style' }>,
): Planned {
  if (!index(step.startRow) || !index(step.startColumn) || !index(step.endRow) || !index(step.endColumn)) {
    return { detail: `style range ${step.startRow}:${step.startColumn}-${step.endRow}:${step.endColumn}` }
  }
  if (step.startRow > step.endRow || step.startColumn > step.endColumn) {
    return { detail: `style range ${step.startRow}:${step.startColumn}-${step.endRow}:${step.endColumn}` }
  }
  if (step.endRow >= sheet.getMaxRows() || step.endColumn >= sheet.getMaxColumns()) {
    return { detail: `style range ${step.startRow}:${step.startColumn}-${step.endRow}:${step.endColumn}` }
  }
  const style = cellStyle(step.style)
  if (style === undefined) {
    return { detail: `style ${Array.isArray(step.style) ? 'array' : typeof step.style}` }
  }
  return {
    command: {
      kind: 'style',
      range: sheet.getRange(step.startRow, step.startColumn, step.endRow, step.endColumn),
      style,
    },
  }
}

/** Resolve one single-cell range. */
function cellRange(sheet: SheetTarget, row: number, column: number): SheetRange {
  return sheet.getRange(row, column, row, column)
}

/** Whether one cell index exists in the worksheet. */
function contains(sheet: SheetTarget, row: number, column: number): boolean {
  return index(row) && index(column) && row < sheet.getMaxRows() && column < sheet.getMaxColumns()
}

/** Whether a value is a non-negative safe index. */
function index(value: number): boolean {
  return Number.isSafeInteger(value) && value >= 0
}

/** Whether a value is a positive safe count. */
function positive(value: number): boolean {
  return Number.isSafeInteger(value) && value > 0
}

/**
 * Admit one persisted style value as a cell-style update.
 * @param style - the step's JSON style value.
 * @returns the style update, `null` to clear the style, or `undefined` when the value is not an object.
 */
function cellStyle(style: JsonValue): CellStyle | null | undefined {
  if (style === null) return null
  if (typeof style !== 'object' || Array.isArray(style)) return undefined
  return style
}
