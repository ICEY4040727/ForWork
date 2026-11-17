// @vitest-environment jsdom
/**
 * Sheet edit-step execution through the adapter seam.
 *
 * A model workbook stands in for the Univer facade: it applies the same
 * commands the runtime issues, keeps the native undo granularity (one history
 * entry per facade command), and projects the same wire snapshot. The suites
 * cover every declared step kind, atomic refusal, and the adapter's undo path.
 */
import { describe, expect, it, vi } from 'vitest'
import type { ICellData } from '@univerjs/core'
import type { DocumentEditOperation } from '@forwork/session/src/document-events.ts'
import type { JsonValue } from '@forwork/util-values'
import type { OfficeRuntime, OfficeSnapshot } from '../src/client/adapter.ts'
import { createAdapter } from '../src/client/adapter.ts'
import { applySheetOperation } from '../src/client/editor/sheet-operations.ts'
import type { SheetRange, SheetTarget, SheetWorkbook } from '../src/client/editor/sheet-operations.ts'
import { nativeHistory } from '../src/client/undo.ts'
import { jsonValueOf } from '../src/client/wire.ts'

const SHEET_ID = 'sheet-1'

/** One modelled cell: stored value, formula, and style. */
interface ModelCell {
  v: string | number | boolean | null
  f: string | null
  s: JsonValue
}

/** One modelled sheet's mutable state. */
interface SheetState {
  rowCount: number
  columnCount: number
  cells: Map<string, ModelCell>
}

/** One modelled workbook state, captured for undo and redo. */
type BookState = Map<string, SheetState>

class ModelRange implements SheetRange {
  constructor(
    private readonly book: ModelWorkbook,
    private readonly sheet: ModelSheet,
    private readonly startRow: number,
    private readonly startColumn: number,
    private readonly endRow: number,
    private readonly endColumn: number,
  ) {}

  setValue(value: string | number | boolean | ICellData): void {
    this.book.checkpoint()
    for (let row = this.startRow; row <= this.endRow; row += 1) {
      for (let column = this.startColumn; column <= this.endColumn; column += 1) {
        this.write(row, column, value)
      }
    }
  }

  setFormula(formula: string): void {
    this.book.checkpoint()
    for (let row = this.startRow; row <= this.endRow; row += 1) {
      for (let column = this.startColumn; column <= this.endColumn; column += 1) {
        this.cell(row, column).f = formula
      }
    }
  }

  private write(row: number, column: number, value: string | number | boolean | ICellData): void {
    const cell = this.cell(row, column)
    if (typeof value === 'object') {
      if ('s' in value) {
        cell.s = value.s === null || value.s === undefined ? null : jsonValueOf(value.s)
        return
      }
      if ('v' in value) cell.v = value.v ?? null
      if ('f' in value) cell.f = value.f ?? null
      return
    }
    cell.v = value
    cell.f = null
  }

  private cell(row: number, column: number): ModelCell {
    const key = `${row}:${column}`
    const existing = this.sheet.cells.get(key)
    if (existing !== undefined) return existing
    const cell: ModelCell = { v: null, f: null, s: null }
    this.sheet.cells.set(key, cell)
    return cell
  }
}

class ModelSheet implements SheetTarget {
  readonly cells = new Map<string, ModelCell>()

  constructor(
    private readonly book: ModelWorkbook,
    public rowCount: number,
    public columnCount: number,
  ) {}

  getMaxRows(): number {
    return this.rowCount
  }

  getMaxColumns(): number {
    return this.columnCount
  }

  insertRows(row: number, count: number): void {
    this.book.checkpoint()
    const moved = new Map<string, ModelCell>()
    for (const [key, cell] of this.cells) {
      const [index, column] = split(key)
      moved.set(index >= row ? `${index + count}:${column}` : key, cell)
    }
    this.replace(moved)
    this.rowCount += count
  }

  deleteRows(row: number, count: number): void {
    this.book.checkpoint()
    const moved = new Map<string, ModelCell>()
    for (const [key, cell] of this.cells) {
      const [index, column] = split(key)
      if (index >= row && index < row + count) continue
      moved.set(index >= row + count ? `${index - count}:${column}` : key, cell)
    }
    this.replace(moved)
    this.rowCount -= count
  }

  insertColumns(column: number, count: number): void {
    this.book.checkpoint()
    const moved = new Map<string, ModelCell>()
    for (const [key, cell] of this.cells) {
      const [row, index] = split(key)
      moved.set(index >= column ? `${row}:${index + count}` : key, cell)
    }
    this.replace(moved)
    this.columnCount += count
  }

  deleteColumns(column: number, count: number): void {
    this.book.checkpoint()
    const moved = new Map<string, ModelCell>()
    for (const [key, cell] of this.cells) {
      const [row, index] = split(key)
      if (index >= column && index < column + count) continue
      moved.set(index >= column + count ? `${row}:${index - count}` : key, cell)
    }
    this.replace(moved)
    this.columnCount -= count
  }

  getRange(startRow: number, startColumn: number, endRow: number, endColumn: number): SheetRange {
    return new ModelRange(this.book, this, startRow, startColumn, endRow, endColumn)
  }

  json(): JsonValue {
    const cellData: Record<string, JsonValue> = {}
    for (const [key, cell] of this.cells) {
      if (cell.v === null && cell.f === null && cell.s === null) continue
      const [row, column] = split(key)
      const data: Record<string, JsonValue> = {}
      if (cell.v !== null) data['v'] = cell.v
      if (cell.f !== null) data['f'] = cell.f
      if (cell.s !== null) data['s'] = cell.s
      const rowCells = recordOf(cellData[row])
      rowCells[String(column)] = data
      cellData[String(row)] = rowCells
    }
    return { rowCount: this.rowCount, columnCount: this.columnCount, cellData }
  }

  state(): SheetState {
    const cells = new Map<string, ModelCell>()
    for (const [key, cell] of this.cells) cells.set(key, { ...cell })
    return { rowCount: this.rowCount, columnCount: this.columnCount, cells }
  }

  private replace(cells: Map<string, ModelCell>): void {
    this.cells.clear()
    for (const [key, cell] of cells) this.cells.set(key, cell)
  }
}

class ModelWorkbook implements SheetWorkbook {
  private readonly sheets = new Map<string, ModelSheet>()
  private readonly past: BookState[] = []
  private readonly future: BookState[] = []

  constructor(rowCount: number, columnCount: number) {
    this.sheets.set(SHEET_ID, new ModelSheet(this, rowCount, columnCount))
  }

  get sheet(): ModelSheet {
    const sheet = this.sheets.get(SHEET_ID)
    if (sheet === undefined) throw new Error('model workbook has no sheet')
    return sheet
  }

  get historyLength(): number {
    return this.past.length
  }

  getSheetBySheetId(sheetId: string): SheetTarget | null {
    return this.sheets.get(sheetId) ?? null
  }

  checkpoint(): void {
    this.past.push(this.state())
    this.future.length = 0
  }

  undo(): void {
    const previous = this.past.pop()
    if (previous === undefined) return
    this.future.push(this.state())
    this.restore(previous)
  }

  redo(): void {
    const next = this.future.pop()
    if (next === undefined) return
    this.past.push(this.state())
    this.restore(next)
  }

  snapshot(): OfficeSnapshot {
    const sheets: Record<string, JsonValue> = {}
    for (const [id, sheet] of this.sheets) sheets[id] = sheet.json()
    return { kind: 'workbook', data: { sheets } }
  }

  private state(): BookState {
    const state: BookState = new Map()
    for (const [id, sheet] of this.sheets) state.set(id, sheet.state())
    return state
  }

  private restore(state: BookState): void {
    this.sheets.clear()
    for (const [id, saved] of state) {
      const sheet = new ModelSheet(this, saved.rowCount, saved.columnCount)
      for (const [key, cell] of saved.cells) sheet.cells.set(key, { ...cell })
      this.sheets.set(id, sheet)
    }
  }
}

/** The workbook every suite starts from: three populated cells in an 8×6 sheet. */
function workbookWith(): ModelWorkbook {
  const book = new ModelWorkbook(8, 6)
  const sheet = book.sheet
  sheet.cells.set('0:0', { v: 1, f: null, s: null })
  sheet.cells.set('0:1', { v: null, f: '=A1', s: null })
  sheet.cells.set('1:0', { v: 2, f: null, s: null })
  return book
}

/** A runtime over the model workbook, using the adapter's native history. */
function runtimeOf(book: ModelWorkbook): OfficeRuntime {
  const listeners = new Set<() => void>()
  const notify = (): void => {
    for (const listener of [...listeners]) listener()
  }
  const history = nativeHistory({
    undo: async () => {
      book.undo()
      notify()
      return true
    },
    redo: async () => {
      book.redo()
      notify()
      return true
    },
  })
  return {
    getSnapshot: () => book.snapshot(),
    apply(operation) {
      const outcome = applySheetOperation(book, operation)
      if (outcome.status === 'applied') notify()
      return outcome
    },
    undo: () => { history.undo() },
    redo: () => { history.redo() },
    onMutation(listener) {
      listeners.add(listener)
      return () => { listeners.delete(listener) }
    },
    dispose() { listeners.clear() },
  }
}

/** Split one model cell key into its row and column indexes. */
function split(key: string): [number, number] {
  const [row, column] = key.split(':')
  return [Number(row), Number(column)]
}

/** Narrow a JSON value to a plain object, keeping absence. */
function objectOf(value: JsonValue | undefined): Record<string, JsonValue> | undefined {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return undefined
  return value
}

/** Narrow a JSON value to a plain object for construction. */
function recordOf(value: JsonValue | undefined): Record<string, JsonValue> {
  return objectOf(value) ?? {}
}

/** The modelled sheet's projection inside one snapshot. */
function sheetOf(snapshot: OfficeSnapshot): Record<string, JsonValue> {
  const sheet = objectOf(objectOf(snapshot.data)?.['sheets'])?.[SHEET_ID]
  if (sheet === undefined) throw new Error('snapshot has no modelled sheet')
  return objectOf(sheet) ?? {}
}

/** One cell's projection, or undefined when the cell stores nothing. */
function cellAt(snapshot: OfficeSnapshot, row: number, column: number): Record<string, JsonValue> | undefined {
  const cells = objectOf(sheetOf(snapshot)['cellData'])
  const rowCells = objectOf(cells?.[String(row)])
  return objectOf(rowCells?.[String(column)])
}

describe('applySheetOperation', () => {
  it('writes and clears a formula', () => {
    const book = workbookWith()
    expect(applySheetOperation(book, [
      { kind: 'formula', sheetId: SHEET_ID, row: 2, column: 1, formula: '=SUM(A1:A2)' },
    ])).toEqual({ status: 'applied' })
    expect(cellAt(book.snapshot(), 2, 1)).toEqual({ f: '=SUM(A1:A2)' })

    expect(applySheetOperation(book, [
      { kind: 'formula', sheetId: SHEET_ID, row: 0, column: 1, formula: null },
    ])).toEqual({ status: 'applied' })
    expect(cellAt(book.snapshot(), 0, 1)).toBeUndefined()
  })

  it('inserts and deletes rows', () => {
    const book = workbookWith()
    expect(applySheetOperation(book, [
      { kind: 'row', sheetId: SHEET_ID, row: 1, count: 2, action: 'insert' },
    ])).toEqual({ status: 'applied' })
    expect(sheetOf(book.snapshot())['rowCount']).toBe(10)
    expect(cellAt(book.snapshot(), 3, 0)).toEqual({ v: 2 })

    expect(applySheetOperation(book, [
      { kind: 'row', sheetId: SHEET_ID, row: 1, count: 1, action: 'delete' },
    ])).toEqual({ status: 'applied' })
    expect(sheetOf(book.snapshot())['rowCount']).toBe(9)
    expect(cellAt(book.snapshot(), 2, 0)).toEqual({ v: 2 })
  })

  it('inserts and deletes columns', () => {
    const book = workbookWith()
    expect(applySheetOperation(book, [
      { kind: 'column', sheetId: SHEET_ID, column: 1, count: 2, action: 'insert' },
    ])).toEqual({ status: 'applied' })
    expect(sheetOf(book.snapshot())['columnCount']).toBe(8)
    expect(cellAt(book.snapshot(), 0, 3)).toEqual({ f: '=A1' })

    expect(applySheetOperation(book, [
      { kind: 'column', sheetId: SHEET_ID, column: 1, count: 1, action: 'delete' },
    ])).toEqual({ status: 'applied' })
    expect(sheetOf(book.snapshot())['columnCount']).toBe(7)
    expect(cellAt(book.snapshot(), 0, 2)).toEqual({ f: '=A1' })
  })

  it('styles every cell in a rectangular range', () => {
    const book = workbookWith()
    const style = { bl: 1, bg: { rgb: '#ff0000' } } as const
    expect(applySheetOperation(book, [
      {
        kind: 'style',
        sheetId: SHEET_ID,
        startRow: 0,
        startColumn: 0,
        endRow: 1,
        endColumn: 1,
        style,
      },
    ])).toEqual({ status: 'applied' })
    const snapshot = book.snapshot()
    expect(cellAt(snapshot, 0, 0)?.['s']).toEqual(style)
    expect(cellAt(snapshot, 0, 1)?.['s']).toEqual(style)
    expect(cellAt(snapshot, 1, 0)?.['s']).toEqual(style)
    expect(cellAt(snapshot, 1, 1)?.['s']).toEqual(style)
  })

  it('applies a mixed batch in declared order', () => {
    const book = workbookWith()
    expect(applySheetOperation(book, [
      { kind: 'cell', sheetId: SHEET_ID, row: 3, column: 0, value: 'Sprocket' },
      { kind: 'formula', sheetId: SHEET_ID, row: 3, column: 1, formula: '=A4' },
      {
        kind: 'style',
        sheetId: SHEET_ID,
        startRow: 3,
        startColumn: 3,
        endRow: 3,
        endColumn: 3,
        style: { bl: 1 },
      },
      { kind: 'row', sheetId: SHEET_ID, row: 6, count: 1, action: 'insert' },
      { kind: 'column', sheetId: SHEET_ID, column: 4, count: 1, action: 'insert' },
    ])).toEqual({ status: 'applied' })
    const snapshot = book.snapshot()
    expect(sheetOf(snapshot)['rowCount']).toBe(9)
    expect(sheetOf(snapshot)['columnCount']).toBe(7)
    expect(cellAt(snapshot, 3, 0)).toEqual({ v: 'Sprocket' })
    expect(cellAt(snapshot, 3, 1)).toEqual({ f: '=A4' })
    expect(cellAt(snapshot, 3, 3)?.['s']).toEqual({ bl: 1 })
  })

  const refused: readonly (readonly [string, DocumentEditOperation])[] = [
    ['an unknown sheet', [{ kind: 'cell', sheetId: 'missing', row: 0, column: 0, value: 1 }]],
    ['a cell outside the sheet', [{ kind: 'cell', sheetId: SHEET_ID, row: 8, column: 0, value: 1 }]],
    ['a row insertion at the last row', [{ kind: 'row', sheetId: SHEET_ID, row: 8, count: 1, action: 'insert' }]],
    ['a row deletion past the sheet', [{ kind: 'row', sheetId: SHEET_ID, row: 7, count: 2, action: 'delete' }]],
    ['an empty row count', [{ kind: 'column', sheetId: SHEET_ID, column: 0, count: 0, action: 'insert' }]],
    ['a reversed style range', [{
      kind: 'style',
      sheetId: SHEET_ID,
      startRow: 2,
      startColumn: 0,
      endRow: 1,
      endColumn: 0,
      style: { bl: 1 },
    }]],
    ['a style range past the sheet', [{
      kind: 'style',
      sheetId: SHEET_ID,
      startRow: 0,
      startColumn: 0,
      endRow: 0,
      endColumn: 6,
      style: { bl: 1 },
    }]],
    ['a paragraph step', [{ kind: 'paragraph', index: 0, text: 'Hello' }]],
  ]

  it.each(refused)('refuses %s without applying anything', (_name, operation) => {
    const book = workbookWith()
    const before = book.snapshot()
    expect(applySheetOperation(book, operation).status).toBe('unsupported-feature')
    expect(book.snapshot()).toEqual(before)
    expect(book.historyLength).toBe(0)
  })

  it('refuses a batch whose later step is invalid and keeps earlier steps unapplied', () => {
    const book = workbookWith()
    const before = book.snapshot()
    const outcome = applySheetOperation(book, [
      { kind: 'cell', sheetId: SHEET_ID, row: 3, column: 0, value: 'Sprocket' },
      { kind: 'row', sheetId: SHEET_ID, row: 0, count: 1, action: 'insert' },
      { kind: 'formula', sheetId: 'missing', row: 0, column: 0, formula: '=A1' },
    ])
    expect(outcome).toEqual({
      status: 'unsupported-feature',
      features: [{ kind: 'other', detail: 'sheet missing' }],
    })
    expect(book.snapshot()).toEqual(before)
    expect(book.historyLength).toBe(0)
  })

  it('undoes and redoes an applied operation through the native history', async () => {
    const book = workbookWith()
    const before = book.snapshot()
    const adapter = createAdapter(async () => runtimeOf(book))
    const handle = await adapter.mount(document.createElement('div'), {
      kind: 'workbook',
      snapshot: before,
      onChange: vi.fn(),
    })
    expect(handle.apply([
      { kind: 'formula', sheetId: SHEET_ID, row: 2, column: 0, formula: '=A1+A2' },
    ])).toEqual({ status: 'applied' })
    expect(cellAt(handle.getSnapshot(), 2, 0)).toEqual({ f: '=A1+A2' })

    handle.undo()
    expect(handle.getSnapshot()).toEqual(before)

    handle.redo()
    expect(cellAt(handle.getSnapshot(), 2, 0)).toEqual({ f: '=A1+A2' })
    handle.dispose()
  })
})
