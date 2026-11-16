/**
 * Model-facing structure summaries of one editable Office snapshot.
 *
 * The summary is derived from the live editor's snapshot so the model can read
 * sheet, cell, formula, paragraph, and table structure without receiving raw
 * file bytes or the complete unit. Cell previews are bounded; counts describe
 * the complete unit so a truncated preview is never mistaken for a small one.
 * @module @forwork/tool-office/summary
 */

import type { OfficeDocumentSnapshot } from '@forwork/office-exchange'
import type { JsonValue } from '@forwork/util-values'
import { OfficeToolResultError } from './errors.ts'

/** One spreadsheet cell in a bounded preview. */
export interface OfficeCellSummary {
  /** Zero-based row index. */
  readonly row: number
  /** Zero-based column index. */
  readonly column: number
  /** JSON cell value, or `null` for an empty cell. */
  readonly value: JsonValue
  /** Formula text when the cell holds one. */
  readonly formula?: string
}

/** One sheet's bounded cell preview and complete counts. */
export interface OfficeSheetSummary {
  /** Sheet identity inside the workbook snapshot. */
  readonly id: string
  /** Display name of the sheet. */
  readonly name: string
  /** Rows in the complete sheet. */
  readonly rowCount: number
  /** Columns in the complete sheet. */
  readonly columnCount: number
  /** Bounded cell preview in row-major order. */
  readonly cells: readonly OfficeCellSummary[]
  /** Formulas present in the complete sheet, including cells outside the preview. */
  readonly formulaCount: number
}

/** One text paragraph. */
export interface OfficeParagraphSummary {
  /** Zero-based paragraph index used by paragraph edits. */
  readonly index: number
  /** Paragraph text without its trailing break. */
  readonly text: string
}

/** One flow table inside a text document. */
export interface OfficeTableSummary {
  /** Zero-based table index inside the document. */
  readonly index: number
  /** Rows in the complete table. */
  readonly rows: number
  /** Columns in the complete table. */
  readonly columns: number
}

/** Model-facing structure of one open document. */
export interface OfficeDocumentSummary {
  /** Kind of editable unit the summary describes. */
  readonly kind: 'workbook' | 'document'
  /** Spreadsheet sheets; empty for a text document. */
  readonly sheets: readonly OfficeSheetSummary[]
  /** Text paragraphs; empty for a workbook. */
  readonly paragraphs: readonly OfficeParagraphSummary[]
  /** Flow tables; empty for a workbook. */
  readonly tables: readonly OfficeTableSummary[]
  /** Cells in the complete unit, including cells outside the preview. */
  readonly cellCount: number
  /** Whether the cell preview omits cells present in the unit. */
  readonly truncated: boolean
}

/** A JSON object view of a snapshot value. */
type JsonObject = Record<string, JsonValue>

/** Narrow a JSON value to a plain object. */
function objectOf(value: JsonValue | undefined): JsonObject | undefined {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return undefined
  return value
}

/** Require a JSON object or refuse the snapshot. */
function requireObject(value: JsonValue | undefined, label: string): JsonObject {
  const object = objectOf(value)
  if (object === undefined) throw new OfficeToolResultError('conversion-failed', `The ${label} is not a JSON object.`)
  return object
}

/** Require a JSON array or treat an absent field as empty. */
function arrayOf(value: JsonValue | undefined): readonly JsonValue[] {
  return Array.isArray(value) ? value : []
}

/** Numeric keys of a record in ascending order, discarding non-index keys. */
function indexKeys(record: JsonObject): number[] {
  const keys: number[] = []
  for (const key of Object.keys(record)) {
    const index = Number(key)
    if (Number.isSafeInteger(index) && index >= 0) keys.push(index)
  }
  return keys.sort((left, right) => left - right)
}

/** A declared count, or the observed maximum index plus one. */
function countOf(declared: JsonValue | undefined, observed: number): number {
  return typeof declared === 'number' && Number.isSafeInteger(declared) && declared >= 0 ? declared : observed
}

/** Summarize one sheet and stop adding preview cells once the budget is spent. */
function sheetSummary(
  id: string,
  value: JsonValue | undefined,
  budget: { remaining: number; total: number },
): { readonly summary: OfficeSheetSummary; readonly truncated: boolean } {
  const sheet = requireObject(value, `sheet "${id}"`)
  const cells = objectOf(sheet['cellData']) ?? {}
  const rows = indexKeys(cells)
  let maxColumn = -1
  let formulaCount = 0
  const preview: OfficeCellSummary[] = []
  for (const row of rows) {
    const rowCells = requireObject(cells[String(row)], `row ${row} of sheet "${id}"`)
    for (const column of indexKeys(rowCells)) {
      maxColumn = Math.max(maxColumn, column)
      const cell = requireObject(rowCells[String(column)], `cell ${row}:${column} of sheet "${id}"`)
      const formula = typeof cell['f'] === 'string' ? cell['f'] : undefined
      if (formula !== undefined) formulaCount += 1
      budget.total += 1
      if (budget.remaining > 0) {
        preview.push({
          row,
          column,
          value: cell['v'] ?? null,
          ...formula === undefined ? {} : { formula },
        })
        budget.remaining -= 1
      }
    }
  }
  const lastRow = rows.at(-1) ?? -1
  return {
    summary: {
      id,
      name: typeof sheet['name'] === 'string' ? sheet['name'] : id,
      rowCount: countOf(sheet['rowCount'], lastRow + 1),
      columnCount: countOf(sheet['columnCount'], maxColumn + 1),
      cells: preview,
      formulaCount,
    },
    truncated: budget.remaining === 0 && budget.total > preview.length,
  }
}

/** Summarize a workbook snapshot. */
function workbookSummary(data: JsonValue, maxCells: number): OfficeDocumentSummary {
  const workbook = requireObject(data, 'workbook snapshot')
  const order = workbook['sheetOrder']
  const sheets = requireObject(workbook['sheets'], 'workbook sheets')
  if (!Array.isArray(order)) throw new OfficeToolResultError('conversion-failed', 'The workbook has no sheet order.')
  const budget = { remaining: maxCells, total: 0 }
  let truncated = false
  const summaries = order.map((entry) => {
    if (typeof entry !== 'string') {
      throw new OfficeToolResultError('conversion-failed', 'The workbook sheet order is not a list of sheet ids.')
    }
    const result = sheetSummary(entry, sheets[entry], budget)
    truncated ||= result.truncated
    return result.summary
  })
  return {
    kind: 'workbook',
    sheets: summaries,
    paragraphs: [],
    tables: [],
    cellCount: budget.total,
    truncated,
  }
}

/** Derive paragraph text boundaries when the snapshot carries no paragraph list. */
function derivedParagraphs(stream: string): OfficeParagraphSummary[] {
  const paragraphs: OfficeParagraphSummary[] = []
  let start = 0
  let index = 0
  for (let offset = 0; offset <= stream.length; offset += 1) {
    const character = stream[offset]
    if (offset < stream.length && character !== '\r' && character !== '\n') continue
    if (offset > start) paragraphs.push({ index, text: stream.slice(start, offset) })
    index += 1
    start = offset + 1
  }
  return paragraphs
}

/** Summarize a text-document snapshot. */
function documentSummary(data: JsonValue): OfficeDocumentSummary {
  const document = requireObject(data, 'document snapshot')
  const body = requireObject(document['body'], 'document body')
  const stream = body['dataStream']
  if (typeof stream !== 'string') {
    throw new OfficeToolResultError('conversion-failed', 'The document has no data stream.')
  }
  const declared = arrayOf(body['paragraphs'])
  const paragraphs: OfficeParagraphSummary[] = declared.length === 0
    ? derivedParagraphs(stream)
    : declared.map((entry, index) => {
      const paragraph = requireObject(entry, `paragraph ${index}`)
      const start = paragraph['startIndex']
      if (typeof start !== 'number' || !Number.isSafeInteger(start) || start < 0 || start > stream.length) {
        throw new OfficeToolResultError('conversion-failed', `Paragraph ${index} has an invalid start index.`)
      }
      const next = index + 1 < declared.length ? requireObject(declared[index + 1], `paragraph ${index + 1}`)['startIndex'] : stream.length
      if (typeof next !== 'number' || !Number.isSafeInteger(next) || next < start || next > stream.length) {
        throw new OfficeToolResultError('conversion-failed', `Paragraph ${index} has an invalid end index.`)
      }
      return { index, text: stream.slice(start, next).replace(/[\r\n]+$/u, '') }
    })
  const tableSource = objectOf(document['tableSource']) ?? {}
  const tables = arrayOf(body['tables']).map((entry, index) => {
    const table = requireObject(entry, `table ${index}`)
    const tableId = table['tableId']
    if (typeof tableId !== 'string') {
      throw new OfficeToolResultError('conversion-failed', `Table ${index} has no identity.`)
    }
    const source = objectOf(tableSource[tableId]) ?? {}
    return {
      index,
      rows: arrayOf(source['tableRows']).length,
      columns: arrayOf(source['tableColumns']).length,
    }
  })
  return { kind: 'document', sheets: [], paragraphs, tables, cellCount: 0, truncated: false }
}

/**
 * Summarize one editable snapshot for the model.
 * @param snapshot - live unit snapshot returned by the editor.
 * @param maxCells - largest cell preview the summary may list.
 * @returns structure counts plus a bounded cell preview.
 * @throws {OfficeToolResultError} `conversion-failed` when the snapshot does not match its kind.
 */
export function summarizeDocument(snapshot: OfficeDocumentSnapshot, maxCells: number): OfficeDocumentSummary {
  return snapshot.kind === 'workbook'
    ? workbookSummary(snapshot.data, maxCells)
    : documentSummary(snapshot.data)
}
