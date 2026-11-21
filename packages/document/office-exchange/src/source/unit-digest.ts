/**
 * Canonical unit projections for in-place save change detection.
 *
 * The digest of a unit must survive an editor round trip that moves stream
 * offsets, reorders unrelated units, or reallocates snapshot ids, while it must
 * change when the unit's text, style, or structure changes. Each projection
 * below therefore contains only the modelled content named by the data model:
 * it excludes `startIndex`, run positions, array ordinals, and snapshot ids,
 * and canonicalizes object keys before hashing.
 *
 * A digest is the lowercase SHA-256 hex of the canonical JSON serialization.
 * @module @forwork/office-exchange/source/unit-digest
 */

import { createHash } from 'node:crypto'
import type { ICellData, IDocumentData, IParagraph, ITable, ITableCell, ITableRow, ITextRun } from '@univerjs/core'

type CanonicalValue =
  | null
  | boolean
  | number
  | string
  | CanonicalValue[]
  | { [key: string]: CanonicalValue }

/**
 * One paragraph's modelled content, independent of stream positions.
 *
 * The caller supplies the stream text the paragraph covers and the runs that
 * belong to it. The digest never reads `startIndex`, run `st`/`ed`, or
 * `paragraphId`.
 */
export interface ParagraphUnit {
  /** Stream text from the paragraph's first character through its break. */
  readonly dataStream: string
  /** Runs inside the paragraph, in stream order. */
  readonly textRuns?: readonly ITextRun[]
  /** Paragraph record from `IDocumentBody.paragraphs`. */
  readonly paragraph: IParagraph
}

/** One table cell's modelled content. */
export interface TableCellUnit {
  /** Cell properties from `ITableRow.tableCells`. */
  readonly cell: ITableCell
  /** Paragraphs directly inside the cell, in stream order. */
  readonly paragraphs: readonly ParagraphUnit[]
  /** Nested tables inside the cell, in stream order. */
  readonly tables: readonly TableUnit[]
}

/** One table row's modelled content. */
export interface TableRowUnit {
  /** Row properties excluding `tableCells`, which {@link cells} carries. */
  readonly row: ITableRow
  /** Cells in this row, in column order. */
  readonly cells: readonly TableCellUnit[]
}

/** One table's modelled content. */
export interface TableUnit {
  /** Table properties excluding `tableRows` and `tableId`. */
  readonly table: ITable
  /** Rows in this table, in row order. */
  readonly rows: readonly TableRowUnit[]
}

/**
 * Digest one paragraph unit.
 * @param unit - Stream slice, runs, and paragraph properties of the paragraph.
 * @returns Lowercase SHA-256 hex of the canonical paragraph projection.
 */
export function digestParagraph(unit: ParagraphUnit): string {
  return hashCanonical(paragraphProjection(unit))
}

/**
 * Digest one table cell unit.
 * @param unit - Cell properties plus the paragraphs and nested tables inside it.
 * @returns Lowercase SHA-256 hex of the canonical cell projection.
 */
export function digestTableCell(unit: TableCellUnit): string {
  return hashCanonical(tableCellProjection(unit))
}

/**
 * Digest one sheet cell.
 * @param cell - Workbook cell data carrying the value, type, formula, hyperlink document, and style.
 * @returns Lowercase SHA-256 hex of the canonical sheet-cell projection.
 */
export function digestSheetCell(cell: ICellData): string {
  return hashCanonical(sheetCellProjection(cell))
}

function paragraphProjection(unit: ParagraphUnit): CanonicalValue {
  return {
    dataStream: unit.dataStream,
    textRuns: (unit.textRuns ?? []).map(run => ({
      styleId: run.sId ?? null,
      style: canonicalize(run.ts ?? null),
    })),
    paragraphStyle: canonicalize(unit.paragraph.paragraphStyle ?? null),
    bullet: canonicalize(unit.paragraph.bullet ?? null),
    listId: unit.paragraph.bullet?.listId ?? null,
  }
}

function tableCellProjection(unit: TableCellUnit): CanonicalValue {
  return {
    properties: canonicalize(unit.cell),
    paragraphs: unit.paragraphs.map(paragraphProjection),
    tables: unit.tables.map(tableProjection),
  }
}

function tableRowProjection(unit: TableRowUnit): CanonicalValue {
  return {
    properties: canonicalize({ ...unit.row, tableCells: undefined }),
    cells: unit.cells.map(tableCellProjection),
  }
}

function tableProjection(unit: TableUnit): CanonicalValue {
  return {
    properties: canonicalize({ ...unit.table, tableRows: undefined, tableId: undefined }),
    rows: unit.rows.map(tableRowProjection),
  }
}

function sheetCellProjection(cell: ICellData): CanonicalValue {
  return {
    value: canonicalize(cell.v ?? null),
    type: cell.t ?? null,
    formula: cell.f ?? null,
    hyperlink: hyperlinkProjection(cell.p ?? null),
    style: canonicalize(cell.s ?? null),
  }
}

function hyperlinkProjection(document: IDocumentData | null): CanonicalValue {
  if (document === null) return null
  const body = document.body
  return {
    dataStream: body?.dataStream ?? '',
    textRuns: (body?.textRuns ?? []).map(run => ({
      styleId: run.sId ?? null,
      style: canonicalize(run.ts ?? null),
    })),
    paragraphs: (body?.paragraphs ?? []).map(paragraph => ({
      styleId: paragraph.styleId ?? null,
      paragraphStyle: canonicalize(paragraph.paragraphStyle ?? null),
      bullet: canonicalize(paragraph.bullet ?? null),
    })),
    ranges: (body?.customRanges ?? []).map(range => ({
      rangeType: range.rangeType,
      properties: canonicalize(range.properties ?? null),
      wholeEntity: range.wholeEntity ?? null,
    })),
  }
}

function hashCanonical(projection: CanonicalValue): string {
  return createHash('sha256').update(JSON.stringify(projection), 'utf8').digest('hex')
}

function canonicalize(value: unknown): CanonicalValue {
  if (value === null || value === undefined) return null
  if (typeof value === 'string' || typeof value === 'boolean') return value
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value === 'bigint') return value.toString()
  if (Array.isArray(value)) {
    const items: unknown[] = value
    return items.map(item => canonicalize(item))
  }
  if (typeof value === 'object') {
    const record = value as Record<string, unknown>
    const result: { [key: string]: CanonicalValue } = {}
    for (const key of Object.keys(record).sort()) {
      const child = record[key]
      if (child === undefined) continue
      result[key] = canonicalize(child)
    }
    return result
  }
  return null
}
