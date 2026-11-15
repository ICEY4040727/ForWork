/** Structure summaries: counts, bounded previews, formulas, paragraphs, and malformed snapshots. */
import type { OfficeDocumentSnapshot } from '@deepseek-ai/dsh-office-exchange'
import type { JsonValue } from '@deepseek-ai/dsh-util-values'
import { describe, expect, it } from 'vitest'
import { OfficeToolResultError } from '../src/errors.ts'
import { summarizeDocument } from '../src/summary.ts'

/** Refuse one snapshot and return its stable code. */
function refusalOf(snapshot: OfficeDocumentSnapshot): string {
  try {
    summarizeDocument(snapshot, 10)
  } catch (error) {
    if (error instanceof OfficeToolResultError) return error.code
    throw error
  }
  throw new Error('expected a summary refusal')
}

/** Build a snapshot envelope around deliberately malformed JSON data. */
function badSnapshot(kind: 'workbook' | 'document', data: JsonValue): OfficeDocumentSnapshot {
  return { kind, data }
}

/** A workbook snapshot with one sheet and the supplied cell data. */
function workbook(cellData: JsonValue, extra: Record<string, JsonValue> = {}): OfficeDocumentSnapshot {
  return {
    kind: 'workbook',
    data: {
      sheetOrder: ['sheet-1'],
      sheets: {
        'sheet-1': { name: 'Report', cellData, ...extra },
      },
    },
  }
}

describe('workbook summaries', () => {
  it('lists cells, formulas, and complete counts', () => {
    const summary = summarizeDocument(workbook({
      0: { 0: { v: 'Name' }, 1: { v: 2, f: '=SUM(A1:A2)' } },
      2: { 3: { v: true } },
    }, { rowCount: 9, columnCount: 7 }), 10)

    expect(summary.kind).toBe('workbook')
    expect(summary.sheets).toEqual([{
      id: 'sheet-1',
      name: 'Report',
      rowCount: 9,
      columnCount: 7,
      cells: [
        { row: 0, column: 0, value: 'Name' },
        { row: 0, column: 1, value: 2, formula: '=SUM(A1:A2)' },
        { row: 2, column: 3, value: true },
      ],
      formulaCount: 1,
    }])
    expect(summary.cellCount).toBe(3)
    expect(summary.truncated).toBe(false)
  })

  it('derives counts when the sheet omits them and truncates the preview', () => {
    const summary = summarizeDocument(workbook({
      0: { 0: { v: 1 }, 2: { v: 2 } },
      4: { 1: { v: 3 } },
    }), 2)

    expect(summary.sheets[0]?.rowCount).toBe(5)
    expect(summary.sheets[0]?.columnCount).toBe(3)
    expect(summary.sheets[0]?.cells).toHaveLength(2)
    expect(summary.cellCount).toBe(3)
    expect(summary.truncated).toBe(true)
  })

  it('keeps an unnamed sheet under its id and an empty sheet at zero', () => {
    const summary = summarizeDocument({
      kind: 'workbook',
      data: { sheetOrder: ['sheet-1'], sheets: { 'sheet-1': {} } },
    }, 5)
    expect(summary.sheets[0]).toEqual({
      id: 'sheet-1',
      name: 'sheet-1',
      rowCount: 0,
      columnCount: 0,
      cells: [],
      formulaCount: 0,
    })
    expect(summary.truncated).toBe(false)
  })

  it('ignores non-index keys in rectangular data', () => {
    const summary = summarizeDocument(workbook({
      0: { 0: { v: 1 }, label: { v: 2 } },
      summary: { 0: { v: 3 } },
    }), 5)
    expect(summary.cellCount).toBe(1)
    expect(summary.sheets[0]?.cells).toEqual([{ row: 0, column: 0, value: 1 }])
  })

  const malformed: readonly (readonly [string, OfficeDocumentSnapshot])[] = [
    ['a non-object root', badSnapshot('workbook', null)],
    ['no sheets record', badSnapshot('workbook', { sheetOrder: [] })],
    ['no sheet order', badSnapshot('workbook', { sheets: {} })],
    ['a non-string sheet id', badSnapshot('workbook', { sheetOrder: [1], sheets: {} })],
    ['a missing sheet', badSnapshot('workbook', { sheetOrder: ['x'], sheets: {} })],
    ['a non-object row', workbook({ 0: [] })],
    ['a non-object cell', workbook({ 0: { 0: [] } })],
  ]
  it.each(malformed)('refuses %s', (_label, snapshot) => {
    expect(refusalOf(snapshot)).toBe('conversion-failed')
  })
})

describe('document summaries', () => {
  it('reads declared paragraphs and table dimensions', () => {
    const summary = summarizeDocument({
      kind: 'document',
      data: {
        body: {
          dataStream: 'Hello\rWorld\r',
          paragraphs: [{ startIndex: 0 }, { startIndex: 6 }],
          tables: [{ tableId: 'table-1' }],
        },
        tableSource: { 'table-1': { tableRows: [{}, {}], tableColumns: [{}, {}, {}] } },
      },
    }, 10)

    expect(summary.kind).toBe('document')
    expect(summary.paragraphs).toEqual([
      { index: 0, text: 'Hello' },
      { index: 1, text: 'World' },
    ])
    expect(summary.tables).toEqual([{ index: 0, rows: 2, columns: 3 }])
    expect(summary.cellCount).toBe(0)
  })

  it('derives paragraphs when the snapshot carries none', () => {
    const summary = summarizeDocument({
      kind: 'document',
      data: { body: { dataStream: 'A\rB\nC' } },
    }, 10)
    expect(summary.paragraphs).toEqual([
      { index: 0, text: 'A' },
      { index: 1, text: 'B' },
      { index: 2, text: 'C' },
    ])
    expect(summary.tables).toEqual([])
  })

  it('skips empty paragraphs and tolerates a missing table source', () => {
    const summary = summarizeDocument({
      kind: 'document',
      data: {
        body: { dataStream: 'A\r\rB', tables: [{ tableId: 'table-1' }] },
      },
    }, 10)
    expect(summary.paragraphs).toEqual([
      { index: 0, text: 'A' },
      { index: 2, text: 'B' },
    ])
    expect(summary.tables).toEqual([{ index: 0, rows: 0, columns: 0 }])
  })

  const malformed: readonly (readonly [string, OfficeDocumentSnapshot])[] = [
    ['a non-object root', badSnapshot('document', null)],
    ['no body', badSnapshot('document', {})],
    ['no stream', badSnapshot('document', { body: {} })],
    ['a non-object paragraph', badSnapshot('document', { body: { dataStream: 'x', paragraphs: [1] } })],
    ['an invalid paragraph start', badSnapshot('document', { body: { dataStream: 'x', paragraphs: [{ startIndex: -1 }] } })],
    ['an invalid paragraph end', badSnapshot('document', {
      body: { dataStream: 'abc', paragraphs: [{ startIndex: 2 }, { startIndex: 1 }] },
    })],
    ['a table without identity', badSnapshot('document', { body: { dataStream: 'x', tables: [{}] } })],
  ]
  it.each(malformed)('refuses %s', (_label, snapshot) => {
    expect(refusalOf(snapshot)).toBe('conversion-failed')
  })
})
