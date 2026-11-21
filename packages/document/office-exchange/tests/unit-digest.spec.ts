/** Canonical unit projections: content changes alter the digest; positions and ids do not. */
import {
  BooleanNumber,
  CellValueType,
  CustomRangeType,
  ObjectRelativeFromH,
  ObjectRelativeFromV,
  TableAlignmentType,
  TableRowHeightRule,
  TableSizeType,
  TableTextWrapType,
} from '@univerjs/core'
import type {
  ICellData,
  IDocumentBody,
  IDocumentData,
  IParagraph,
  ITable,
  ITableCell,
  ITableRow,
  ITextRun,
  ITextStyle,
} from '@univerjs/core'
import { describe, expect, it } from 'vitest'
import type { ParagraphUnit, TableCellUnit, TableRowUnit, TableUnit } from '../src/source/unit-digest.ts'
import { digestParagraph, digestSheetCell, digestTableCell } from '../src/source/unit-digest.ts'

function paragraph(overrides: Partial<IParagraph> = {}): IParagraph {
  return { startIndex: 5, paragraphId: 'para_1', ...overrides }
}

function run(st: number, ed: number, style: string, textStyle: ITextStyle): ITextRun {
  return { st, ed, sId: style, ts: textStyle }
}

function paragraphUnit(overrides: Partial<ParagraphUnit> = {}): ParagraphUnit {
  return {
    dataStream: 'hello\r',
    textRuns: [run(0, 5, 'style-1', { ff: 'Arial', fs: 12 })],
    paragraph: paragraph({ paragraphStyle: { textStyle: { fs: 12 } } }),
    ...overrides,
  }
}

function emptyTableCell(): ITableCell {
  return {}
}

function tableCellUnit(overrides: Partial<TableCellUnit> = {}): TableCellUnit {
  return {
    cell: emptyTableCell(),
    paragraphs: [paragraphUnit()],
    tables: [],
    ...overrides,
  }
}

function row(overrides: Partial<ITableRow> = {}): ITableRow {
  return {
    tableCells: [],
    trHeight: { val: { v: 20 }, hRule: TableRowHeightRule.AUTO },
    ...overrides,
  }
}

function rowUnit(overrides: Partial<TableRowUnit> = {}): TableRowUnit {
  return { row: row(), cells: [tableCellUnit()], ...overrides }
}

function table(overrides: Partial<ITable> = {}): ITable {
  return {
    tableRows: [],
    tableColumns: [{ size: { type: TableSizeType.SPECIFIED, width: { v: 100 } } }],
    tableId: 'table-1',
    align: TableAlignmentType.START,
    indent: { v: 0 },
    textWrap: TableTextWrapType.NONE,
    position: {
      positionH: { relativeFrom: ObjectRelativeFromH.PAGE },
      positionV: { relativeFrom: ObjectRelativeFromV.PAGE },
    },
    dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
    size: { type: TableSizeType.SPECIFIED, width: { v: 100 } },
    ...overrides,
  }
}

function tableUnit(overrides: Partial<TableUnit> = {}): TableUnit {
  return { table: table(), rows: [rowUnit()], ...overrides }
}

function hyperlinkBody(overrides: Partial<IDocumentBody> = {}): IDocumentBody {
  return {
    dataStream: 'link\r',
    textRuns: [run(0, 4, 'link-style', { ff: 'Arial' })],
    paragraphs: [paragraph({
      startIndex: 4,
      paragraphId: 'para_link',
      styleId: 'LinkStyle',
      paragraphStyle: { textStyle: { ff: 'Arial' } },
      bullet: { listType: 'bullet', listId: 'list-1', nestingLevel: 0 },
    })],
    customRanges: [{
      startIndex: 0,
      endIndex: 3,
      rangeId: 'hyperlink-1',
      rangeType: CustomRangeType.HYPERLINK,
      properties: { url: 'https://example.com' },
    }],
    ...overrides,
  }
}

function hyperlinkDocument(body: IDocumentBody, id = 'doc-1'): IDocumentData {
  return { id, documentStyle: {}, body }
}

function sheetCell(overrides: Partial<ICellData> = {}): ICellData {
  return {
    v: 'hello',
    t: CellValueType.STRING,
    f: '=A1',
    p: hyperlinkDocument(hyperlinkBody()),
    s: { ff: 'Arial', fs: 12 },
    ...overrides,
  }
}

describe('digestParagraph', () => {
  it('returns a 64-character lowercase SHA-256 digest', () => {
    expect(digestParagraph(paragraphUnit())).toMatch(/^[0-9a-f]{64}$/u)
  })

  it('changes when text, run style, paragraph style, bullet, or list id changes', () => {
    const base = digestParagraph(paragraphUnit())
    expect(digestParagraph(paragraphUnit({ dataStream: 'world\r' }))).not.toBe(base)
    expect(digestParagraph(paragraphUnit({ textRuns: [run(0, 5, 'style-2', { ff: 'Arial', fs: 12 })] }))).not.toBe(base)
    expect(digestParagraph(paragraphUnit({ textRuns: [run(0, 5, 'style-1', { ff: 'Arial', fs: 14 })] }))).not.toBe(base)
    expect(digestParagraph(paragraphUnit({
      paragraph: paragraph({ paragraphStyle: { textStyle: { fs: 14 } } }),
    }))).not.toBe(base)
    expect(digestParagraph(paragraphUnit({
      paragraph: paragraph({ bullet: { listType: 'bullet', listId: 'list-1', nestingLevel: 0 } }),
    }))).not.toBe(base)
    expect(digestParagraph(paragraphUnit({
      paragraph: paragraph({ bullet: { listType: 'number', listId: 'list-1', nestingLevel: 0 } }),
    }))).not.toBe(digestParagraph(paragraphUnit({
      paragraph: paragraph({ bullet: { listType: 'bullet', listId: 'list-1', nestingLevel: 0 } }),
    })))
  })

  it('ignores startIndex, paragraph id, and run bounds', () => {
    const shifted = paragraphUnit({
      textRuns: [run(40, 45, 'style-1', { ff: 'Arial', fs: 12 })],
      paragraph: paragraph({ startIndex: 99, paragraphId: 'para_other', paragraphStyle: { textStyle: { fs: 12 } } }),
    })
    expect(digestParagraph(shifted)).toBe(digestParagraph(paragraphUnit()))
  })

  it('ignores imported bullet fields that the canonical projection does not name', () => {
    const withSource = paragraphUnit({
      paragraph: paragraph({
        paragraphStyle: { textStyle: { fs: 12 } },
        sourceBullet: { listType: 'bullet', listId: 'list-9', nestingLevel: 1 },
        directBullet: null,
      }),
    })
    expect(digestParagraph(withSource)).toBe(digestParagraph(paragraphUnit()))
  })

  it('canonicalizes object key order and drops undefined style fields', () => {
    const reordered = paragraphUnit({
      textRuns: [run(0, 5, 'style-1', { fs: 12, ff: 'Arial' })],
    })
    expect(digestParagraph(reordered)).toBe(digestParagraph(paragraphUnit()))
    type StyleWithUndefined = ITextStyle & { readonly unused?: string | undefined }
    const withUndefined: StyleWithUndefined = { ff: 'Arial', fs: 12, unused: undefined }
    const dropped = paragraphUnit({ textRuns: [run(0, 5, 'style-1', withUndefined)] })
    expect(digestParagraph(dropped)).toBe(digestParagraph(paragraphUnit()))
  })

  it('treats absent optional run and paragraph fields as empty projections', () => {
    const sparse: ParagraphUnit = {
      dataStream: 'x\r',
      paragraph: { startIndex: 1, paragraphId: 'para_sparse' },
    }
    expect(digestParagraph(sparse)).toMatch(/^[0-9a-f]{64}$/u)
    const bareRun = paragraphUnit({ textRuns: [{ st: 0, ed: 5 }] })
    expect(digestParagraph(bareRun)).toMatch(/^[0-9a-f]{64}$/u)
    expect(digestParagraph(bareRun)).not.toBe(digestParagraph(paragraphUnit()))
  })

  it('canonicalizes arrays and non-JSON style values deterministically', () => {
    const gradient: ITextStyle = {
      textFill: {
        type: 'gradient',
        gradient: { stops: [{ offset: 0, color: '#000000' }, { offset: 1, color: '#ffffff' }] },
      },
    }
    const first = digestParagraph(paragraphUnit({ textRuns: [run(0, 5, 'style-1', gradient)] }))
    expect(first).toBe(digestParagraph(paragraphUnit({ textRuns: [run(0, 5, 'style-1', gradient)] })))
    type StyleWithExtras = ITextStyle & {
      readonly bad?: number
      readonly big?: bigint
      readonly fn?: () => void
      readonly extras?: readonly (string | undefined)[]
    }
    const extras: StyleWithExtras = {
      fs: Number.NaN,
      bad: Number.POSITIVE_INFINITY,
      big: 12n,
      fn: () => undefined,
      extras: ['a', undefined],
    }
    expect(digestParagraph(paragraphUnit({ textRuns: [run(0, 5, 'style-1', extras)] }))).toMatch(/^[0-9a-f]{64}$/u)
  })
})

describe('digestTableCell', () => {
  it('changes when cell content, cell properties, row properties, table properties, or a nested table changes', () => {
    const base = digestTableCell(tableCellUnit())
    expect(digestTableCell(tableCellUnit({
      paragraphs: [paragraphUnit({ dataStream: 'changed\r' })],
    }))).not.toBe(base)
    expect(digestTableCell(tableCellUnit({
      cell: { tcFitText: BooleanNumber.TRUE },
    }))).not.toBe(base)
    const nested = tableUnit()
    const withNested = tableCellUnit({ tables: [nested] })
    expect(digestTableCell(withNested)).not.toBe(base)
    expect(digestTableCell(tableCellUnit({
      tables: [tableUnit({ table: table({ align: TableAlignmentType.CENTER }) })],
    }))).not.toBe(digestTableCell(withNested))
    expect(digestTableCell(tableCellUnit({
      tables: [tableUnit({ table: table({ tableColumns: [{ size: { type: TableSizeType.SPECIFIED, width: { v: 200 } } }] }) })],
    }))).not.toBe(digestTableCell(withNested))
    expect(digestTableCell(tableCellUnit({
      tables: [tableUnit({ rows: [rowUnit({ row: row({ trHeight: { val: { v: 30 }, hRule: TableRowHeightRule.AUTO } }) })] })],
    }))).not.toBe(digestTableCell(withNested))
  })

  it('ignores table ids and the structural tableRows/tableCells containers', () => {
    const base = digestTableCell(tableCellUnit({ tables: [tableUnit()] }))
    const renamed = tableUnit({
      table: table({ tableId: 'table-other', tableRows: [row()] }),
      rows: [rowUnit({ row: row({ tableCells: [{ tcFitText: BooleanNumber.TRUE }] }) })],
    })
    expect(digestTableCell(tableCellUnit({ tables: [renamed] }))).toBe(base)
  })

  it('is deterministic across equivalent nested content', () => {
    const first = digestTableCell(tableCellUnit({ tables: [tableUnit()] }))
    const second = digestTableCell(tableCellUnit({ tables: [tableUnit()] }))
    expect(first).toBe(second)
    expect(first).toMatch(/^[0-9a-f]{64}$/u)
  })
})

describe('digestSheetCell', () => {
  it('changes for value, type, formula, hyperlink, and style changes', () => {
    const base = digestSheetCell(sheetCell())
    expect(digestSheetCell(sheetCell({ v: 'world' }))).not.toBe(base)
    expect(digestSheetCell(sheetCell({ t: CellValueType.NUMBER, v: 1 }))).not.toBe(base)
    expect(digestSheetCell(sheetCell({ f: '=B1' }))).not.toBe(base)
    expect(digestSheetCell(sheetCell({ s: { ff: 'Times New Roman' } }))).not.toBe(base)
    expect(digestSheetCell(sheetCell({ s: 'style-2' }))).not.toBe(base)
    const body = hyperlinkBody()
    expect(digestSheetCell(sheetCell({ p: hyperlinkDocument(body) }))).toBe(base)
    expect(digestSheetCell(sheetCell({ p: hyperlinkDocument(hyperlinkBody({ dataStream: 'other\r' })) }))).not.toBe(base)
    expect(digestSheetCell(sheetCell({
      p: hyperlinkDocument(hyperlinkBody({
        customRanges: [{
          startIndex: 0,
          endIndex: 3,
          rangeId: 'hyperlink-1',
          rangeType: CustomRangeType.HYPERLINK,
          properties: { url: 'https://example.org' },
        }],
      })),
    }))).not.toBe(base)
  })

  it('ignores hyperlink document ids and positional fields', () => {
    const shifted = hyperlinkDocument(hyperlinkBody({
      textRuns: [run(50, 54, 'link-style', { ff: 'Arial' })],
      paragraphs: [paragraph({
        startIndex: 99,
        paragraphId: 'para_other',
        styleId: 'LinkStyle',
        paragraphStyle: { textStyle: { ff: 'Arial' } },
        bullet: { listType: 'bullet', listId: 'list-1', nestingLevel: 0 },
      })],
      customRanges: [{
        startIndex: 10,
        endIndex: 13,
        rangeId: 'hyperlink-other',
        rangeType: CustomRangeType.HYPERLINK,
        properties: { url: 'https://example.com' },
      }],
    }), 'doc-2')
    expect(digestSheetCell(sheetCell({ p: shifted }))).toBe(digestSheetCell(sheetCell()))
  })

  it('treats a missing hyperlink like a null one and canonicalizes object key order', () => {
    const missing = sheetCell()
    delete missing.p
    expect(digestSheetCell(missing)).toBe(digestSheetCell(sheetCell({ p: null })))
    expect(digestSheetCell(sheetCell({ s: { fs: 12, ff: 'Arial' } }))).toBe(digestSheetCell(sheetCell()))
  })

  it('projects sparse hyperlink documents and ranges without falling back to undefined', () => {
    const noBody: IDocumentData = { id: 'doc-no-body', documentStyle: {} }
    expect(digestSheetCell(sheetCell({ p: noBody }))).toMatch(/^[0-9a-f]{64}$/u)
    const sparseBody = hyperlinkBody({
      textRuns: [{ st: 0, ed: 4 }],
      paragraphs: [{ startIndex: 4, paragraphId: 'para_sparse' }],
      customRanges: [{
        startIndex: 0,
        endIndex: 1,
        rangeId: 'range-sparse',
        rangeType: CustomRangeType.HYPERLINK,
        wholeEntity: true,
      }],
    })
    expect(digestSheetCell(sheetCell({ p: hyperlinkDocument(sparseBody) }))).toMatch(/^[0-9a-f]{64}$/u)
  })

  it('covers non-string values and an empty cell deterministically', () => {
    expect(digestSheetCell({ v: true, t: CellValueType.BOOLEAN })).toMatch(/^[0-9a-f]{64}$/u)
    expect(digestSheetCell({ v: 0, t: CellValueType.NUMBER })).toMatch(/^[0-9a-f]{64}$/u)
    expect(digestSheetCell({})).toMatch(/^[0-9a-f]{64}$/u)
    expect(digestSheetCell({})).toBe(digestSheetCell({ v: null }))
  })
})
