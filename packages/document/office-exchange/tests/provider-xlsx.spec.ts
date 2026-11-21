/** XLSX provider mapping: fixture semantics and the write/read round-trip. */
import { readFile } from 'node:fs/promises'
import ExcelJS from 'exceljs'
import { expect, it, vi } from 'vitest'
import {
  BaselineOffset,
  BooleanNumber,
  CellValueType,
  DateSystem,
  HorizontalAlign,
  LocaleType,
  TextDecoration,
  VerticalAlign,
  WrapStrategy,
  WorksheetHiddenState,
  type ICellData,
  type IColumnData,
  type IRowData,
  type IStyleData,
  type IWorkbookData,
  type IWorksheetData,
} from '@univerjs/core'
import { readWorkbook, writeWorkbook } from '../src/provider/xlsx.ts'

const FIXTURE = new URL('./fixtures/roundtrip.xlsx', import.meta.url)

function sheetOf(workbook: IWorkbookData, name: string): Partial<IWorksheetData> {
  const id = workbook.sheetOrder.find(candidate => workbook.sheets[candidate]?.name === name)
  if (id === undefined) throw new Error(`missing sheet ${name}`)
  return workbook.sheets[id] ?? {}
}

function cellAt(sheet: Partial<IWorksheetData>, row: number, column: number): ICellData | undefined {
  return sheet.cellData?.[row]?.[column]
}

function styleOf(cell: ICellData | undefined): IStyleData | undefined {
  const style = cell?.s
  return style !== undefined && style !== null && typeof style !== 'string' ? style : undefined
}

function fieldStyleOf(entry: Partial<IRowData> | Partial<IColumnData> | undefined): IStyleData | undefined {
  const style = entry?.s
  return style !== undefined && style !== null && typeof style !== 'string' ? style : undefined
}

function expectSemantics(workbook: IWorkbookData): void {
  expect(workbook.sheetOrder.map(id => workbook.sheets[id]?.name)).toEqual(['Sales', 'Notes'])
  const sales = sheetOf(workbook, 'Sales')
  expect(styleOf(cellAt(sales, 0, 0))?.bl).toBe(1)
  expect(styleOf(cellAt(sales, 0, 1))?.bl).toBe(1)
  expect(styleOf(cellAt(sales, 0, 2))?.bl).toBe(1)
  for (const row of [1, 2, 3]) {
    expect(styleOf(cellAt(sales, row, 2))?.n?.pattern).toBe('$#,##0.00')
  }
  expect(cellAt(sales, 4, 1)?.f).toBe('=SUM(B2:B4)')
  expect(sales.mergeData).toContainEqual({
    startRow: 6,
    startColumn: 0,
    endRow: 7,
    endColumn: 3,
  })
  expect(cellAt(sales, 0, 0)?.v).toBe('Item')
  expect(cellAt(sales, 1, 0)?.v).toBe('Widget')
  expect(cellAt(sales, 1, 2)?.v).toBe(1200.5)
  expect(cellAt(sales, 4, 0)?.v).toBe('Total units')
}

it('reads workbook values, formulas, styles, and merges from the fixture', async () => {
  const bytes = new Uint8Array(await readFile(FIXTURE))
  expectSemantics(await readWorkbook(bytes))
})

it('preserves supported semantics through a write/read round-trip', async () => {
  const original = await readWorkbook(new Uint8Array(await readFile(FIXTURE)))
  const encoded = await writeWorkbook(original)
  expect(encoded.byteLength).toBeGreaterThan(0)
  expectSemantics(await readWorkbook(encoded))
})

it('maps dates, rich text, hyperlinks, errors, formulas, and supported styles', async () => {
  const workbook = new ExcelJS.Workbook()
  workbook.properties.date1904 = true
  const sheet = workbook.addWorksheet('Data')
  sheet.getCell('A1').value = new Date('2024-01-02T03:04:05.000Z')
  sheet.getCell('A1').numFmt = 'yyyy-mm-dd'
  sheet.getCell('A2').value = { richText: [{ text: 'rich ' }, { text: 'text' }] }
  sheet.getCell('A3').value = { text: 'link', hyperlink: 'https://example.com' }
  sheet.getCell('A4').value = { error: '#VALUE!' }
  sheet.getCell('A5').value = { formula: 'SUM(B1:B2)', result: 3 }
  sheet.getCell('B1').value = 1
  sheet.getCell('B2').value = 2
  sheet.getCell('F1').value = 'invalid color'
  sheet.getCell('F1').font = { color: { argb: 'ZZZZZZ' } }
  sheet.getCell('F2').value = 'short color'
  sheet.getCell('F2').font = { color: { argb: 'FFF' } }
  sheet.getCell('G1').value = 'invalid fill'
  sheet.getCell('G1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'ZZZZZZ' } }
  const styled = sheet.getCell('C1')
  styled.value = 'styled'
  styled.font = {
    bold: true,
    italic: true,
    underline: true,
    size: 14,
    name: 'Arial',
    color: { argb: 'FF112233' },
  }
  styled.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFAABBCC' } }
  styled.border = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'double' },
  }
  styled.numFmt = '0.00'
  sheet.getRow(3).height = 30
  sheet.getColumn(2).width = 20
  sheet.getColumn(4).width = 0.5
  sheet.getRow(4).hidden = true
  sheet.getColumn(3).hidden = true
  sheet.views = [{ state: 'normal', showGridLines: false }]
  sheet.properties.defaultColWidth = 12
  sheet.properties.defaultRowHeight = 18
  sheet.mergeCells('D1:E2')
  sheet.getCell('D1').value = 'merge'
  const snapshot = await readWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  expect(snapshot.dateSystem).toBe(DateSystem.Date1904)
  const data = sheetOf(snapshot, 'Data')
  const date = cellAt(data, 0, 0)
  expect(typeof date?.v).toBe('number')
  expect(date?.t).toBe(CellValueType.NUMBER)
  expect(styleOf(date)?.n?.pattern).toBe('yyyy-mm-dd')
  expect(cellAt(data, 1, 0)?.v).toBe('rich text')
  expect(cellAt(data, 2, 0)?.v).toBe('link')
  expect(cellAt(data, 3, 0)?.v).toBe('#VALUE!')
  expect(cellAt(data, 4, 0)?.f).toBe('=SUM(B1:B2)')
  expect(cellAt(data, 4, 0)?.v).toBe(3)
  const style = styleOf(cellAt(data, 0, 2))
  expect(style?.bl).toBe(BooleanNumber.TRUE)
  expect(style?.it).toBe(BooleanNumber.TRUE)
  expect(style?.ul?.t).toBe(TextDecoration.SINGLE)
  expect(style?.fs).toBe(14)
  expect(style?.ff).toBe('Arial')
  expect(style?.cl?.rgb).toBe('#112233')
  expect(style?.bg?.rgb).toBe('#AABBCC')
  expect(style?.bd?.t?.s).toBe(1)
  expect(style?.bd?.b?.s).toBe(7)
  expect(style?.n?.pattern).toBe('0.00')
  expect(styleOf(cellAt(data, 0, 5))?.cl).toBeUndefined()
  expect(styleOf(cellAt(data, 1, 5))?.cl).toBeUndefined()
  expect(styleOf(cellAt(data, 0, 6))?.bg).toBeUndefined()
  expect(data.rowData?.[2]?.h).toBe(40)
  expect(data.rowData?.[3]?.hd).toBe(BooleanNumber.TRUE)
  expect(data.columnData?.[1]?.w).toBe(145)
  expect(data.columnData?.[2]?.hd).toBe(BooleanNumber.TRUE)
  expect(data.columnData?.[3]?.w).toBe(6)
  expect(data.defaultColumnWidth).toBe(89)
  expect(data.defaultRowHeight).toBe(24)
  expect(data.showGridlines).toBe(BooleanNumber.FALSE)
  expect(data.mergeData).toContainEqual({
    startRow: 0,
    startColumn: 3,
    endRow: 2,
    endColumn: 5,
  })
})

it('writes style ids, cached formula results, merges, and layout', async () => {
  const snapshot: IWorkbookData = {
    id: 'synthetic',
    name: '',
    appVersion: '1.4.0',
    locale: LocaleType.EN_US,
    dateSystem: DateSystem.Date1904,
    styles: {
      full: {
        bl: BooleanNumber.TRUE,
        it: BooleanNumber.TRUE,
        ul: { s: BooleanNumber.TRUE, t: TextDecoration.DOUBLE },
        fs: 12,
        ff: 'Arial',
        cl: { rgb: '#112233' },
        n: { pattern: '0.00' },
        bg: { rgb: '#AABBCC' },
        bd: {
          t: { s: 1, cl: { rgb: '#000000' } },
          r: { s: 0, cl: { rgb: '#000000' } },
          b: null,
          l: { s: 7, cl: { rgb: '#FFFFFF' } },
        },
      },
    },
    sheetOrder: ['synthetic', 'plain', 'nameless'],
    sheets: {
      nameless: { id: 'nameless' },
      synthetic: {
        id: 'synthetic',
        name: 'Synthetic',
        cellData: {
          0: {
            0: { v: 'styled', s: 'full' },
            1: { f: '=SUM(C1:C2)', v: 3, t: CellValueType.NUMBER },
            2: { v: 1 },
            3: {
              v: 'inline',
              s: {
                bl: BooleanNumber.TRUE,
                ul: { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE },
                cl: { rgb: '112233' },
                bg: { rgb: 'rgb(1,2,3)' },
              },
            },
            4: { v: 'missing', s: 'missing' },
            5: { v: 'null', s: null },
            6: { v: false, t: CellValueType.BOOLEAN },
            7: { f: '=1+1', s: { n: { pattern: '$#,##0.00' } } },
            8: { f: 'SUM(C1:C2)' },
            9: { v: 'default border', s: { bd: { t: { s: 1, cl: {} } } } },
          },
          1: { 2: { v: 2 }, 0: { s: { bd: {} } } },
          2: { 0: { s: { bd: { b: null } } } },
          3: { 0: { s: { bd: { t: { s: 0, cl: { rgb: '#000000' } } } } } },
        },
        mergeData: [{ startRow: 0, startColumn: 26, endRow: 2, endColumn: 28 }],
        rowData: { 0: { h: 20 }, 1: { hd: BooleanNumber.TRUE }, 2: { h: 30, hd: BooleanNumber.TRUE } },
        columnData: {
          0: { w: 70 },
          1: { hd: BooleanNumber.TRUE },
          2: { w: 90, hd: BooleanNumber.TRUE },
          3: { w: 6 },
        },
        defaultRowHeight: 24,
        defaultColumnWidth: 64,
        showGridlines: BooleanNumber.FALSE,
        hidden: BooleanNumber.TRUE,
      },
      plain: {
        id: 'plain',
        name: 'Plain',
        showGridlines: BooleanNumber.TRUE,
        hidden: BooleanNumber.FALSE,
      },
    },
  }
  const snapshotRead = await readWorkbook(new Uint8Array(await writeWorkbook(snapshot)))
  expect(snapshotRead.sheetOrder.map(id => snapshotRead.sheets[id]?.name))
    .toEqual(['Synthetic', 'Plain', 'nameless'])
  const data = sheetOf(snapshotRead, 'Synthetic')
  const style = styleOf(cellAt(data, 0, 0))
  expect(style?.bl).toBe(BooleanNumber.TRUE)
  expect(style?.ul?.t).toBe(TextDecoration.DOUBLE)
  expect(style?.bd?.t?.s).toBe(1)
  expect(style?.bd?.l?.s).toBe(7)
  expect(cellAt(data, 0, 1)?.f).toBe('=SUM(C1:C2)')
  expect(cellAt(data, 0, 1)?.v).toBe(3)
  const inline = styleOf(cellAt(data, 0, 3))
  expect(inline?.cl?.rgb).toBe('#112233')
  expect(inline?.bg).toBeUndefined()
  expect(styleOf(cellAt(data, 0, 4))).toBeUndefined()
  expect(cellAt(data, 0, 6)?.v).toBe(false)
  expect(cellAt(data, 0, 7)?.f).toBe('=1+1')
  expect(styleOf(cellAt(data, 0, 7))?.n?.pattern).toBe('$#,##0.00')
  expect(cellAt(data, 0, 8)?.f).toBe('=SUM(C1:C2)')
  expect(styleOf(cellAt(data, 0, 9))?.bd?.t?.cl?.rgb).toBe('#000000')
  expect(data.mergeData).toContainEqual({
    startRow: 0,
    startColumn: 26,
    endRow: 2,
    endColumn: 28,
  })
  expect(data.rowData?.[0]?.h).toBe(20)
  expect(data.rowData?.[1]?.hd).toBe(BooleanNumber.TRUE)
  expect(data.rowData?.[2]?.h).toBe(30)
  expect(data.columnData?.[0]?.w).toBe(70)
  expect(data.columnData?.[1]?.hd).toBe(BooleanNumber.TRUE)
  expect(data.columnData?.[2]?.w).toBe(90)
  expect(data.columnData?.[3]?.w).toBe(6)
  expect(data.defaultRowHeight).toBe(24)
  expect(data.defaultColumnWidth).toBe(64)
  expect(data.showGridlines).toBe(BooleanNumber.FALSE)
  expect(data.hidden).toBe(BooleanNumber.TRUE)
  const plain = sheetOf(snapshotRead, 'Plain')
  expect(plain.showGridlines).toBe(BooleanNumber.TRUE)
  expect(plain.hidden).toBe(BooleanNumber.FALSE)
})

it('maps a 1900-system date to a serial number', async () => {
  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('Dated').getCell('A1').value = new Date('2024-01-02T03:04:05.000Z')
  const snapshot = await readWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  expect(typeof cellAt(sheetOf(snapshot, 'Dated'), 0, 0)?.v).toBe('number')
})

it('maps alignment, strike, baseline, view state, and row and column styles', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Formatting')
  const aligned = sheet.getCell('A1')
  aligned.value = 'aligned'
  aligned.alignment = {
    horizontal: 'center',
    vertical: 'top',
    wrapText: true,
    shrinkToFit: true,
    textRotation: 45,
  }
  aligned.font = { strike: true, vertAlign: 'superscript' }
  sheet.getCell('A2').value = 'rotated'
  sheet.getCell('A2').alignment = { horizontal: 'right', vertical: 'bottom', textRotation: -30 }
  sheet.getCell('A3').value = 'vertical'
  sheet.getCell('A3').alignment = { textRotation: 'vertical' }
  sheet.getCell('A4').value = 'subscript'
  sheet.getCell('A4').font = { vertAlign: 'subscript' }
  sheet.getCell('A5').value = 'row style'
  sheet.getRow(5).font = { bold: true }
  sheet.getColumn(2).alignment = { horizontal: 'center' }
  sheet.views = [{
    state: 'frozen',
    xSplit: 2,
    ySplit: 1,
    topLeftCell: 'C2',
    showGridLines: false,
    zoomScale: 150,
    rightToLeft: true,
    showRowColHeaders: false,
  }]
  sheet.properties.tabColor = { argb: 'FF336699' }
  const snapshot = await readWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  assertFormattingSheet(sheetOf(snapshot, 'Formatting'))
  const reread = await readWorkbook(new Uint8Array(await writeWorkbook(snapshot)))
  assertFormattingSheet(sheetOf(reread, 'Formatting'))
})

function assertFormattingSheet(data: Partial<IWorksheetData>): void {
  const aligned = styleOf(cellAt(data, 0, 0))
  expect(aligned?.ht).toBe(HorizontalAlign.CENTER)
  expect(aligned?.vt).toBe(VerticalAlign.TOP)
  expect(aligned?.tb).toBe(WrapStrategy.WRAP)
  expect(aligned?.stf).toBe(BooleanNumber.TRUE)
  expect(aligned?.tr).toEqual({ a: 45 })
  expect(aligned?.st?.s).toBe(BooleanNumber.TRUE)
  expect(aligned?.va).toBe(BaselineOffset.SUPERSCRIPT)
  const rotated = styleOf(cellAt(data, 1, 0))
  expect(rotated?.ht).toBe(HorizontalAlign.RIGHT)
  expect(rotated?.vt).toBe(VerticalAlign.BOTTOM)
  expect(rotated?.tr).toEqual({ a: -30 })
  expect(styleOf(cellAt(data, 2, 0))?.tr).toEqual({ a: 0, v: BooleanNumber.TRUE })
  expect(styleOf(cellAt(data, 3, 0))?.va).toBe(BaselineOffset.SUBSCRIPT)
  expect(fieldStyleOf(data.rowData?.[4])?.bl).toBe(BooleanNumber.TRUE)
  expect(fieldStyleOf(data.columnData?.[1])?.ht).toBe(HorizontalAlign.CENTER)
  expect(data.freeze).toEqual({ xSplit: 2, ySplit: 1, startRow: 1, startColumn: 2 })
  expect(data.zoomRatio).toBe(150)
  expect(data.rightToLeft).toBe(BooleanNumber.TRUE)
  expect(data.rowHeader?.hidden).toBe(BooleanNumber.TRUE)
  expect(data.columnHeader?.hidden).toBe(BooleanNumber.TRUE)
  expect(data.showGridlines).toBe(BooleanNumber.FALSE)
  expect(data.tabColor).toBe('#336699')
}

it('reads frozen panes that pin only one axis, and writes them back', async () => {
  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('Cols').views = [{ state: 'frozen', xSplit: 2 }]
  workbook.addWorksheet('Rows').views = [{ state: 'frozen', ySplit: 1 }]
  const snapshot = await readWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  expect(sheetOf(snapshot, 'Cols').freeze).toEqual({ xSplit: 2, ySplit: 0, startRow: 0, startColumn: 2 })
  expect(sheetOf(snapshot, 'Rows').freeze).toEqual({ xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 })
  const reread = await readWorkbook(new Uint8Array(await writeWorkbook(snapshot)))
  expect(sheetOf(reread, 'Cols').freeze).toEqual({ xSplit: 2, ySplit: 0, startRow: 0, startColumn: 2 })
  expect(sheetOf(reread, 'Rows').freeze).toEqual({ xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 })
})

it('maps the very-hidden sheet state through a write/read round trip', async () => {
  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('Shown')
  workbook.addWorksheet('Hidden').state = 'veryHidden'
  const snapshot = await readWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  expect(sheetOf(snapshot, 'Hidden').hidden).toBe(WorksheetHiddenState.VERY_HIDDEN)
  const reread = await readWorkbook(new Uint8Array(await writeWorkbook(snapshot)))
  expect(sheetOf(reread, 'Hidden').hidden).toBe(WorksheetHiddenState.VERY_HIDDEN)
})

it('omits styles that map to no ExcelJS field and invalid tab colors', async () => {
  const snapshot: IWorkbookData = {
    id: 'empty-formatting',
    name: '',
    appVersion: '1.4.0',
    locale: LocaleType.EN_US,
    dateSystem: DateSystem.Date1900,
    styles: {},
    sheetOrder: ['empty', 'header'],
    sheets: {
      empty: {
        id: 'empty',
        name: 'Empty',
        tabColor: 'not-a-color',
        cellData: {
          0: {
            0: { v: 'no alignment', s: { ht: 0, vt: 0, tb: 2, va: null } },
            1: { v: 'no rotation', s: { tr: null } },
            2: { v: 'empty style', s: {} },
          },
        },
        rowData: { 0: { s: {} } },
        columnData: { 0: { s: {} }, 3: { hd: BooleanNumber.TRUE } },
      },
      header: {
        id: 'header',
        name: 'Header',
        columnHeader: { height: 20, hidden: BooleanNumber.TRUE },
      },
    },
  }
  const reread = await readWorkbook(new Uint8Array(await writeWorkbook(snapshot)))
  const empty = sheetOf(reread, 'Empty')
  expect(empty.tabColor).toBe('')
  expect(styleOf(cellAt(empty, 0, 0))).toBeUndefined()
  expect(styleOf(cellAt(empty, 0, 1))).toBeUndefined()
  expect(styleOf(cellAt(empty, 0, 2))).toBeUndefined()
  const header = sheetOf(reread, 'Header')
  expect(header.rowHeader?.hidden).toBe(BooleanNumber.TRUE)
  expect(header.columnHeader?.hidden).toBe(BooleanNumber.TRUE)
})

it('rejects a workbook with more sheets than the configured cell limit', async () => {
  const bytes = new Uint8Array(await readFile(FIXTURE))
  await expect(readWorkbook(bytes, {
    maxFileBytes: 1_000_000,
    maxCells: 1,
    maxConversionMs: 1_000,
    commandTimeoutMs: 5_000,
    maxSourceAnchors: 100_000,
    maxSourceAnchorBytes: 8 * 1024 * 1024,
  })).rejects.toMatchObject({ code: 'too-large' })
})

it('rejects a snapshot with more sheets than the configured cell limit', async () => {
  const workbook = await readWorkbook(new Uint8Array(await readFile(FIXTURE)))
  await expect(writeWorkbook(workbook, {
    maxFileBytes: 1_000_000,
    maxCells: 1,
    maxConversionMs: 1_000,
    commandTimeoutMs: 5_000,
    maxSourceAnchors: 100_000,
    maxSourceAnchorBytes: 8 * 1024 * 1024,
  })).rejects.toMatchObject({ code: 'too-large' })
})

it('rejects a snapshot whose sheet order names a missing sheet', async () => {
  const snapshot: IWorkbookData = {
    id: 'missing',
    name: '',
    appVersion: '1.4.0',
    locale: LocaleType.EN_US,
    styles: {},
    sheetOrder: ['absent'],
    sheets: {},
  }
  await expect(writeWorkbook(snapshot)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('reports workbook serialization failures', async () => {
  const workbook = await readWorkbook(new Uint8Array(await readFile(FIXTURE)))
  const spy = vi.spyOn(ExcelJS.Workbook.prototype, 'xlsx', 'get').mockImplementation(() => {
    throw new Error('write rejected')
  })
  try {
    await expect(writeWorkbook(workbook)).rejects.toMatchObject({ code: 'conversion-failed' })
  } finally {
    spy.mockRestore()
  }
})

it('reports non-Error serialization failures with their text', async () => {
  const workbook = await readWorkbook(new Uint8Array(await readFile(FIXTURE)))
  const spy = vi.spyOn(ExcelJS.Workbook.prototype, 'xlsx', 'get').mockImplementation(() => {
    throw { toString: () => 'write rejected' }
  })
  try {
    const error: unknown = await writeWorkbook(workbook).catch((cause: unknown) => cause)
    expect(error).toMatchObject({ code: 'conversion-failed' })
    expect(String(error)).toContain('write rejected')
  } finally {
    spy.mockRestore()
  }
})

it('rejects bytes above the configured file bound', async () => {
  await expect(readWorkbook(new Uint8Array(9), {
    maxFileBytes: 8,
    maxCells: 100,
    maxConversionMs: 1_000,
    commandTimeoutMs: 5_000,
    maxSourceAnchors: 100_000,
    maxSourceAnchorBytes: 8 * 1024 * 1024,
  })).rejects.toMatchObject({ code: 'too-large' })
})

it('rejects a workbook above the configured cell bound', async () => {
  const bytes = new Uint8Array(await readFile(FIXTURE))
  await expect(readWorkbook(bytes, {
    maxFileBytes: 1_000_000,
    maxCells: 6,
    maxConversionMs: 1_000,
    commandTimeoutMs: 5_000,
    maxSourceAnchors: 100_000,
    maxSourceAnchorBytes: 8 * 1024 * 1024,
  })).rejects.toMatchObject({ code: 'too-large' })
})

it('rejects a snapshot above the configured cell bound before writing', async () => {
  const workbook = await readWorkbook(new Uint8Array(await readFile(FIXTURE)))
  await expect(writeWorkbook(workbook, {
    maxFileBytes: 1_000_000,
    maxCells: 6,
    maxConversionMs: 1_000,
    commandTimeoutMs: 5_000,
    maxSourceAnchors: 100_000,
    maxSourceAnchorBytes: 8 * 1024 * 1024,
  })).rejects.toMatchObject({ code: 'too-large' })
})

it('rejects unreadable bytes as conversion failures', async () => {
  await expect(readWorkbook(new Uint8Array([0, 1, 2, 3]))).rejects.toMatchObject({
    name: 'OfficeExchangeError',
    code: 'conversion-failed',
  })
})
