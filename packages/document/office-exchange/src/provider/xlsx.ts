/**
 * XLSX mapping between OOXML workbooks and Univer `IWorkbookData`.
 *
 * SUPPORTED: sheet order and names, string/number/boolean/date values, A1
 * formulas with cached results, number formats, bold/italic/underline, font
 * family/size/ARGB color, solid fills, cell borders, merged ranges, row
 * heights, and column widths.
 *
 * UNSUPPORTED: charts, drawings, images, comments, macros, PivotTables,
 * conditional formatting, data validation, theme-derived colors, array and
 * shared-formula metadata, hyperlink targets, error values (read as text), tab
 * colors, freeze panes, print settings, and workbook metadata. A read/write
 * round-trip does not preserve these.
 */
import ExcelJS, {
  type BorderStyle,
  type Borders,
  type Cell as ExcelCell,
  type CellValue as ExcelCellValue,
  type Color as ExcelColor,
  type Font as ExcelFont,
} from 'exceljs'
import type {
  BooleanNumber,
  BorderStyleTypes,
  CellValueType,
  DateSystem,
  IBorderData,
  IColorStyle,
  ICellData,
  IColumnData,
  IRowData,
  IStyleData,
  IWorkbookData,
  IWorksheetData,
  IRange,
  LocaleType,
  TextDecoration,
} from '@univerjs/core'
import { type OfficeExchangeConfig, resolveConfig } from '../config.ts'
import { OfficeExchangeError } from '../errors.ts'

const UNIVER_APP_VERSION = '1.4.0'
const BOOLEAN_FALSE: BooleanNumber = 0
const BOOLEAN_TRUE: BooleanNumber = 1
const CELL_STRING: CellValueType = 1
const CELL_NUMBER: CellValueType = 2
const CELL_BOOLEAN: CellValueType = 3
const DATE_1900 = 'date1900' as DateSystem
const DATE_1904 = 'date1904' as DateSystem
const LOCALE_EN_US = 'enUS' as LocaleType
const TEXT_DECORATION_SINGLE: TextDecoration = 12
const TEXT_DECORATION_DOUBLE: TextDecoration = 10
const DETERMINISTIC_DATE = new Date('2000-01-01T00:00:00.000Z')

const BORDER_STYLE_BY_EXCEL: Readonly<Record<BorderStyle, BorderStyleTypes>> = {
  thin: 1,
  hair: 2,
  dotted: 3,
  dashed: 4,
  dashDot: 5,
  dashDotDot: 6,
  double: 7,
  medium: 8,
  mediumDashed: 9,
  mediumDashDot: 10,
  mediumDashDotDot: 11,
  slantDashDot: 12,
  thick: 13,
}

const EXCEL_BORDER_BY_STYLE: Readonly<Record<number, BorderStyle>> = {
  1: 'thin',
  2: 'hair',
  3: 'dotted',
  4: 'dashed',
  5: 'dashDot',
  6: 'dashDotDot',
  7: 'double',
  8: 'medium',
  9: 'mediumDashed',
  10: 'mediumDashDot',
  11: 'mediumDashDotDot',
  12: 'slantDashDot',
  13: 'thick',
}

interface ReadValue {
  /** Cell value in the Univer model. */
  readonly v: NonNullable<ICellData['v']>
  /** Univer value classification for `v`. */
  readonly t: NonNullable<ICellData['t']>
}

/**
 * Read `.xlsx` bytes into a Univer workbook snapshot. Values, formulas, number
 * formats, fonts, fills, borders, merged ranges, and row/column sizes are
 * mapped; the unsupported features listed in the module note are not.
 * @param bytes - Complete workbook bytes.
 * @param config - Bounds applied before parsing; defaults when omitted.
 * @returns Workbook snapshot with inline styles and `=`-prefixed formulas.
 * @throws {OfficeExchangeError} `too-large` when a bound is exceeded, `conversion-failed` when the input is not a readable workbook.
 */
export async function readWorkbook(bytes: Uint8Array, config?: OfficeExchangeConfig): Promise<IWorkbookData> {
  const bounds = resolveConfig(config ?? {})
  if (bytes.byteLength > bounds.maxFileBytes) {
    throw new OfficeExchangeError(
      'too-large',
      `Workbook is ${bytes.byteLength} bytes; the limit is ${bounds.maxFileBytes}`,
    )
  }
  const workbook = new ExcelJS.Workbook()
  try {
    await workbook.xlsx.load(Uint8Array.from(bytes).buffer)
  } catch (error) {
    throw new OfficeExchangeError('conversion-failed', `Cannot read workbook: ${messageOf(error)}`)
  }
  if (workbook.worksheets.length > bounds.maxCells) {
    throw new OfficeExchangeError(
      'too-large',
      `Workbook has ${workbook.worksheets.length} sheets; the cell limit is ${bounds.maxCells}`,
    )
  }
  const date1904 = parsed(workbook.properties.date1904) === true
  let cellCount = 0
  const sheetOrder: string[] = []
  const sheets: Record<string, Partial<IWorksheetData>> = {}
  for (const [order, worksheet] of workbook.worksheets.entries()) {
    const id = `sheet-${order + 1}`
    sheetOrder.push(id)
    sheets[id] = readWorksheet(worksheet, id, date1904, () => {
      cellCount += 1
      if (cellCount > bounds.maxCells) {
        throw new OfficeExchangeError('too-large', `Workbook has more than ${bounds.maxCells} cells`)
      }
    })
  }
  return {
    id: 'office-exchange-workbook',
    name: workbook.title,
    appVersion: UNIVER_APP_VERSION,
    locale: LOCALE_EN_US,
    dateSystem: date1904 ? DATE_1904 : DATE_1900,
    styles: {},
    sheetOrder,
    sheets,
    resources: [],
  }
}

/**
 * Write a Univer workbook snapshot to `.xlsx` bytes. The inverse of
 * `readWorkbook` for the supported features; the unsupported features listed in
 * the module note are absent from the output.
 * @param data - Workbook snapshot from `readWorkbook` or the editor.
 * @param config - Bounds validated before writing; defaults when omitted.
 * @returns Encoded workbook bytes.
 * @throws {OfficeExchangeError} `too-large` when a bound is exceeded, `conversion-failed` when serialization fails.
 */
export async function writeWorkbook(data: IWorkbookData, config?: OfficeExchangeConfig): Promise<Uint8Array> {
  const bounds = resolveConfig(config ?? {})
  if (data.sheetOrder.length > bounds.maxCells) {
    throw new OfficeExchangeError(
      'too-large',
      `Workbook has ${data.sheetOrder.length} sheets; the cell limit is ${bounds.maxCells}`,
    )
  }
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Office Exchange'
  workbook.lastModifiedBy = 'Office Exchange'
  workbook.created = DETERMINISTIC_DATE
  workbook.modified = DETERMINISTIC_DATE
  workbook.properties.date1904 = data.dateSystem === DATE_1904
  let cellCount = 0
  for (const id of data.sheetOrder) {
    const sheet = data.sheets[id]
    if (sheet === undefined) {
      throw new OfficeExchangeError('conversion-failed', `Workbook snapshot is missing sheet ${id}`)
    }
    const worksheet = workbook.addWorksheet(sheet.name ?? id)
    const cellData = sheet.cellData ?? {}
    for (const [rowKey, cells] of entriesOf(cellData)) {
      const row = Number(rowKey) + 1
      for (const [columnKey, cell] of entriesOf(cells)) {
        cellCount += 1
        if (cellCount > bounds.maxCells) {
          throw new OfficeExchangeError('too-large', `Workbook has more than ${bounds.maxCells} cells`)
        }
        writeCell(worksheet.getCell(row, Number(columnKey) + 1), cell, data)
      }
    }
    writeLayout(worksheet, sheet)
    for (const merge of sheet.mergeData ?? []) worksheet.mergeCells(mergeAddress(merge))
  }
  try {
    const buffer = await workbook.xlsx.writeBuffer()
    return new Uint8Array(buffer)
  } catch (error) {
    throw new OfficeExchangeError('conversion-failed', `Cannot write workbook: ${messageOf(error)}`)
  }
}

function readWorksheet(
  worksheet: ExcelJS.Worksheet,
  id: string,
  date1904: boolean,
  countCell: () => void,
): IWorksheetData {
  const cellData: Record<number, Record<number, ICellData>> = {}
  const rowData: Record<number, Partial<IRowData>> = {}
  const columnData: Record<number, Partial<IColumnData>> = {}
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    const rowIndex = row.number - 1
    const height = parsed(row.height)
    if (height !== undefined) rowData[rowIndex] = { h: pointsToPixels(height) }
    if (row.hidden) rowData[rowIndex] = { ...rowData[rowIndex], hd: BOOLEAN_TRUE }
    row.eachCell({ includeEmpty: false }, (cell) => {
      countCell()
      const columnIndex = Number(cell.col) - 1
      cellData[rowIndex] = cellData[rowIndex] ?? {}
      cellData[rowIndex][columnIndex] = readCell(cell, date1904)
    })
  })
  for (let column = 1; column <= worksheet.columnCount; column += 1) {
    const source = worksheet.getColumn(column)
    const width = parsed(source.width)
    if (width !== undefined) columnData[column - 1] = { w: columnPixels(width) }
    if (source.hidden) columnData[column - 1] = { ...columnData[column - 1], hd: BOOLEAN_TRUE }
  }
  const properties = worksheet.properties
  const views = parsed(worksheet.views)
  return {
    id,
    name: worksheet.name,
    tabColor: '',
    hidden: worksheet.state === 'visible' ? BOOLEAN_FALSE : BOOLEAN_TRUE,
    freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
    rowCount: worksheet.rowCount,
    columnCount: worksheet.columnCount,
    zoomRatio: 100,
    scrollTop: 0,
    scrollLeft: 0,
    defaultColumnWidth: columnPixels(parsed(properties.defaultColWidth) ?? 8.43),
    defaultRowHeight: pointsToPixels(properties.defaultRowHeight),
    mergeData: worksheet.model.merges.map(range => readMerge(worksheet, range)),
    cellData,
    rowData,
    columnData,
    rowHeader: { width: 46 },
    columnHeader: { height: 20 },
    showGridlines: parsed(views?.[0]?.showGridLines) === false ? BOOLEAN_FALSE : BOOLEAN_TRUE,
    rightToLeft: BOOLEAN_FALSE,
  }
}

function readMerge(worksheet: ExcelJS.Worksheet, range: string): IRange {
  const [startAddress, endAddress] = range.split(':') as [string, string]
  const start = worksheet.getCell(startAddress)
  const end = worksheet.getCell(endAddress)
  return {
    startRow: Number(start.row) - 1,
    startColumn: Number(start.col) - 1,
    endRow: Number(end.row),
    endColumn: Number(end.col),
  }
}

function readCell(cell: ExcelCell, date1904: boolean): ICellData {
  const data: ICellData = {}
  const style = readStyle(cell)
  if (style !== undefined) data.s = style
  const formula = parsed(cell.formula)
  if (formula !== undefined) {
    data.f = `=${formula}`
    const cached = parsed(cell.result)
    if (cached !== undefined) {
      const value = readValue(cached, date1904)
      /* v8 ignore next -- A defined cached result is a primitive, date, or error, all classified by readValue. */
      if (value !== undefined) {
        data.v = value.v
        data.t = value.t
      }
    }
    return data
  }
  const value = readValue(cell.value, date1904)
  if (value !== undefined) {
    data.v = value.v
    data.t = value.t
  }
  return data
}

function readValue(value: ExcelCellValue | undefined, date1904: boolean): ReadValue | undefined {
  if (value === null || value === undefined) return undefined
  if (value instanceof Date) return { v: dateSerial(value, date1904), t: CELL_NUMBER }
  if (typeof value === 'string') return { v: value, t: CELL_STRING }
  if (typeof value === 'number') return { v: value, t: CELL_NUMBER }
  if (typeof value === 'boolean') return { v: value, t: CELL_BOOLEAN }
  if ('richText' in value && Array.isArray(value.richText)) {
    const runs = value.richText as readonly { readonly text: string }[]
    return { v: runs.map(run => run.text).join(''), t: CELL_STRING }
  }
  if ('error' in value && typeof value.error === 'string') return { v: value.error, t: CELL_STRING }
  /* v8 ignore next -- Every non-formula object value carries text or was classified above. */
  if ('text' in value && typeof value.text === 'string') return { v: value.text, t: CELL_STRING }
  /* v8 ignore next -- Formula objects are read through their cached result, never as a cell value. */
  return undefined
}

function readStyle(cell: ExcelCell): IStyleData | undefined {
  const style: IStyleData = {}
  const font = parsed(cell.font)
  if (font !== undefined) {
    if (font.bold === true) style.bl = BOOLEAN_TRUE
    if (font.italic === true) style.it = BOOLEAN_TRUE
    const underline = parsed(font.underline)
    if (underline !== undefined && underline !== false && underline !== 'none') {
      style.ul = {
        s: BOOLEAN_TRUE,
        t: underline === 'double' || underline === 'doubleAccounting'
          ? TEXT_DECORATION_DOUBLE
          : TEXT_DECORATION_SINGLE,
      }
    }
    const size = parsed(font.size)
    if (size !== undefined) style.fs = size
    const name = parsed(font.name)
    if (name !== undefined) style.ff = name
    const color = readColor(parsed(font.color))
    if (color !== undefined) style.cl = color
  }
  const numFmt = parsed(cell.numFmt)
  if (numFmt !== undefined && numFmt !== 'General') style.n = { pattern: numFmt }
  const fill = parsed(cell.fill)
  if (fill !== undefined && fill.type === 'pattern' && fill.pattern === 'solid') {
    const color = readColor(parsed(fill.fgColor))
    if (color !== undefined) style.bg = color
  }
  const border = readBorder(parsed(cell.border))
  if (border !== undefined) style.bd = border
  return Object.keys(style).length === 0 ? undefined : style
}

function readBorder(border: Partial<Borders> | undefined): IBorderData | undefined {
  if (border === undefined) return undefined
  const data: IBorderData = {}
  for (const [side, key] of [['top', 't'], ['right', 'r'], ['bottom', 'b'], ['left', 'l']] as const) {
    const edge = border[side]
    if (edge === undefined) continue
    data[key] = {
      s: BORDER_STYLE_BY_EXCEL[edge.style as BorderStyle],
      cl: readColor(parsed(edge.color)) ?? { rgb: '#000000' },
    }
  }
  return Object.keys(data).length === 0 ? undefined : data
}

function readColor(color: Partial<ExcelColor> | undefined): IColorStyle | undefined {
  if (color === undefined) return undefined
  const argb = parsed(color.argb)
  if (argb === undefined || argb.length < 6) return undefined
  const rgb = argb.slice(-6)
  return /^[0-9a-f]{6}$/iu.test(rgb) ? { rgb: `#${rgb}` } : undefined
}

function writeCell(cell: ExcelCell, data: ICellData, workbook: IWorkbookData): void {
  if (data.f !== undefined && data.f !== null && data.f.length > 0) {
    const formula = data.f.startsWith('=') ? data.f.slice(1) : data.f
    cell.value = data.v === undefined || data.v === null
      ? { formula }
      : { formula, result: data.v }
  } else if (data.v !== undefined && data.v !== null) {
    cell.value = data.v
  }
  const style = resolveStyle(workbook, data.s)
  if (style !== undefined) writeStyle(cell, style)
}

function resolveStyle(workbook: IWorkbookData, style: ICellData['s']): IStyleData | undefined {
  if (style === undefined || style === null) return undefined
  return typeof style === 'string' ? workbook.styles[style] ?? undefined : style
}

function writeStyle(cell: ExcelCell, style: IStyleData): void {
  const font: Partial<ExcelFont> = {}
  if (style.bl === BOOLEAN_TRUE) font.bold = true
  if (style.it === BOOLEAN_TRUE) font.italic = true
  if (style.ul?.s === BOOLEAN_TRUE) {
    font.underline = style.ul.t === TEXT_DECORATION_DOUBLE ? 'double' : 'single'
  }
  if (style.fs !== undefined) font.size = style.fs
  if (style.ff !== undefined && style.ff !== null) font.name = style.ff
  const fontColor = argbOf(style.cl)
  if (fontColor !== undefined) font.color = { argb: fontColor }
  if (Object.keys(font).length > 0) cell.font = font
  if (style.n !== undefined && style.n !== null) cell.numFmt = style.n.pattern
  const background = argbOf(style.bg)
  if (background !== undefined) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: background } }
  }
  const border = writeBorder(style.bd)
  if (border !== undefined) cell.border = border
}

function writeBorder(border: IStyleData['bd']): Partial<Borders> | undefined {
  if (border === undefined || border === null) return undefined
  const result: Partial<Borders> = {}
  for (const [side, key] of [['top', 't'], ['right', 'r'], ['bottom', 'b'], ['left', 'l']] as const) {
    const edge = border[key]
    if (edge === undefined || edge === null) continue
    const style = EXCEL_BORDER_BY_STYLE[edge.s]
    if (style === undefined) continue
    result[side] = { style, color: { argb: argbOf(edge.cl) ?? 'FF000000' } }
  }
  return Object.keys(result).length === 0 ? undefined : result
}

function writeLayout(worksheet: ExcelJS.Worksheet, sheet: Partial<IWorksheetData>): void {
  if (sheet.defaultRowHeight !== undefined) {
    worksheet.properties.defaultRowHeight = pixelsToPoints(sheet.defaultRowHeight)
  }
  if (sheet.defaultColumnWidth !== undefined) {
    worksheet.properties.defaultColWidth = pixelsToColumnWidth(sheet.defaultColumnWidth)
  }
  if (sheet.showGridlines !== undefined) {
    worksheet.views = [{ state: 'normal', showGridLines: sheet.showGridlines === BOOLEAN_TRUE }]
  }
  if (sheet.hidden !== undefined) {
    worksheet.state = sheet.hidden === BOOLEAN_FALSE ? 'visible' : 'hidden'
  }
  for (const [rowKey, row] of entriesOf(sheet.rowData ?? {})) {
    const excelRow = worksheet.getRow(Number(rowKey) + 1)
    if (row.h !== undefined) excelRow.height = pixelsToPoints(row.h)
    if (row.hd === BOOLEAN_TRUE) excelRow.hidden = true
  }
  for (const [columnKey, column] of entriesOf(sheet.columnData ?? {})) {
    const excelColumn = worksheet.getColumn(Number(columnKey) + 1)
    if (column.w !== undefined) excelColumn.width = pixelsToColumnWidth(column.w)
    if (column.hd === BOOLEAN_TRUE) excelColumn.hidden = true
  }
}

function mergeAddress(range: IRange): string {
  const start = cellAddress(range.startRow + 1, range.startColumn + 1)
  const end = cellAddress(range.endRow, range.endColumn)
  return `${start}:${end}`
}

function cellAddress(row: number, column: number): string {
  let label = ''
  let value = column
  while (value > 0) {
    label = String.fromCharCode(65 + ((value - 1) % 26)) + label
    value = Math.floor((value - 1) / 26)
  }
  return `${label}${row}`
}

function argbOf(color: IColorStyle | null | undefined | void): string | undefined {
  const rgb = color?.rgb
  if (rgb === undefined || rgb === null) return undefined
  const hex = rgb.startsWith('#') ? rgb.slice(1) : rgb
  return /^[0-9a-f]{6}$/iu.test(hex) ? `FF${hex}` : undefined
}

function dateSerial(date: Date, date1904: boolean): number {
  return date.getTime() / 86_400_000 + 25_569 - (date1904 ? 1_462 : 0)
}

function pointsToPixels(points: number): number {
  return points * 96 / 72
}

function pixelsToPoints(pixels: number): number {
  return pixels * 72 / 96
}

function columnPixels(width: number): number {
  return Math.floor(width < 1 ? width * 12 : width * 7 + 5)
}

function pixelsToColumnWidth(pixels: number): number {
  return pixels < 12 ? pixels / 12 : (pixels - 5) / 7
}

/**
 * Widen a field ExcelJS declares as required but omits when a file lacks it.
 * @param value - Declared field value.
 * @returns The runtime value, including the absent case.
 */
function parsed<T>(value: T): T | undefined {
  return value
}

/**
 * Enumerate a numeric-index record with string keys.
 * @param record - Numeric-index record to enumerate.
 * @returns Key/value pairs for every present entry.
 */
function entriesOf<T>(record: { readonly [key: number]: T }): Array<[string, T]> {
  return Object.entries(record)
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
