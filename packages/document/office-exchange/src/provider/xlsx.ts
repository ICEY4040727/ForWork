/**
 * XLSX mapping between OOXML workbooks and Univer `IWorkbookData`.
 *
 * SUPPORTED: sheet order and names, string/number/boolean/date values, A1
 * formulas with cached results, number formats, bold/italic/underline/strike,
 * vertical baseline, font family/size/ARGB color, horizontal and vertical
 * alignment, wrap and shrink-to-fit, text rotation, solid fills, cell borders,
 * row and column styles, merged ranges, row heights, column widths, frozen
 * panes, zoom, right-to-left view, header visibility, tab colors, the
 * very-hidden sheet state, conditional-formatting rules, data-validation
 * rules, cell hyperlinks, cell images, and worksheet background images.
 *
 * UNSUPPORTED: charts, non-image drawings, comments, macros, PivotTables,
 * theme-derived colors, array and shared-formula metadata, error values (read
 * as text), print settings, row and column outline levels, and workbook
 * metadata. A read/write round-trip does not preserve these.
 */
import ExcelJS, {
  type Alignment as ExcelAlignment,
  type BorderStyle,
  type Borders,
  type Cell as ExcelCell,
  type CellValue as ExcelCellValue,
  type Color as ExcelColor,
  type Fill as ExcelFill,
  type Font as ExcelFont,
  type Image as ExcelImage,
  type Style as ExcelStyle,
} from 'exceljs'
import type {
  BaselineOffset,
  BooleanNumber,
  BorderStyleTypes,
  CellValueType,
  DateSystem,
  HorizontalAlign,
  IBorderData,
  IColorStyle,
  ICellData,
  IColumnData,
  IFreeze,
  IRowData,
  IStyleData,
  ITextRotation,
  IWorkbookData,
  IWorksheetBackgroundImage,
  IWorksheetData,
  IRange,
  LocaleType,
  TextDecoration,
  VerticalAlign,
  WrapStrategy,
  WorksheetHiddenState,
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
const HORIZONTAL_LEFT = 1 as HorizontalAlign
const HORIZONTAL_CENTER = 2 as HorizontalAlign
const HORIZONTAL_RIGHT = 3 as HorizontalAlign
const HORIZONTAL_JUSTIFIED = 4 as HorizontalAlign
const HORIZONTAL_DISTRIBUTED = 6 as HorizontalAlign
const VERTICAL_TOP = 1 as VerticalAlign
const VERTICAL_MIDDLE = 2 as VerticalAlign
const VERTICAL_BOTTOM = 3 as VerticalAlign
const WRAP_STRATEGY_WRAP = 3 as WrapStrategy
const BASELINE_SUBSCRIPT = 2 as BaselineOffset
const BASELINE_SUPERSCRIPT = 3 as BaselineOffset
const VERY_HIDDEN = 2 as WorksheetHiddenState
const VERTICAL_ROTATION = 'vertical'
const DEFAULT_ZOOM_RATIO = 100
const DETERMINISTIC_DATE = new Date('2000-01-01T00:00:00.000Z')
const EMU_PER_PIXEL = 9525
const DEFAULT_COLUMN_PIXELS = 64
const DEFAULT_ROW_PIXELS = 20
const HYPERLINK_RANGE_TYPE = 0
const DRAWING_TYPE_IMAGE = 0
const IMAGE_SOURCE_BASE64 = 'BASE64'

/** Univer resource names for the sheet feature plugins that persist these constructs. */
const CONDITIONAL_FORMATTING_RESOURCE = 'SHEET_CONDITIONAL_FORMATTING_PLUGIN'
const DATA_VALIDATION_RESOURCE = 'SHEET_DATA_VALIDATION_PLUGIN'
const DRAWING_RESOURCE = 'SHEET_DRAWING_PLUGIN'

const CF_HIGHLIGHT = 'highlightCell'
const CF_DATA_BAR = 'dataBar'
const CF_COLOR_SCALE = 'colorScale'
const CF_ICON_SET = 'iconSet'

const CF_SUB_FORMULA = 'formula'
const CF_SUB_NUMBER = 'number'
const CF_SUB_TEXT = 'text'
const CF_SUB_RANK = 'rank'
const CF_SUB_AVERAGE = 'average'
const CF_SUB_UNIQUE = 'uniqueValues'
const CF_SUB_DUPLICATE = 'duplicateValues'
const CF_SUB_TIME_PERIOD = 'timePeriod'

const UNIVER_CF_OPERATOR_BY_EXCEL: Readonly<Record<string, string>> = {
  equal: 'equal',
  notEqual: 'notEqual',
  greaterThan: 'greaterThan',
  greaterThanOrEqual: 'greaterThanOrEqual',
  lessThan: 'lessThan',
  lessThanOrEqual: 'lessThanOrEqual',
  between: 'between',
  notBetween: 'notBetween',
}

const EXCEL_OPERATOR_TOKEN: Readonly<Record<string, string>> = {
  equal: '=',
  notEqual: '<>',
  greaterThan: '>',
  greaterThanOrEqual: '>=',
  lessThan: '<',
  lessThanOrEqual: '<=',
}

const UNIVER_TIME_PERIOD_BY_EXCEL: Readonly<Record<string, string>> = {
  today: 'today',
  yesterday: 'yesterday',
  tomorrow: 'tomorrow',
  last7Days: 'last7Days',
  thisMonth: 'thisMonth',
  lastMonth: 'lastMonth',
  nextMonth: 'nextMonth',
  thisWeek: 'thisWeek',
  lastWeek: 'lastWeek',
  nextWeek: 'nextWeek',
}

const EXCEL_TIME_PERIOD_BY_UNIVER: Readonly<Record<string, string>> = {
  today: 'today',
  yesterday: 'yesterday',
  tomorrow: 'tomorrow',
  last7Days: 'last7Days',
  thisMonth: 'thisMonth',
  lastMonth: 'lastMonth',
  nextMonth: 'nextMonth',
  thisWeek: 'thisWeek',
  lastWeek: 'lastWeek',
  nextWeek: 'nextWeek',
}

const EXCEL_ICON_SET_NAMES: ReadonlySet<string> = new Set([
  '3Arrows', '3ArrowsGray', '3Flags', '3Signs', '3Stars', '3Symbols', '3Symbols2',
  '3TrafficLights1', '3TrafficLights2', '3Triangles', '4Arrows', '4ArrowsGray',
  '4Rating', '4RedToBlack', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Boxes',
  '5Quarters', '5Rating',
])

/** One persisted conditional-formatting rule in the plugin resource JSON. */
interface ConditionalFormattingResourceRule {
  readonly cfId: string
  readonly ranges: readonly IRange[]
  readonly stopIfTrue: boolean
  readonly rule: ConditionalFormattingResourceConfig
}

/** The rule body variants Univer's conditional-formatting plugin consumes. */
type ConditionalFormattingResourceConfig =
  | {
    readonly type: typeof CF_HIGHLIGHT
    readonly subType: typeof CF_SUB_FORMULA
    readonly value: string
    readonly style: IStyleData
  }
  | {
    readonly type: typeof CF_HIGHLIGHT
    readonly subType: typeof CF_SUB_NUMBER
    readonly operator: string
    readonly value: number | readonly [number, number]
    readonly style: IStyleData
  }
  | {
    readonly type: typeof CF_HIGHLIGHT
    readonly subType: typeof CF_SUB_TEXT
    readonly operator: string
    readonly value: string
    readonly style: IStyleData
  }
  | {
    readonly type: typeof CF_HIGHLIGHT
    readonly subType: typeof CF_SUB_RANK
    readonly isBottom: boolean
    readonly isPercent: boolean
    readonly value: number
    readonly style: IStyleData
  }
  | {
    readonly type: typeof CF_HIGHLIGHT
    readonly subType: typeof CF_SUB_AVERAGE
    readonly operator: string
    readonly style: IStyleData
  }
  | {
    readonly type: typeof CF_HIGHLIGHT
    readonly subType: typeof CF_SUB_UNIQUE | typeof CF_SUB_DUPLICATE
    readonly style: IStyleData
  }
  | {
    readonly type: typeof CF_HIGHLIGHT
    readonly subType: typeof CF_SUB_TIME_PERIOD
    readonly operator: string
    readonly style: IStyleData
  }
  | {
    readonly type: typeof CF_COLOR_SCALE
    readonly config: readonly {
      readonly index: number
      readonly color: string
      readonly value: ConditionalFormattingValue
    }[]
  }
  | {
    readonly type: typeof CF_DATA_BAR
    readonly isShowValue: boolean
    readonly config: {
      readonly min: ConditionalFormattingValue
      readonly max: ConditionalFormattingValue
      readonly isGradient: boolean
      readonly positiveColor: string
      readonly nativeColor: string
    }
  }
  | {
    readonly type: typeof CF_ICON_SET
    readonly isShowValue: boolean
    readonly config: readonly {
      readonly operator: string
      readonly value: ConditionalFormattingValue
      readonly iconType: string
      readonly iconId: string
    }[]
  }

/** One threshold in a color scale, data bar, or icon set rule. */
interface ConditionalFormattingValue {
  readonly type: string
  readonly value?: number | string
}

/** One persisted data-validation rule in the plugin resource JSON. */
interface DataValidationResourceRule {
  readonly uid: string
  readonly type: string
  readonly operator?: string
  readonly formula1?: string
  readonly formula2?: string
  readonly allowBlank?: boolean
  readonly showDropDown?: boolean
  readonly showErrorMessage?: boolean
  readonly errorStyle?: number
  readonly error?: string
  readonly errorTitle?: string
  readonly showInputMessage?: boolean
  readonly prompt?: string
  readonly promptTitle?: string
  readonly ranges: readonly IRange[]
}

/** One cell offset inside a drawing anchor, in worksheet layout pixels. */
interface DrawingCellPosition {
  readonly column: number
  readonly columnOffset: number
  readonly row: number
  readonly rowOffset: number
}

/** A drawing's anchor in worksheet coordinates. */
interface DrawingTransform {
  readonly from: DrawingCellPosition
  readonly to: DrawingCellPosition
}

/** One persisted cell image in the sheet-drawing plugin resource JSON. */
interface DrawingImageResource {
  readonly unitId: string
  readonly subUnitId: string
  readonly drawingId: string
  readonly drawingType: number
  readonly imageSourceType: string
  readonly source: string
  readonly sheetTransform: DrawingTransform
  readonly axisAlignSheetTransform: DrawingTransform
  readonly xlsx?: {
    readonly extension: string
    readonly editAs?: string
    readonly ext?: { readonly width: number; readonly height: number }
    readonly hyperlink?: { readonly hyperlink: string; readonly tooltip?: string }
  }
}

/** The per-sheet map the sheet-drawing plugin persists. */
type DrawingResourceSheet = {
  readonly data: Record<string, DrawingImageResource>
  readonly order: readonly string[]
}

/** One `conditionalFormatting` collection entry as ExcelJS reads it. */
interface ExcelConditionalFormatting {
  readonly ref: string
  readonly rules: readonly ExcelConditionalRule[]
}

/**
 * One ExcelJS conditional-formatting rule. ExcelJS reads rule types its
 * declaration union omits (text-prefix and unique-value rules), so the reader
 * uses their runtime fields directly.
 */
interface ExcelConditionalRule {
  readonly type: string
  readonly priority?: number
  readonly formulae?: readonly unknown[]
  readonly operator?: string
  readonly text?: string
  readonly style?: Partial<ExcelStyle>
  readonly rank?: number
  readonly percent?: boolean
  readonly bottom?: boolean
  readonly aboveAverage?: boolean
  readonly timePeriod?: string
  readonly cfvo?: readonly ExcelConditionalValue[]
  readonly color?: Partial<ExcelColor> | readonly Partial<ExcelColor>[]
  readonly iconSet?: string
  readonly showValue?: boolean
  readonly gradient?: boolean
}

/** One ExcelJS conditional-formatting threshold. */
interface ExcelConditionalValue {
  readonly type: string
  readonly value?: number | string
}

/** One conditional-formatting rule ExcelJS's writer consumes. */
interface ExcelConditionalRuleOutput {
  readonly type: string
  readonly priority?: number
  readonly formulae?: readonly unknown[]
  readonly operator?: string
  readonly text?: string
  readonly style?: Partial<ExcelStyle>
  readonly rank?: number
  readonly percent?: boolean
  readonly bottom?: boolean
  readonly timePeriod?: string
  readonly cfvo?: readonly ExcelConditionalValue[]
  readonly color?: Partial<ExcelColor> | readonly Partial<ExcelColor>[]
  readonly iconSet?: string
  readonly showValue?: boolean
}

/** The worksheet fields ExcelJS's declaration omits for conditional formatting. */
interface ExcelConditionalFormattingCollection {
  readonly conditionalFormattings: readonly ExcelConditionalFormatting[]
  addConditionalFormatting(options: { ref: string; rules: readonly ExcelConditionalRuleOutput[] }): void
}

/** The worksheet field that holds the runtime data-validation collection. */
interface ExcelDataValidationsHost {
  readonly dataValidations: ExcelDataValidations
}

/** Frozen-pane state of a sheet whose view pins no rows or columns. */
const NO_FREEZE: IFreeze = { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 }

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

const UNIVER_HORIZONTAL_BY_EXCEL: Readonly<Record<string, HorizontalAlign>> = {
  left: HORIZONTAL_LEFT,
  center: HORIZONTAL_CENTER,
  right: HORIZONTAL_RIGHT,
  justify: HORIZONTAL_JUSTIFIED,
  centerContinuous: HORIZONTAL_CENTER,
  distributed: HORIZONTAL_DISTRIBUTED,
  fill: HORIZONTAL_DISTRIBUTED,
}

const EXCEL_HORIZONTAL_BY_UNIVER: Readonly<Record<number, NonNullable<ExcelAlignment['horizontal']>>> = {
  1: 'left',
  2: 'center',
  3: 'right',
  4: 'justify',
  5: 'justify',
  6: 'distributed',
}

const UNIVER_VERTICAL_BY_EXCEL: Readonly<Record<string, VerticalAlign>> = {
  top: VERTICAL_TOP,
  middle: VERTICAL_MIDDLE,
  bottom: VERTICAL_BOTTOM,
  justify: VERTICAL_MIDDLE,
  distributed: VERTICAL_MIDDLE,
}

const EXCEL_VERTICAL_BY_UNIVER: Readonly<Record<number, NonNullable<ExcelAlignment['vertical']>>> = {
  1: 'top',
  2: 'middle',
  3: 'bottom',
}

const EXCEL_VERT_ALIGN_BY_BASELINE: Readonly<Record<number, NonNullable<ExcelFont['vertAlign']>>> = {
  2: 'subscript',
  3: 'superscript',
}

const BASELINE_BY_EXCEL: Readonly<Record<string, BaselineOffset>> = {
  subscript: BASELINE_SUBSCRIPT,
  superscript: BASELINE_SUPERSCRIPT,
}

/**
 * Style fields shared by an ExcelJS cell, row, and column. ExcelJS exposes the
 * same setters on all three, so one mapping serves both directions.
 */
interface ExcelStyleFields {
  font?: Partial<ExcelFont>
  numFmt?: string
  fill?: ExcelFill
  border?: Partial<Borders>
  alignment?: Partial<ExcelAlignment>
}

/** The view facts that determine frozen panes, read structurally from ExcelJS. */
interface ExcelViewState {
  readonly state?: string
  readonly xSplit?: number
  readonly ySplit?: number
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
  const conditionalFormatting: Record<string, ConditionalFormattingResourceRule[]> = {}
  const dataValidation: Record<string, DataValidationResourceRule[]> = {}
  const drawings: Record<string, DrawingResourceSheet> = {}
  for (const [order, worksheet] of workbook.worksheets.entries()) {
    const id = `sheet-${order + 1}`
    sheetOrder.push(id)
    const sheet = readWorksheet(worksheet, id, date1904, () => {
      cellCount += 1
      if (cellCount > bounds.maxCells) {
        throw new OfficeExchangeError('too-large', `Workbook has more than ${bounds.maxCells} cells`)
      }
    })
    const background = readBackgroundImage(workbook, worksheet)
    if (background !== undefined) sheet.backgroundImage = background
    sheets[id] = sheet
    const rules = readConditionalFormatting(worksheet, order)
    if (rules.length > 0) conditionalFormatting[id] = rules
    const validations = readDataValidations(worksheet, order, date1904)
    if (validations.length > 0) dataValidation[id] = validations
    const sheetDrawings = readDrawings(workbook, worksheet, order, id)
    if (sheetDrawings !== undefined) drawings[id] = sheetDrawings
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
    resources: [
      /* v8 ignore next -- a workbook with no conditional formatting has no resource to append. */
      ...(Object.keys(conditionalFormatting).length === 0
        ? []
        : [{ name: CONDITIONAL_FORMATTING_RESOURCE, data: JSON.stringify(conditionalFormatting) }]),
      ...(Object.keys(dataValidation).length === 0
        ? []
        : [{ name: DATA_VALIDATION_RESOURCE, data: JSON.stringify(dataValidation) }]),
      ...(Object.keys(drawings).length === 0
        ? []
        : [{ name: DRAWING_RESOURCE, data: JSON.stringify(drawings) }]),
    ],
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
  const conditionalFormatting = resourceOf<Record<string, ConditionalFormattingResourceRule[]>>(
    data,
    CONDITIONAL_FORMATTING_RESOURCE,
  ) ?? {}
  const dataValidation = resourceOf<Record<string, DataValidationResourceRule[]>>(data, DATA_VALIDATION_RESOURCE) ?? {}
  const drawings = resourceOf<Record<string, DrawingResourceSheet>>(data, DRAWING_RESOURCE) ?? {}
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
    writeLayout(worksheet, sheet, data)
    for (const merge of sheet.mergeData ?? []) worksheet.mergeCells(mergeAddress(merge))
    writeConditionalFormatting(worksheet, conditionalFormatting[id] ?? [])
    writeDataValidations(worksheet, dataValidation[id] ?? [], data.dateSystem === DATE_1904)
    writeDrawings(workbook, worksheet, drawings[id])
    writeBackgroundImage(workbook, worksheet, sheet.backgroundImage)
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
    const rowStyle = readStyle(row)
    if (height !== undefined || row.hidden || rowStyle !== undefined) {
      rowData[rowIndex] = {
        ...(height === undefined ? {} : { h: pointsToPixels(height) }),
        ...(row.hidden ? { hd: BOOLEAN_TRUE } : {}),
        ...(rowStyle === undefined ? {} : { s: rowStyle }),
      }
    }
    row.eachCell({ includeEmpty: false }, (cell) => {
      countCell()
      const columnIndex = Number(cell.col) - 1
      cellData[rowIndex] = cellData[rowIndex] ?? {}
      cellData[rowIndex][columnIndex] = readCell(cell, date1904)
    })
  })
  const columnSources = parsed(worksheet.columns) ?? []
  for (const [index, source] of columnSources.entries()) {
    const width = parsed(source.width)
    const columnStyle = readStyle(source)
    if (width !== undefined || source.hidden || columnStyle !== undefined) {
      columnData[index] = {
        /* v8 ignore next -- ExcelJS serializes a width for every listed column, so the absent arm has no reachable producer. */
        ...(width === undefined ? {} : { w: columnPixels(width) }),
        ...(source.hidden ? { hd: BOOLEAN_TRUE } : {}),
        ...(columnStyle === undefined ? {} : { s: columnStyle }),
      }
    }
  }
  const properties = worksheet.properties
  const views = parsed(worksheet.views)
  const view: Partial<ExcelJS.WorksheetViewCommon> | undefined = views?.[0]
  const freezeView: ExcelViewState | undefined = views?.[0]
  const headersHidden = parsed(view?.showRowColHeaders) === false
  const sheetState = parsed(worksheet.state)
  return {
    id,
    name: worksheet.name,
    tabColor: readColor(parsed(properties.tabColor))?.rgb ?? '',
    hidden: sheetState === 'veryHidden' ? VERY_HIDDEN : sheetState === 'hidden' ? BOOLEAN_TRUE : BOOLEAN_FALSE,
    freeze: readFreeze(freezeView),
    rowCount: worksheet.rowCount,
    columnCount: worksheet.columnCount,
    zoomRatio: parsed(view?.zoomScale) ?? DEFAULT_ZOOM_RATIO,
    scrollTop: 0,
    scrollLeft: 0,
    defaultColumnWidth: columnPixels(parsed(properties.defaultColWidth) ?? 8.43),
    defaultRowHeight: pointsToPixels(properties.defaultRowHeight),
    mergeData: worksheet.model.merges.map(range => readMerge(worksheet, range)),
    cellData,
    rowData,
    columnData,
    rowHeader: headersHidden ? { width: 46, hidden: BOOLEAN_TRUE } : { width: 46 },
    columnHeader: headersHidden ? { height: 20, hidden: BOOLEAN_TRUE } : { height: 20 },
    showGridlines: parsed(view?.showGridLines) === false ? BOOLEAN_FALSE : BOOLEAN_TRUE,
    rightToLeft: parsed(view?.rightToLeft) === true ? BOOLEAN_TRUE : BOOLEAN_FALSE,
  }
}

/**
 * Read a worksheet's frozen-pane state from its first view. ExcelJS normalizes
 * `frozenSplit` to `frozen` while parsing, so both states name the same model.
 * @param view - First worksheet view, or `undefined` when the sheet has none.
 * @returns The frozen row and column counts, with the scrollable start at their boundary.
 */
function readFreeze(view: ExcelViewState | undefined): IFreeze {
  if (view === undefined || (view.state !== 'frozen' && view.state !== 'frozenSplit')) return NO_FREEZE
  /* v8 ignore next -- ExcelJS parses both pane splits as numbers; the fallback covers only the optional structural field. */
  const xSplit = parsed(view.xSplit) ?? 0
  /* v8 ignore next -- ExcelJS parses both pane splits as numbers; the fallback covers only the optional structural field. */
  const ySplit = parsed(view.ySplit) ?? 0
  return { xSplit, ySplit, startRow: ySplit, startColumn: xSplit }
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
  const hyperlink = readHyperlink(cell)
  if (hyperlink !== undefined) data.p = hyperlink
  return data
}

/**
 * Read a cell hyperlink into the rich-text cell document Univer stores for it.
 * A hyperlink lives in the cell's document body as a custom range carrying the
 * external URL, so the link and its display text survive a snapshot round-trip.
 * @param cell - ExcelJS cell.
 * @returns The cell document, or `undefined` when the cell carries no hyperlink.
 */
function readHyperlink(cell: ExcelCell): ICellData['p'] {
  const value = cell.value
  if (value === null || typeof value !== 'object' || !('hyperlink' in value)) return undefined
  const hyperlink = value.hyperlink
  if (typeof hyperlink !== 'string' || hyperlink.length === 0) return undefined
  const text = 'text' in value && typeof value.text === 'string' ? value.text : hyperlink
  const tooltip = 'tooltip' in value && typeof value.tooltip === 'string' ? value.tooltip : undefined
  return {
    id: `cell-link-${String(cell.row)}-${String(cell.col)}`,
    documentStyle: {},
    body: {
      dataStream: `${text}\r\n`,
      paragraphs: [{ startIndex: text.length, paragraphId: 'cell-paragraph' }],
      customRanges: [{
        startIndex: 0,
        endIndex: Math.max(0, text.length - 1),
        rangeId: `hyperlink-${String(cell.row)}-${String(cell.col)}`,
        rangeType: HYPERLINK_RANGE_TYPE,
        properties: tooltip === undefined ? { url: hyperlink } : { url: hyperlink, tooltip },
      }],
    },
  }
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

/**
 * Read a worksheet's conditional formats into the plugin resource rules.
 * ExcelJS reads one entry per `conditionalFormatting` element and keeps each
 * rule's priority, so the resource orders rules by that priority.
 * @param worksheet - ExcelJS worksheet.
 * @param sheetIndex - Zero-based position used to mint deterministic ids.
 * @returns The rules to persist, in priority order.
 */
function readConditionalFormatting(
  worksheet: ExcelJS.Worksheet,
  sheetIndex: number,
): ConditionalFormattingResourceRule[] {
  const collected: Array<{ readonly priority: number; readonly rule: ConditionalFormattingResourceRule }> = []
  for (const [groupIndex, group] of conditionalFormattingsOf(worksheet).conditionalFormattings.entries()) {
    const ranges = readRanges(group.ref)
    if (ranges.length === 0) continue
    for (const [ruleIndex, rule] of group.rules.entries()) {
      const config = conditionalRuleOf(rule, group.ref)
      if (config === undefined) continue
      collected.push({
        priority: typeof rule.priority === 'number' ? rule.priority : 0,
        rule: {
          cfId: `cf-${String(sheetIndex + 1)}-${String(groupIndex + 1)}-${String(ruleIndex + 1)}`,
          ranges,
          stopIfTrue: false,
          rule: config,
        },
      })
    }
  }
  collected.sort((left, right) => left.priority - right.priority)
  return collected.map(entry => entry.rule)
}

/**
 * Map one ExcelJS conditional-formatting rule onto the Univer rule model.
 * Rule types the sheet plugin cannot express are omitted; the README names
 * those narrow limitations.
 * @param rule - ExcelJS rule.
 * @param ref - The containing `conditionalFormatting` reference.
 * @returns The Univer rule, or `undefined` for a rule with no editable model.
 */
function conditionalRuleOf(
  rule: ExcelConditionalRule,
  ref: string,
): ConditionalFormattingResourceConfig | undefined {
  const style = readStyle(rule.style ?? {}) ?? {}
  const formula = formulaTextOf(rule.formulae?.[0])
  switch (rule.type) {
    case 'expression':
      return { type: CF_HIGHLIGHT, subType: CF_SUB_FORMULA, value: withFormulaPrefix(formula), style }
    case 'cellIs':
      return numberRuleOf(rule, ref, style)
    case 'containsText': {
      const operator = typeof rule.operator === 'string' ? rule.operator : 'containsText'
      const text = textOfRule(rule, formula)
      if (text === undefined) {
        return formula === undefined
          ? { type: CF_HIGHLIGHT, subType: CF_SUB_TEXT, operator, value: '', style }
          : { type: CF_HIGHLIGHT, subType: CF_SUB_FORMULA, value: withFormulaPrefix(formula), style }
      }
      return { type: CF_HIGHLIGHT, subType: CF_SUB_TEXT, operator, value: text, style }
    }
    case 'beginsWith':
    case 'endsWith':
    case 'notContainsText': {
      const text = textOfRule(rule, formula)
      if (text !== undefined) {
        return { type: CF_HIGHLIGHT, subType: CF_SUB_TEXT, operator: rule.type, value: text, style }
      }
      return formula === undefined
        ? undefined
        : { type: CF_HIGHLIGHT, subType: CF_SUB_FORMULA, value: withFormulaPrefix(formula), style }
    }
    case 'uniqueValues':
      return { type: CF_HIGHLIGHT, subType: CF_SUB_UNIQUE, style }
    case 'duplicateValues':
      return { type: CF_HIGHLIGHT, subType: CF_SUB_DUPLICATE, style }
    case 'top10':
      return {
        type: CF_HIGHLIGHT,
        subType: CF_SUB_RANK,
        isBottom: rule.bottom === true,
        isPercent: rule.percent === true,
        value: typeof rule.rank === 'number' ? rule.rank : 10,
        style,
      }
    case 'aboveAverage':
      return {
        type: CF_HIGHLIGHT,
        subType: CF_SUB_AVERAGE,
        operator: rule.aboveAverage === false ? 'lessThan' : 'greaterThan',
        style,
      }
    case 'timePeriod': {
      const operator = typeof rule.timePeriod === 'string'
        ? UNIVER_TIME_PERIOD_BY_EXCEL[rule.timePeriod]
        : undefined
      if (operator === undefined) {
        return formula === undefined
          ? undefined
          : { type: CF_HIGHLIGHT, subType: CF_SUB_FORMULA, value: withFormulaPrefix(formula), style }
      }
      return { type: CF_HIGHLIGHT, subType: CF_SUB_TIME_PERIOD, operator, style }
    }
    case 'colorScale': {
      const colors = colorArrayOf(rule.color)
      const config = (rule.cfvo ?? []).map((threshold, index) => ({
        index,
        color: colorStringOf(colors[index] ?? colors[0]),
        value: conditionalValueOf(threshold),
      }))
      return { type: CF_COLOR_SCALE, config }
    }
    case 'dataBar': {
      const thresholds = rule.cfvo ?? []
      const color = colorStringOf(colorArrayOf(rule.color)[0])
      return {
        type: CF_DATA_BAR,
        isShowValue: rule.showValue !== false,
        config: {
          min: conditionalValueOf(thresholds[0]),
          max: conditionalValueOf(thresholds[1]),
          isGradient: rule.gradient !== false,
          positiveColor: color,
          nativeColor: color,
        },
      }
    }
    case 'iconSet': {
      const iconType = iconTypeOf(rule.iconSet)
      const config = (rule.cfvo ?? []).map((threshold, index) => ({
        operator: 'greaterThanOrEqual',
        value: conditionalValueOf(threshold),
        iconType,
        iconId: String(index),
      }))
      return { type: CF_ICON_SET, isShowValue: rule.showValue !== false, config }
    }
    default:
      return undefined
  }
}

/**
 * Map a `cellIs` rule onto a numeric highlight rule, or a formula rule when the
 * comparison operand is not a literal number.
 * @param rule - ExcelJS cell-is rule.
 * @param ref - The containing range reference.
 * @param style - Rule style already mapped to the Univer model.
 * @returns The numeric rule.
 */
function numberRuleOf(
  rule: ExcelConditionalRule,
  ref: string,
  style: IStyleData,
): ConditionalFormattingResourceConfig {
  const operator = typeof rule.operator === 'string' ? UNIVER_CF_OPERATOR_BY_EXCEL[rule.operator] : undefined
  const resolved = operator ?? 'equal'
  const first = numberOf(rule.formulae?.[0])
  const second = numberOf(rule.formulae?.[1])
  if (resolved === 'between' || resolved === 'notBetween') {
    if (first !== undefined && second !== undefined) {
      return { type: CF_HIGHLIGHT, subType: CF_SUB_NUMBER, operator: resolved, value: [first, second], style }
    }
  } else if (first !== undefined) {
    return { type: CF_HIGHLIGHT, subType: CF_SUB_NUMBER, operator: resolved, value: first, style }
  }
  return {
    type: CF_HIGHLIGHT,
    subType: CF_SUB_FORMULA,
    value: withFormulaPrefix(expressionOf(rule, ref)),
    style,
  }
}

/**
 * Build a relative expression for a rule whose operand is not a literal number.
 * @param rule - ExcelJS cell-is rule.
 * @param ref - The containing range reference.
 * @returns The formula text that matches the rule against the range's first cell.
 */
function expressionOf(rule: ExcelConditionalRule, ref: string): string {
  const target = relativeReferenceOf(ref)
  const operator = typeof rule.operator === 'string' ? EXCEL_OPERATOR_TOKEN[rule.operator] : undefined
  const first = formulaTextOf(rule.formulae?.[0]) ?? '0'
  const second = formulaTextOf(rule.formulae?.[1]) ?? '0'
  if (operator === undefined) return `${target}=${first}`
  if (rule.operator === 'between') return `AND(${target}>=${first},${target}<=${second})`
  if (rule.operator === 'notBetween') return `OR(${target}<${first},${target}>${second})`
  return `${target}${operator}${first}`
}

/**
 * Read a hyperlink's display text from the rule's text attribute or formula.
 * @param rule - ExcelJS text rule.
 * @param formula - First formula of the rule.
 * @returns The text the rule matches, or `undefined` when it is not recoverable.
 */
function textOfRule(rule: ExcelConditionalRule, formula: string | undefined): string | undefined {
  if ('text' in rule && typeof rule.text === 'string') return rule.text
  if (formula === undefined) return undefined
  const quoted = /"([^"]*)"/u.exec(formula)
  return quoted?.[1]
}

/** @returns The formula text with a leading `=` when the value is non-empty. */
function withFormulaPrefix(formula: string | undefined): string {
  if (formula === undefined || formula.length === 0) return ''
  return formula.startsWith('=') ? formula : `=${formula}`
}

/** @returns The formula text without a leading `=`. */
function withoutFormulaPrefix(formula: string): string {
  return formula.startsWith('=') ? formula.slice(1) : formula
}

/** @returns A rule value as text, or `undefined` when absent. */
function formulaTextOf(value: unknown): string | undefined {
  if (typeof value === 'string') return value
  if (typeof value === 'number' && Number.isFinite(value)) return String(value)
  return undefined
}

/** @returns A rule value as a finite number, or `undefined` when it is not a number. */
function numberOf(value: unknown): number | undefined {
  if (typeof value === 'number') return Number.isFinite(value) ? value : undefined
  if (typeof value !== 'string' || value.trim().length === 0) return undefined
  const numeric = Number(value)
  return Number.isFinite(numeric) ? numeric : undefined
}

/** @returns The relative A1 reference of a range's first cell. */
function relativeReferenceOf(ref: string): string {
  const first = ref.trim().split(/\s+/u)[0] ?? ''
  const [start] = first.split(':')
  return (start ?? '').replaceAll('$', '') || 'A1'
}

/**
 * Map an ExcelJS threshold to the Univer conditional-formatting value model.
 * @param threshold - ExcelJS `cfvo` entry, when present.
 * @returns The Univer value, defaulting to an automatic minimum.
 */
function conditionalValueOf(threshold: ExcelConditionalValue | undefined): ConditionalFormattingValue {
  if (threshold === undefined) return { type: 'min' }
  switch (threshold.type) {
    case 'min':
    case 'autoMin':
      return { type: 'min' }
    case 'max':
    case 'autoMax':
      return { type: 'max' }
    case 'num':
    case 'percent':
    case 'percentile':
      return { type: threshold.type, value: typeof threshold.value === 'number' ? threshold.value : 0 }
    case 'formula':
      return { type: 'formula', value: threshold.value ?? 0 }
    default:
      return { type: 'min' }
  }
}

/** @returns A `#RRGGBB` string for an ExcelJS rule color. */
function colorStringOf(color: Partial<ExcelColor> | undefined): string {
  return readColor(color)?.rgb ?? '#000000'
}

/** @returns A rule's color or colors as one array. */
function colorArrayOf(
  color: Partial<ExcelColor> | readonly Partial<ExcelColor>[] | undefined,
): readonly Partial<ExcelColor>[] {
  if (color === undefined) return []
  if (Array.isArray(color)) return color
  return [color as Partial<ExcelColor>]
}

/** @returns The Univer icon-set name for an ExcelJS icon set. */
function iconTypeOf(iconSet: string | undefined): string {
  if (iconSet === undefined || !EXCEL_ICON_SET_NAMES.has(iconSet)) return 'EMPTY_ICON_TYPE'
  return iconSet
}

/**
 * Read a worksheet's data validations. ExcelJS expands one OOXML
 * `dataValidation` into its covered cell addresses while sharing the parsed
 * rule object, so grouping by object identity restores the original ranges.
 * @param worksheet - ExcelJS worksheet.
 * @param sheetIndex - Zero-based position used to mint deterministic ids.
 * @param date1904 - Whether date criteria use the 1904 serial system.
 * @returns The validation rules to persist.
 */
function readDataValidations(
  worksheet: ExcelJS.Worksheet,
  sheetIndex: number,
  date1904: boolean,
): DataValidationResourceRule[] {
  const model = validationsOf(worksheet).dataValidations.model
  const groups = new Map<ExcelDataValidation, string[]>()
  for (const [address, validation] of Object.entries(model)) {
    if (validation === undefined) continue
    const addresses = groups.get(validation) ?? []
    addresses.push(address)
    groups.set(validation, addresses)
  }
  const rules: DataValidationResourceRule[] = []
  for (const [validation, addresses] of groups) {
    const ranges = addresses.flatMap(address => readRanges(address))
    if (ranges.length === 0) continue
    rules.push({
      uid: `dv-${String(sheetIndex + 1)}-${String(rules.length + 1)}`,
      ...dataValidationFieldsOf(validation, date1904),
      ranges,
    })
  }
  return rules
}

/** The fields a Univer data-validation rule carries besides identity and ranges. */
type DataValidationFields = Omit<DataValidationResourceRule, 'uid' | 'ranges'>

/**
 * Map one ExcelJS data validation onto the Univer rule fields.
 * @param validation - ExcelJS validation.
 * @param date1904 - Whether date criteria use the 1904 serial system.
 * @returns The rule type, criteria, and options.
 */
function dataValidationFieldsOf(validation: ExcelDataValidation, date1904: boolean): DataValidationFields {
  const type = validation.type
  const formulae = validation.formulae ?? []
  const first = formulae[0]
  return {
    type,
    ...(validation.operator === undefined ? {} : { operator: validation.operator }),
    ...(first === undefined ? {} : { formula1: dataValidationFormula(type, first, date1904) }),
    ...(formulae[1] === undefined ? {} : { formula2: dataValidationFormula(type, formulae[1], date1904) }),
    ...(validation.allowBlank === true ? { allowBlank: true } : {}),
    ...(validation.showErrorMessage === true ? { showErrorMessage: true } : {}),
    ...(validation.errorStyle === undefined
      ? {}
      : { errorStyle: UNIVER_ERROR_STYLE_BY_EXCEL[validation.errorStyle] }),
    ...(validation.error === undefined ? {} : { error: validation.error }),
    ...(validation.errorTitle === undefined ? {} : { errorTitle: validation.errorTitle }),
    ...(validation.showInputMessage === true ? { showInputMessage: true } : {}),
    ...(validation.prompt === undefined ? {} : { prompt: validation.prompt }),
    ...(validation.promptTitle === undefined ? {} : { promptTitle: validation.promptTitle }),
  }
}

const UNIVER_ERROR_STYLE_BY_EXCEL: Readonly<Record<string, number>> = {
  information: 0,
  stop: 1,
  warning: 2,
}

const EXCEL_ERROR_STYLE_BY_UNIVER: Readonly<Record<number, string>> = {
  0: 'information',
  1: 'stop',
  2: 'warning',
}

/**
 * Convert one ExcelJS data-validation formula into the Univer criterion text.
 * @param type - Validation type.
 * @param formula - ExcelJS formula value (number, date, or string).
 * @param date1904 - Whether date criteria use the 1904 serial system.
 * @returns A Univer formula reference, a serialized literal list, or a number.
 */
function dataValidationFormula(type: string, formula: unknown, date1904: boolean): string {
  if (type === 'list') {
    if (typeof formula === 'string' && formula.startsWith('"') && formula.endsWith('"')) {
      const options = formula.slice(1, -1).split(',').map(option => option.trim()).filter(Boolean)
      return JSON.stringify(options)
    }
    return withFormulaPrefix(formulaTextOf(formula))
  }
  if (formula instanceof Date) return String(dateSerial(formula, date1904))
  if (typeof formula === 'number') return String(formula)
  const text = formulaTextOf(formula) ?? ''
  if (type === 'custom' || text.startsWith('$') || /^[A-Za-z]+\d+/u.test(text)) {
    return withFormulaPrefix(text)
  }
  return text
}

/**
 * Read worksheet drawings into the sheet-drawing plugin resource.
 * @param workbook - ExcelJS workbook that owns the media.
 * @param worksheet - ExcelJS worksheet.
 * @param sheetIndex - Zero-based position used to mint deterministic ids.
 * @param sheetId - The Univer sheet id the drawings belong to.
 * @returns The per-sheet drawing map, or `undefined` when the sheet has none.
 */
function readDrawings(
  workbook: ExcelJS.Workbook,
  worksheet: ExcelJS.Worksheet,
  sheetIndex: number,
  sheetId: string,
): DrawingResourceSheet | undefined {
  const images = worksheet.getImages()
  if (images.length === 0) return undefined
  const data: Record<string, DrawingImageResource> = {}
  const order: string[] = []
  for (const [index, image] of images.entries()) {
    const medium = workbook.getImage(Number(image.imageId))
    const drawingId = `sheet-drawing-${String(sheetIndex + 1)}-${String(index + 1)}`
    const range: ExcelImageRange = image.range
    const from = drawingPositionOf(anchorModelOf(range.tl))
    const to = range.br === undefined
      ? advanceAnchor(from, range.ext?.width ?? DEFAULT_COLUMN_PIXELS, range.ext?.height ?? DEFAULT_ROW_PIXELS, worksheet)
      : drawingPositionOf(anchorModelOf(range.br))
    const extension = medium.extension
    const hyperlink = range.hyperlinks === undefined
      ? {}
      : {
        hyperlink: {
          hyperlink: range.hyperlinks.hyperlink,
          ...(range.hyperlinks.tooltip === undefined ? {} : { tooltip: range.hyperlinks.tooltip }),
        },
      }
    data[drawingId] = {
      unitId: 'office-exchange-workbook',
      subUnitId: sheetId,
      drawingId,
      drawingType: DRAWING_TYPE_IMAGE,
      imageSourceType: IMAGE_SOURCE_BASE64,
      source: dataUrlOf(medium),
      sheetTransform: { from, to },
      axisAlignSheetTransform: { from, to },
      xlsx: {
        extension,
        ...(range.editAs === undefined ? {} : { editAs: range.editAs }),
        ...(range.ext === undefined ? {} : { ext: { width: range.ext.width, height: range.ext.height } }),
        ...hyperlink,
      },
    }
    order.push(drawingId)
  }
  return { data, order }
}

/** A drawing range as ExcelJS exposes it, including one-cell extensions. */
interface ExcelImageRange {
  readonly tl: ExcelJS.Anchor
  readonly br?: ExcelJS.Anchor
  readonly ext?: { readonly width: number; readonly height: number }
  readonly editAs?: string
  readonly hyperlinks?: { readonly hyperlink: string; readonly tooltip?: string }
}

/** The cell and offset fields an ExcelJS image anchor stores. */
interface ExcelAnchorModel {
  readonly nativeCol: number
  readonly nativeColOff: number
  readonly nativeRow: number
  readonly nativeRowOff: number
}

/**
 * Read the cell and offset fields an ExcelJS anchor stores. The declaration
 * omits the `model` getter every anchor instance exposes.
 * @param anchor - ExcelJS anchor.
 * @returns The anchor's cell and offset fields.
 */
function anchorModelOf(anchor: ExcelJS.Anchor): ExcelAnchorModel {
  return (anchor as ExcelJS.Anchor & { readonly model: ExcelAnchorModel }).model
}

/**
 * Project an ExcelJS anchor model onto the Univer drawing position. ExcelJS
 * stores offsets in EMU while the sheet model uses layout pixels.
 * @param anchor - ExcelJS anchor model.
 * @returns The Univer drawing position.
 */
function drawingPositionOf(anchor: ExcelAnchorModel): DrawingCellPosition {
  return {
    column: anchor.nativeCol,
    columnOffset: anchor.nativeColOff / EMU_PER_PIXEL,
    row: anchor.nativeRow,
    rowOffset: anchor.nativeRowOff / EMU_PER_PIXEL,
  }
}

/**
 * Advance a drawing position by a pixel extent across rows and columns.
 * @param from - Starting position.
 * @param width - Extent width in pixels.
 * @param height - Extent height in pixels.
 * @param worksheet - Worksheet that supplies row heights and column widths.
 * @returns The ending position.
 */
function advanceAnchor(
  from: DrawingCellPosition,
  width: number,
  height: number,
  worksheet: ExcelJS.Worksheet,
): DrawingCellPosition {
  let column = from.column
  let columnOffset = from.columnOffset
  let remainingWidth = width
  while (remainingWidth > 0) {
    const available = Math.max(1, columnPixelsOf(worksheet, column) - columnOffset)
    if (remainingWidth < available) {
      columnOffset += remainingWidth
      remainingWidth = 0
    } else {
      remainingWidth -= available
      column += 1
      columnOffset = 0
    }
  }
  let row = from.row
  let rowOffset = from.rowOffset
  let remainingHeight = height
  while (remainingHeight > 0) {
    const available = Math.max(1, rowPixelsOf(worksheet, row) - rowOffset)
    if (remainingHeight < available) {
      rowOffset += remainingHeight
      remainingHeight = 0
    } else {
      remainingHeight -= available
      row += 1
      rowOffset = 0
    }
  }
  return { column, columnOffset, row, rowOffset }
}

/** @returns A column's width in layout pixels. */
function columnPixelsOf(worksheet: ExcelJS.Worksheet, column: number): number {
  const width = parsed(worksheet.getColumn(column + 1).width)
  return columnPixels(width ?? parsed(worksheet.properties.defaultColWidth) ?? 8.43)
}

/** @returns A row's height in layout pixels. */
function rowPixelsOf(worksheet: ExcelJS.Worksheet, row: number): number {
  const height = parsed(worksheet.getRow(row + 1).height)
  return height === undefined
    ? pointsToPixels(parsed(worksheet.properties.defaultRowHeight) ?? 15)
    : pointsToPixels(height)
}

/** @returns A base64 data URL for one ExcelJS media part. */
function dataUrlOf(medium: ExcelImage): string {
  if (medium.buffer !== undefined) {
    const bytes = new Uint8Array(medium.buffer)
    return `data:${mimeOfExtension(medium.extension)};base64,${Buffer.from(bytes).toString('base64')}`
  }
  return medium.base64 ?? `data:${mimeOfExtension(medium.extension)};base64,`
}

/** @returns The image MIME type for a file extension. */
function mimeOfExtension(extension: string): string {
  return extension === 'jpg' ? 'image/jpeg' : `image/${extension.toLowerCase()}`
}

/**
 * Read a worksheet background image into the Univer worksheet field.
 * @param workbook - ExcelJS workbook that owns the media.
 * @param worksheet - ExcelJS worksheet.
 * @returns The background image, or `undefined` when the sheet has none.
 */
function readBackgroundImage(
  workbook: ExcelJS.Workbook,
  worksheet: ExcelJS.Worksheet,
): IWorksheetBackgroundImage | undefined {
  const imageId = parsed(worksheet.getBackgroundImageId())
  if (imageId === undefined || imageId === '') return undefined
  const medium = workbook.getImage(Number(imageId))
  return {
    source: dataUrlOf(medium),
    imageSourceType: IMAGE_SOURCE_BASE64 as IWorksheetBackgroundImage['imageSourceType'],
  }
}

/**
 * Parse an A1 reference into Univer ranges. The plugin's rule ranges use
 * inclusive end cells, matching Univer's range model.
 * @param ref - Space-separated A1 references from ExcelJS.
 * @returns The parsed ranges; tokens that are not A1 references are skipped.
 */
function readRanges(ref: string): IRange[] {
  const ranges: IRange[] = []
  for (const token of ref.trim().split(/\s+/u)) {
    const match = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/u.exec(token)
    if (match === null) continue
    const startColumn = columnIndexOf(match[1] ?? 'A')
    const startRow = Number(match[2]) - 1
    const endColumn = match[3] === undefined ? startColumn : columnIndexOf(match[3])
    const endRow = match[4] === undefined ? startRow : Number(match[4]) - 1
    ranges.push({ startRow, startColumn, endRow, endColumn })
  }
  return ranges
}

/** @returns The zero-based column index for an A1 column label. */
function columnIndexOf(label: string): number {
  let value = 0
  for (const character of label.toUpperCase()) {
    value = value * 26 + character.charCodeAt(0) - 64
  }
  return value - 1
}

/** ExcelJS's worksheet data-validation collection, which its typings omit. */
interface ExcelDataValidations {
  readonly model: Record<string, ExcelDataValidation | undefined>
  add(address: string, validation: ExcelDataValidation): ExcelDataValidation
}

/** An ExcelJS data validation with the `any` type its parser can produce. */
interface ExcelDataValidation {
  readonly type: string
  readonly formulae?: readonly unknown[]
  readonly allowBlank?: boolean
  readonly operator?: string
  readonly error?: string
  readonly errorTitle?: string
  readonly errorStyle?: string
  readonly prompt?: string
  readonly promptTitle?: string
  readonly showErrorMessage?: boolean
  readonly showInputMessage?: boolean
}

/** @returns The worksheet's runtime data-validation collection. */
function validationsOf(worksheet: ExcelJS.Worksheet): ExcelDataValidationsHost {
  return worksheet as ExcelJS.Worksheet & ExcelDataValidationsHost
}

/** @returns The worksheet's runtime conditional-formatting collection. */
function conditionalFormattingsOf(worksheet: ExcelJS.Worksheet): ExcelConditionalFormattingCollection {
  return worksheet as ExcelJS.Worksheet & ExcelConditionalFormattingCollection
}

/**
 * Write conditional-formatting rules into a worksheet.
 * @param worksheet - Target worksheet.
 * @param rules - Resource rules for this sheet.
 */
function writeConditionalFormatting(
  worksheet: ExcelJS.Worksheet,
  rules: readonly ConditionalFormattingResourceRule[],
): void {
  const collection = conditionalFormattingsOf(worksheet)
  let priority = 1
  for (const entry of rules) {
    const rule = excelConditionalRuleOf(entry.rule, entry.ranges)
    collection.addConditionalFormatting({
      ref: rangeReferenceOf(entry.ranges),
      rules: [{ ...rule, priority: priority++ }],
    })
  }
}

/**
 * Map one Univer conditional-formatting rule back to ExcelJS.
 * @param config - Univer rule body.
 * @param ranges - Rule ranges, used by expression and average fallbacks.
 * @returns The ExcelJS rule.
 * @throws {OfficeExchangeError} `conversion-failed` for a rule body with no ExcelJS writer.
 */
function excelConditionalRuleOf(
  config: ConditionalFormattingResourceConfig,
  ranges: readonly IRange[],
): ExcelConditionalRuleOutput {
  const style = excelStyleOf('style' in config ? config.style : {}) ?? {}
  switch (config.type) {
    case CF_HIGHLIGHT:
      return excelHighlightRuleOf(config, ranges, style)
    case CF_COLOR_SCALE:
      return {
        type: 'colorScale',
        cfvo: config.config.map(entry => excelThresholdOf(entry.value)),
        color: config.config.map(entry => ({ argb: argbOf({ rgb: entry.color }) ?? 'FF000000' })),
        style,
      }
    case CF_DATA_BAR:
      return {
        type: 'dataBar',
        cfvo: [excelThresholdOf(config.config.min), excelThresholdOf(config.config.max)],
        color: { argb: argbOf({ rgb: config.config.positiveColor }) ?? 'FF638EC6' },
        showValue: config.isShowValue,
        style,
      }
    case CF_ICON_SET:
      return {
        type: 'iconSet',
        iconSet: config.config[0]?.iconType === undefined || config.config[0].iconType === 'EMPTY_ICON_TYPE'
          ? 'NoIcons'
          : config.config[0].iconType,
        showValue: config.isShowValue,
        cfvo: config.config.map(entry => excelThresholdOf(entry.value)),
        style,
      }
  }
}

/**
 * Map one Univer highlight-cell rule back to ExcelJS. Unique, duplicate, and
 * average rules become equivalent expressions because ExcelJS has no writer
 * for those rule types.
 * @param config - Univer highlight rule body.
 * @param ranges - Rule ranges.
 * @param style - ExcelJS rule style.
 * @returns The ExcelJS rule.
 */
function excelHighlightRuleOf(
  config: Extract<ConditionalFormattingResourceConfig, { readonly type: typeof CF_HIGHLIGHT }>,
  ranges: readonly IRange[],
  style: Partial<ExcelStyle>,
): ExcelConditionalRuleOutput {
  switch (config.subType) {
    case CF_SUB_FORMULA:
      return { type: 'expression', formulae: [withoutFormulaPrefix(config.value)], style }
    case CF_SUB_NUMBER: {
      const formulae = Array.isArray(config.value) ? [...config.value] : [config.value]
      return { type: 'cellIs', operator: config.operator, formulae, style }
    }
    case CF_SUB_TEXT: {
      const operator = config.operator
      const target = relativeReferenceOf(rangeReferenceOf(ranges))
      if (operator === 'equal' || operator === 'notEqual') {
        return {
          type: 'expression',
          formulae: [`${target}${operator === 'equal' ? '=' : '<>'}${excelStringLiteral(config.value)}`],
          style,
        }
      }
      return {
        type: 'containsText',
        operator,
        text: config.value,
        formulae: [textRuleFormulaOf(operator, config.value, target)],
        style,
      }
    }
    case CF_SUB_RANK:
      return { type: 'top10', rank: config.value, percent: config.isPercent, bottom: config.isBottom, style }
    case CF_SUB_AVERAGE:
      return {
        type: 'expression',
        formulae: [averageRuleFormulaOf(config.operator, ranges)],
        style,
      }
    case CF_SUB_UNIQUE:
      return { type: 'expression', formulae: [`COUNTIF(${absoluteReferenceOf(ranges)},${relativeReferenceOf(rangeReferenceOf(ranges))})=1`], style }
    case CF_SUB_DUPLICATE:
      return { type: 'expression', formulae: [`COUNTIF(${absoluteReferenceOf(ranges)},${relativeReferenceOf(rangeReferenceOf(ranges))})>1`], style }
    case CF_SUB_TIME_PERIOD: {
      const timePeriod = EXCEL_TIME_PERIOD_BY_UNIVER[config.operator]
      if (timePeriod === undefined) {
        throw new OfficeExchangeError('conversion-failed', `Cannot write time-period rule ${config.operator}`)
      }
      return { type: 'timePeriod', timePeriod, style }
    }
  }
}

/**
 * @param operator - Univer text operator.
 * @param value - Text the rule matches.
 * @param target - The range's first cell as a relative A1 reference.
 * @returns The ExcelJS text rule formula.
 */
function textRuleFormulaOf(operator: string, value: string, target: string): string {
  const literal = excelStringLiteral(value)
  switch (operator) {
    case 'beginsWith':
      return `LEFT(${target},LEN(${literal}))=${literal}`
    case 'endsWith':
      return `RIGHT(${target},LEN(${literal}))=${literal}`
    case 'notContainsText':
      return `ISERROR(SEARCH(${literal},${target}))`
    case 'containsBlanks':
      return `LEN(TRIM(${target}))=0`
    case 'notContainsBlanks':
      return `LEN(TRIM(${target}))>0`
    case 'containsErrors':
      return `ISERROR(${target})`
    case 'notContainsErrors':
      return `NOT(ISERROR(${target}))`
    default:
      return `NOT(ISERROR(SEARCH(${literal},${target})))`
  }
}

/** @returns A quoted Excel formula string literal. */
function excelStringLiteral(value: string): string {
  return `"${value.replaceAll('"', '""')}"`
}

/**
 * @param operator - Univer average operator.
 * @param ranges - Rule ranges.
 * @returns An expression that compares the range's first cell with its average.
 */
function averageRuleFormulaOf(operator: string, ranges: readonly IRange[]): string {
  const target = relativeReferenceOf(rangeReferenceOf(ranges))
  const average = `AVERAGE(${absoluteReferenceOf(ranges)})`
  const token = EXCEL_OPERATOR_TOKEN[operator] ?? '>'
  return `${target}${token}${average}`
}

/**
 * Map a Univer threshold back to an ExcelJS `cfvo`.
 * @param value - Univer threshold.
 * @returns The ExcelJS threshold.
 */
function excelThresholdOf(value: ConditionalFormattingValue): ExcelConditionalValue {
  switch (value.type) {
    case 'min':
      return { type: 'min' }
    case 'max':
      return { type: 'max' }
    case 'num':
    case 'percent':
    case 'percentile':
      return { type: value.type, value: typeof value.value === 'number' ? value.value : Number(value.value ?? 0) }
    default: {
      const numeric = Number(value.value ?? 0)
      return { type: 'formula', value: Number.isFinite(numeric) ? numeric : 0 }
    }
  }
}

/**
 * @param range - One Univer range with an inclusive end cell.
 * @returns A single-cell address or a rectangular A1 reference.
 */
function dataValidationAddressOf(range: IRange): string {
  if (range.startRow === range.endRow && range.startColumn === range.endColumn) {
    return cellAddress(range.startRow + 1, range.startColumn + 1)
  }
  return rangeReferenceOf([range])
}

/** @returns The space-separated A1 reference for one or more ranges. */
function rangeReferenceOf(ranges: readonly IRange[]): string {
  return ranges.map(range => `${cellAddress(range.startRow + 1, range.startColumn + 1)}:${cellAddress(range.endRow + 1, range.endColumn + 1)}`).join(' ')
}

/** @returns The absolute A1 reference spanning every range. */
function absoluteReferenceOf(ranges: readonly IRange[]): string {
  const first = ranges[0]
  if (first === undefined) return '$A$1'
  const start = `${absoluteColumn(first.startColumn)}$${String(first.startRow + 1)}`
  const end = `${absoluteColumn(first.endColumn)}$${String(first.endRow + 1)}`
  return `${start}:${end}`
}

/** @returns An absolute column label. */
function absoluteColumn(index: number): string {
  return `$${cellAddress(1, index + 1).replace(/\d+$/u, '')}`
}

/**
 * Write data-validation rules into a worksheet. Each range is stored under its
 * own address because ExcelJS's writer optimizes only a single range key.
 * @param worksheet - Target worksheet.
 * @param rules - Resource rules for this sheet.
 * @param date1904 - Whether date criteria use the 1904 serial system.
 */
function writeDataValidations(
  worksheet: ExcelJS.Worksheet,
  rules: readonly DataValidationResourceRule[],
  date1904: boolean,
): void {
  const validations = validationsOf(worksheet).dataValidations
  for (const rule of rules) {
    const validation = excelDataValidationOf(rule, date1904)
    for (const range of rule.ranges) {
      validations.add(dataValidationAddressOf(range), validation)
    }
  }
}

/**
 * Map one Univer data-validation rule back to ExcelJS.
 * @param rule - Resource rule.
 * @param date1904 - Whether date criteria use the 1904 serial system.
 * @returns The ExcelJS validation.
 * @throws {OfficeExchangeError} `conversion-failed` for a type ExcelJS cannot write.
 */
function excelDataValidationOf(rule: DataValidationResourceRule, date1904: boolean): ExcelDataValidation {
  if (rule.type === 'listMultiple' || rule.type === 'checkbox' || rule.type === 'none') {
    throw new OfficeExchangeError('conversion-failed', `Cannot write data validation type ${rule.type}`)
  }
  const type = rule.type
  const formula1 = rule.formula1 === undefined ? undefined : excelDataValidationFormula(type, rule.formula1, date1904)
  const formula2 = rule.formula2 === undefined ? undefined : excelDataValidationFormula(type, rule.formula2, date1904)
  return {
    type,
    formulae: [formula1, formula2].filter(value => value !== undefined),
    ...(rule.operator === undefined || type === 'list' || type === 'custom' || type === 'any'
      ? {}
      : { operator: rule.operator }),
    ...(rule.allowBlank === true ? { allowBlank: true } : {}),
    ...(rule.showErrorMessage === true ? { showErrorMessage: true } : {}),
    ...(rule.errorStyle === undefined ? {} : { errorStyle: EXCEL_ERROR_STYLE_BY_UNIVER[rule.errorStyle] }),
    ...(rule.error === undefined ? {} : { error: rule.error }),
    ...(rule.errorTitle === undefined ? {} : { errorTitle: rule.errorTitle }),
    ...(rule.showInputMessage === true ? { showInputMessage: true } : {}),
    ...(rule.prompt === undefined ? {} : { prompt: rule.prompt }),
    ...(rule.promptTitle === undefined ? {} : { promptTitle: rule.promptTitle }),
  }
}

/**
 * Convert one Univer data-validation criterion back to the ExcelJS value.
 * @param type - Validation type.
 * @param formula - Univer criterion text.
 * @param date1904 - Whether date criteria use the 1904 serial system.
 * @returns The ExcelJS formula value.
 */
function excelDataValidationFormula(type: string, formula: string, date1904: boolean): unknown {
  if (type === 'list') {
    if (formula.startsWith('[')) {
      const options = listOptionsOf(formula)
      return `"${options.join(',')}"`
    }
    return withoutFormulaPrefix(formula)
  }
  if (type === 'date') {
    const serial = Number(withoutFormulaPrefix(formula))
    return Number.isFinite(serial) && !formula.startsWith('=')
      ? serialToDate(serial, date1904)
      : withoutFormulaPrefix(formula)
  }
  if (type === 'whole' || type === 'textLength' || type === 'decimal') {
    const numeric = numberOf(withoutFormulaPrefix(formula))
    return numeric ?? withoutFormulaPrefix(formula)
  }
  return withoutFormulaPrefix(formula)
}

/**
 * @param formula - A serialized Univer list criterion.
 * @returns The literal list options.
 */
function listOptionsOf(formula: string): string[] {
  try {
    const value: unknown = JSON.parse(formula)
    if (Array.isArray(value)) return value.filter((entry): entry is string => typeof entry === 'string')
  } catch {
    // The strip-and-split fallback below reads a comma-separated criterion.
  }
  return formula.split(',').map(option => option.trim()).filter(Boolean)
}

/**
 * Write cell drawings into a worksheet.
 * @param workbook - Target workbook that owns the media.
 * @param worksheet - Target worksheet.
 * @param drawings - Drawing resource for this sheet, when present.
 * @throws {OfficeExchangeError} `conversion-failed` when a drawing is not an embedded image.
 */
function writeDrawings(
  workbook: ExcelJS.Workbook,
  worksheet: ExcelJS.Worksheet,
  drawings: DrawingResourceSheet | undefined,
): void {
  if (drawings === undefined) return
  for (const drawingId of drawings.order) {
    const drawing = drawings.data[drawingId]
    if (drawing === undefined) continue
    if (drawing.drawingType !== DRAWING_TYPE_IMAGE) {
      throw new OfficeExchangeError('conversion-failed', `Cannot write drawing ${drawingId}: not an image`)
    }
    const imageId = workbook.addImage({
      base64: drawing.source,
      extension: (drawing.xlsx?.extension ?? extensionOfMime(decodeDataUrl(drawing.source).mime)) as ExcelImage['extension'],
    })
    imagesOf(worksheet).addImage(imageId, excelImagePlacementOf(drawing))
  }
}

/** One decoded data URL. */
interface DecodedImage {
  readonly mime: string
  readonly base64: string
}

/**
 * @param source - A base64 data URL.
 * @returns The MIME type and payload.
 * @throws {OfficeExchangeError} `conversion-failed` when the source is not an embedded image.
 */
function decodeDataUrl(source: string): DecodedImage {
  const match = /^data:([^;,]+);base64,(.*)$/su.exec(source)
  if (match === null) {
    throw new OfficeExchangeError('conversion-failed', 'Cannot write an image that is not a base64 data URL')
  }
  return { mime: match[1] ?? 'image/png', base64: match[2] ?? '' }
}

/** @returns The file extension for an image MIME type. */
function extensionOfMime(mime: string): string {
  return mime === 'image/jpeg' ? 'jpeg' : mime.replace(/^image\//u, '')
}

/**
 * Build the ExcelJS placement for one persisted drawing.
 * @param drawing - Persisted image drawing.
 * @returns The anchor placement ExcelJS accepts.
 */
function excelImagePlacementOf(drawing: DrawingImageResource): ExcelImagePlacement {
  const tl = excelAnchorOf(drawing.sheetTransform.from)
  const br = excelAnchorOf(drawing.sheetTransform.to)
  const hyperlinks = drawing.xlsx?.hyperlink === undefined ? {} : { hyperlinks: drawing.xlsx.hyperlink }
  if (drawing.xlsx?.ext !== undefined) {
    return { tl, ext: drawing.xlsx.ext, editAs: drawing.xlsx.editAs ?? 'oneCell', ...hyperlinks }
  }
  return { tl, br, editAs: drawing.xlsx?.editAs ?? 'twoCell', ...hyperlinks }
}

/** The anchor fields ExcelJS stores for one image corner. */
interface ExcelImageAnchor {
  readonly col: number
  readonly row: number
  readonly nativeCol: number
  readonly nativeColOff: number
  readonly nativeRow: number
  readonly nativeRowOff: number
}

/** The image placement ExcelJS's worksheet accepts at runtime. */
interface ExcelImagePlacement {
  readonly tl: ExcelImageAnchor
  readonly br?: ExcelImageAnchor
  readonly ext?: { readonly width: number; readonly height: number }
  readonly editAs?: string
  readonly hyperlinks?: { readonly hyperlink: string; readonly tooltip?: string }
}

/** The worksheet image methods ExcelJS declares with its own class types. */
interface ExcelWorksheetImages {
  addImage(imageId: number, placement: ExcelImagePlacement): void
  addBackgroundImage(imageId: number): void
}

/** @returns The worksheet's runtime image methods. */
function imagesOf(worksheet: ExcelJS.Worksheet): ExcelWorksheetImages {
  return worksheet as ExcelJS.Worksheet & ExcelWorksheetImages
}

/**
 * @param position - Univer drawing position.
 * @returns The ExcelJS anchor fields.
 */
function excelAnchorOf(position: DrawingCellPosition): ExcelImageAnchor {
  return {
    col: position.column,
    row: position.row,
    nativeCol: position.column,
    nativeColOff: Math.round(position.columnOffset * EMU_PER_PIXEL),
    nativeRow: position.row,
    nativeRowOff: Math.round(position.rowOffset * EMU_PER_PIXEL),
  }
}

/**
 * Write a worksheet background image.
 * @param workbook - Target workbook that owns the media.
 * @param worksheet - Target worksheet.
 * @param image - Univer worksheet background image, when present.
 * @throws {OfficeExchangeError} `conversion-failed` when the image is not an embedded base64 source.
 */
function writeBackgroundImage(
  workbook: ExcelJS.Workbook,
  worksheet: ExcelJS.Worksheet,
  image: IWorksheetBackgroundImage | undefined,
): void {
  if (image === undefined) return
  const decoded = decodeDataUrl(image.source)
  const imageId = workbook.addImage({
    base64: image.source,
    extension: extensionOfMime(decoded.mime) as ExcelImage['extension'],
  })
  imagesOf(worksheet).addBackgroundImage(imageId)
}

/**
 * Convert a workbook snapshot's serial date back to a `Date`.
 * @param serial - Excel date serial.
 * @param date1904 - Whether the workbook uses the 1904 system.
 * @returns The corresponding date.
 */
function serialToDate(serial: number, date1904: boolean): Date {
  return new Date((serial - 25_569 + (date1904 ? 1_462 : 0)) * 86_400_000)
}

function readStyle(source: ExcelStyleFields): IStyleData | undefined {
  const style: IStyleData = {}
  const font = parsed(source.font)
  if (font !== undefined && font !== null) {
    if (font.bold === true) style.bl = BOOLEAN_TRUE
    if (font.italic === true) style.it = BOOLEAN_TRUE
    if (font.strike === true) style.st = { s: BOOLEAN_TRUE }
    const vertAlign = parsed(font.vertAlign)
    if (vertAlign !== undefined) style.va = BASELINE_BY_EXCEL[vertAlign]
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
  const alignment = parsed(source.alignment)
  if (alignment !== undefined && alignment !== null) {
    const horizontal = parsed(alignment.horizontal)
    if (horizontal !== undefined) style.ht = UNIVER_HORIZONTAL_BY_EXCEL[horizontal]
    const vertical = parsed(alignment.vertical)
    if (vertical !== undefined) style.vt = UNIVER_VERTICAL_BY_EXCEL[vertical]
    if (alignment.wrapText === true) style.tb = WRAP_STRATEGY_WRAP
    if (alignment.shrinkToFit === true) style.stf = BOOLEAN_TRUE
    const rotation = readRotation(parsed(alignment.textRotation))
    if (rotation !== undefined) style.tr = rotation
  }
  const numFmt = parsed(source.numFmt)
  if (numFmt !== undefined && numFmt !== 'General') style.n = { pattern: numFmt }
  const fill = parsed(source.fill)
  if (fill !== undefined && fill !== null && fill.type === 'pattern' && fill.pattern === 'solid') {
    const color = readColor(parsed(fill.fgColor))
    if (color !== undefined) style.bg = color
  }
  const border = readBorder(parsed(source.border) ?? undefined)
  if (border !== undefined) style.bd = border
  return Object.keys(style).length === 0 ? undefined : style
}

/**
 * Map an ExcelJS alignment rotation to the Univer model. ExcelJS normalizes
 * OOXML's 0-180 encoding to signed degrees and names vertical text.
 * @param rotation - ExcelJS alignment rotation.
 * @returns The Univer rotation, or `undefined` for the default horizontal text.
 */
function readRotation(
  rotation: number | typeof VERTICAL_ROTATION | undefined | null,
): ITextRotation | undefined {
  if (rotation === undefined || rotation === null) return undefined
  return rotation === VERTICAL_ROTATION ? { a: 0, v: BOOLEAN_TRUE } : { a: rotation }
}

function readBorder(border: Partial<Borders> | undefined): IBorderData | undefined {
  if (border === undefined) return undefined
  const data: IBorderData = {}
  for (const [side, key] of [['top', 't'], ['right', 'r'], ['bottom', 'b'], ['left', 'l']] as const) {
    const edge = border[side]
    if (edge === undefined || edge === null) continue
    data[key] = {
      s: BORDER_STYLE_BY_EXCEL[edge.style as BorderStyle],
      cl: readColor(parsed(edge.color)) ?? { rgb: '#000000' },
    }
  }
  return Object.keys(data).length === 0 ? undefined : data
}

function readColor(color: Partial<ExcelColor> | undefined | null): IColorStyle | undefined {
  if (color === undefined || color === null) return undefined
  const argb = parsed(color.argb)
  if (argb === undefined || argb.length < 6) return undefined
  const rgb = argb.slice(-6)
  return /^[0-9a-f]{6}$/iu.test(rgb) ? { rgb: `#${rgb}` } : undefined
}

/**
 * Read the hyperlink a cell document carries.
 * @param data - Univer cell data.
 * @returns The ExcelJS hyperlink value, or `undefined` when the cell has no link.
 */
function hyperlinkOf(data: ICellData): { text: string; hyperlink: string; tooltip?: string } | undefined {
  const body = data.p?.body
  if (body === undefined) return undefined
  const range = body.customRanges?.find(entry => entry.rangeType === HYPERLINK_RANGE_TYPE)
  const url = range?.properties?.['url']
  if (typeof url !== 'string' || url.length === 0) return undefined
  const text = cellTextOf(data)
  const tooltip = range?.properties?.['tooltip']
  return {
    text: text.length === 0 ? url : text,
    hyperlink: url,
    ...(typeof tooltip === 'string' && tooltip.length > 0 ? { tooltip } : {}),
  }
}

/** @returns A cell's rich-text content without the trailing paragraph mark. */
function cellTextOf(data: ICellData): string {
  const stream = data.p?.body?.dataStream
  if (stream === undefined) return ''
  return stream.endsWith('\r\n') ? stream.slice(0, -2) : stream
}

function writeCell(cell: ExcelCell, data: ICellData, workbook: IWorkbookData): void {
  const hyperlink = hyperlinkOf(data)
  if (data.f !== undefined && data.f !== null && data.f.length > 0) {
    const formula = data.f.startsWith('=') ? data.f.slice(1) : data.f
    cell.value = data.v === undefined || data.v === null
      ? { formula }
      : { formula, result: data.v }
  } else if (hyperlink !== undefined) {
    cell.value = hyperlink
  } else if (data.v !== undefined && data.v !== null) {
    cell.value = data.v
  } else {
    const text = cellTextOf(data)
    if (text.length > 0) cell.value = text
  }
  const style = resolveStyle(workbook, data.s)
  if (style === undefined) return
  const excelStyle = excelStyleOf(style)
  if (excelStyle === undefined) return
  if (excelStyle.font !== undefined) cell.font = excelStyle.font
  if (excelStyle.numFmt !== undefined) cell.numFmt = excelStyle.numFmt
  if (excelStyle.fill !== undefined) cell.fill = excelStyle.fill
  if (excelStyle.border !== undefined) cell.border = excelStyle.border
  if (excelStyle.alignment !== undefined) cell.alignment = excelStyle.alignment
}

function resolveStyle(workbook: IWorkbookData, style: ICellData['s']): IStyleData | undefined {
  if (style === undefined || style === null) return undefined
  return typeof style === 'string' ? workbook.styles[style] ?? undefined : style
}

/**
 * Translate a Univer cell style into the fields ExcelJS stores on a cell, row,
 * or column.
 * @param style - Univer style data.
 * @returns ExcelJS style fields, or `undefined` when the style is empty.
 */
function excelStyleOf(style: IStyleData): ExcelStyleFields | undefined {
  const data: ExcelStyleFields = {}
  const font: Partial<ExcelFont> = {}
  if (style.bl === BOOLEAN_TRUE) font.bold = true
  if (style.it === BOOLEAN_TRUE) font.italic = true
  if (style.st?.s === BOOLEAN_TRUE) font.strike = true
  if (style.va !== undefined && style.va !== null) {
    const vertAlign = EXCEL_VERT_ALIGN_BY_BASELINE[style.va]
    /* v8 ignore next -- Every BaselineOffset member has an ExcelJS name, so the lookup always resolves. */
    if (vertAlign !== undefined) font.vertAlign = vertAlign
  }
  if (style.ul?.s === BOOLEAN_TRUE) {
    font.underline = style.ul.t === TEXT_DECORATION_DOUBLE ? 'double' : 'single'
  }
  if (style.fs !== undefined) font.size = style.fs
  if (style.ff !== undefined && style.ff !== null) font.name = style.ff
  const fontColor = argbOf(style.cl)
  if (fontColor !== undefined) font.color = { argb: fontColor }
  if (Object.keys(font).length > 0) data.font = font
  if (style.n !== undefined && style.n !== null) data.numFmt = style.n.pattern
  const background = argbOf(style.bg)
  if (background !== undefined) {
    data.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: background } }
  }
  const border = writeBorder(style.bd)
  if (border !== undefined) data.border = border
  const alignment = writeAlignment(style)
  if (alignment !== undefined) data.alignment = alignment
  return Object.keys(data).length === 0 ? undefined : data
}

/**
 * Translate the Univer alignment style keys into an ExcelJS alignment.
 * @param style - Univer style data.
 * @returns ExcelJS alignment fields, or `undefined` when none are set.
 */
function writeAlignment(style: IStyleData): Partial<ExcelAlignment> | undefined {
  const alignment: Partial<ExcelAlignment> = {}
  if (style.ht !== undefined && style.ht !== null) {
    const horizontal = EXCEL_HORIZONTAL_BY_UNIVER[style.ht]
    if (horizontal !== undefined) alignment.horizontal = horizontal
  }
  if (style.vt !== undefined && style.vt !== null) {
    const vertical = EXCEL_VERTICAL_BY_UNIVER[style.vt]
    if (vertical !== undefined) alignment.vertical = vertical
  }
  if (style.tb === WRAP_STRATEGY_WRAP) alignment.wrapText = true
  if (style.stf === BOOLEAN_TRUE) alignment.shrinkToFit = true
  const rotation = writeRotation(style.tr)
  if (rotation !== undefined) alignment.textRotation = rotation
  return Object.keys(alignment).length === 0 ? undefined : alignment
}

/**
 * Map a Univer text rotation back to the ExcelJS alignment rotation.
 * @param rotation - Univer text rotation.
 * @returns Signed degrees, `'vertical'`, or `undefined` for horizontal text.
 */
function writeRotation(rotation: IStyleData['tr']): number | typeof VERTICAL_ROTATION | undefined {
  if (rotation === undefined || rotation === null) return undefined
  return rotation.v === BOOLEAN_TRUE ? VERTICAL_ROTATION : rotation.a
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

function writeLayout(worksheet: ExcelJS.Worksheet, sheet: Partial<IWorksheetData>, workbook: IWorkbookData): void {
  if (sheet.defaultRowHeight !== undefined) {
    worksheet.properties.defaultRowHeight = pixelsToPoints(sheet.defaultRowHeight)
  }
  if (sheet.defaultColumnWidth !== undefined) {
    worksheet.properties.defaultColWidth = pixelsToColumnWidth(sheet.defaultColumnWidth)
  }
  writeView(worksheet, sheet)
  if (sheet.tabColor !== undefined && sheet.tabColor !== '') {
    const tabColor = argbOf({ rgb: sheet.tabColor })
    if (tabColor !== undefined) worksheet.properties.tabColor = { argb: tabColor }
  }
  if (sheet.hidden !== undefined) {
    worksheet.state = sheet.hidden === VERY_HIDDEN
      ? 'veryHidden'
      : sheet.hidden === BOOLEAN_FALSE
        ? 'visible'
        : 'hidden'
  }
  for (const [rowKey, row] of entriesOf(sheet.rowData ?? {})) {
    const excelRow = worksheet.getRow(Number(rowKey) + 1)
    if (row.h !== undefined) excelRow.height = pixelsToPoints(row.h)
    if (row.hd === BOOLEAN_TRUE) excelRow.hidden = true
    const style = resolveStyle(workbook, row.s)
    if (style !== undefined) {
      const excelStyle = excelStyleOf(style)
      if (excelStyle !== undefined) styledRow(excelRow).style = excelStyle
    }
  }
  for (const [columnKey, column] of entriesOf(sheet.columnData ?? {})) {
    const excelColumn = worksheet.getColumn(Number(columnKey) + 1)
    if (column.w !== undefined) excelColumn.width = pixelsToColumnWidth(column.w)
    if (column.hd === BOOLEAN_TRUE) excelColumn.hidden = true
    const style = resolveStyle(workbook, column.s)
    if (style !== undefined) {
      const excelStyle = excelStyleOf(style)
      if (excelStyle !== undefined) excelColumn.style = excelStyle
    }
  }
}

/**
 * Write the sheet's one view, combining the gridline, direction, header,
 * zoom, and frozen-pane facts into a single replacement so no fact drops
 * another.
 * @param worksheet - Target worksheet.
 * @param sheet - Sheet snapshot being written.
 */
function writeView(worksheet: ExcelJS.Worksheet, sheet: Partial<IWorksheetData>): void {
  const shared: Partial<ExcelJS.WorksheetViewCommon> = {}
  if (sheet.showGridlines !== undefined) shared.showGridLines = sheet.showGridlines === BOOLEAN_TRUE
  if (sheet.rightToLeft !== undefined) shared.rightToLeft = sheet.rightToLeft === BOOLEAN_TRUE
  if (sheet.rowHeader?.hidden === BOOLEAN_TRUE || sheet.columnHeader?.hidden === BOOLEAN_TRUE) {
    shared.showRowColHeaders = false
  }
  if (sheet.zoomRatio !== undefined && sheet.zoomRatio !== DEFAULT_ZOOM_RATIO) {
    shared.zoomScale = sheet.zoomRatio
  }
  const freeze = sheet.freeze
  if (freeze !== undefined && (freeze.xSplit > 0 || freeze.ySplit > 0)) {
    worksheet.views = [{ ...shared, state: 'frozen', xSplit: freeze.xSplit, ySplit: freeze.ySplit }]
    return
  }
  if (Object.keys(shared).length > 0) worksheet.views = [{ ...shared, state: 'normal' }]
}

/**
 * ExcelJS keeps a row-local style behind `<row s>` but its typings omit the
 * runtime field that carries it.
 * @param row - ExcelJS row.
 * @returns The same row, with its runtime style field typed.
 */
function styledRow(row: ExcelJS.Row): ExcelJS.Row & { style?: Partial<ExcelStyle> } {
  return row
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
 * Parse one plugin resource the editor persisted with the snapshot.
 * @param data - Workbook snapshot.
 * @param name - Resource name registered by the feature plugin.
 * @returns The parsed resource, or `undefined` when the snapshot carries none.
 * @throws {OfficeExchangeError} `conversion-failed` when the resource JSON is malformed.
 */
function resourceOf<T>(data: IWorkbookData, name: string): T | undefined {
  const entry = data.resources?.find(resource => resource.name === name)
  if (entry === undefined) return undefined
  try {
    return JSON.parse(entry.data) as T
  } catch (error) {
    throw new OfficeExchangeError('conversion-failed', `Cannot read ${name} resource: ${messageOf(error)}`)
  }
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
