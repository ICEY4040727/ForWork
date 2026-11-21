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
 *
 * IN-PLACE SAVE: `readWorkbookWithSource` attaches
 * `data.officeExchange.source`, a per-cell anchor index over each resolved
 * worksheet part. `writeWorkbookFromSource` resolves worksheet parts through
 * `xl/workbook.xml` and `xl/_rels/workbook.xml.rels`, recomputes each cell's
 * digest, and splices only the edited `<c>` spans. A changed string becomes an
 * inline string so `xl/sharedStrings.xml` stays untouched, and row or column
 * insert/delete shifts are expressed as row or cell span operations. A
 * structural row or cell shift on a sheet that carries a merged,
 * conditional-formatting, data-validation, auto-filter, or table range refuses
 * because the range is not shifted; malformed table metadata also refuses
 * because the table's range cannot be verified. A cell whose modelled style changes reuses an existing `<xf>`
 * or appends the needed `<numFmt>`, `<font>`, `<fill>`, `<border>`, and `<xf>`
 * entries to `xl/styles.xml`, updating its `count` attributes; the part is
 * never regenerated. Rich-text shared strings and shared or array formulas
 * refuse with a typed error and produce no bytes, and a model style component
 * the writer cannot express refuses instead of being dropped.
 */
import { createHash } from 'node:crypto'
import { posix } from 'node:path'
import { SyntaxValidator } from 'fast-xml-validator'
import ExcelJS, {
  type Alignment as ExcelAlignment,
  type Border as ExcelBorder,
  type BorderStyle,
  type Borders,
  type Cell as ExcelCell,
  type CellValue as ExcelCellValue,
  type Color as ExcelColor,
  type Fill as ExcelFill,
  type Font as ExcelFont,
  type Image as ExcelImage,
  type Media as ExcelMedia,
  type Style as ExcelStyle,
} from 'exceljs'
import { XMLParser } from 'fast-xml-parser'
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
import { assertOpcGraph, open } from '../package/ooxml-package.ts'
import type { OfficePackage } from '../package/ooxml-package.ts'
import type { SourcePackage } from '../package/types.ts'
import { ZipError, bytesEqual } from '../package/zip-format.ts'
import { applySplices } from '../source/apply-splices.ts'
import type { ApplySpliceOperation } from '../source/apply-splices.ts'
import { buildSaveDisposition, byteOperationsByPart, sourceDispositions } from '../source/save-disposition.ts'
import { parseSnapshotSourceState } from '../source/snapshot-state.ts'
import { parseSourceState, SourceStateError } from '../source/source-state.ts'
import type { OfficeSourceFormat, OfficeSourceState, SheetSummary, UnitAnchor } from '../source/source-state.ts'
import { buildSplicePlan, classifyIndependentProperties } from '../source/splice.ts'
import type { CurrentUnit, SpliceOperation } from '../source/splice.ts'
import { digestSheetCell } from '../source/unit-digest.ts'
import {
  CDATA_CLOSE,
  CDATA_OPEN,
  COMMENT_CLOSE,
  COMMENT_OPEN,
  PROCESSING_CLOSE,
  declarationEnd,
  findSequence,
  findTagEnd,
  indexOfByte,
  isWhitespace,
  startsWith,
} from '../source/xml-scan.ts'
import { tagEndInText, isXmlTextWhitespace } from '../source/xml-text.ts'
import { scanSpans } from '../source/xml-spans.ts'
import type { ElementSpan } from '../source/xml-spans.ts'
import type { OfficeWriteResult } from '../snapshots.ts'

const UNIVER_APP_VERSION = '1.4.0'
const MAX_SOURCE_ENTRIES = 10_000
const WORKBOOK_PART = 'xl/workbook.xml'
const WORKBOOK_RELS_PART = 'xl/_rels/workbook.xml.rels'
const SHARED_STRINGS_PART = 'xl/sharedStrings.xml'
const STYLES_PART = 'xl/styles.xml'
const STYLE_XF_ENTRY = 'xf'
const STYLE_COLLECTION_ORDER = [
  'numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst',
] as const
const STYLE_TRACKED_COLLECTIONS: ReadonlySet<string> = new Set(['numFmts', 'fonts', 'fills', 'borders', 'cellXfs'])
const STYLE_UNMODELED_KEYS: Readonly<Record<string, string>> = {
  eastAsiaFontFamily: 'East Asian font family',
  bbl: 'bottom border line',
  ol: 'overline',
  td: 'text direction',
  pd: 'cell padding',
}
/**
 * Built-in number format codes as ExcelJS maps them. A code in this table can
 * use its built-in id instead of appending a custom `numFmt`.
 */
const BUILTIN_NUMBER_FORMATS: ReadonlyMap<string, number> = new Map([
  ['General', 0],
  ['0', 1],
  ['0.00', 2],
  ['#,##0', 3],
  ['#,##0.00', 4],
  ['0%', 9],
  ['0.00%', 10],
  ['0.00E+00', 11],
  ['# ?/?', 12],
  ['# ??/??', 13],
  ['mm-dd-yy', 14],
  ['d-mmm-yy', 15],
  ['d-mmm', 16],
  ['mmm-yy', 17],
  ['h:mm AM/PM', 18],
  ['h:mm:ss AM/PM', 19],
  ['h:mm', 20],
  ['h:mm:ss', 21],
  ['m/d/yy "h":mm', 22],
  ['#,##0 ;(#,##0)', 37],
  ['#,##0 ;[Red](#,##0)', 38],
  ['#,##0.00 ;(#,##0.00)', 39],
  ['#,##0.00 ;[Red](#,##0.00)', 40],
  ['mm:ss', 45],
  ['[h]:mm:ss', 46],
  ['mmss.0', 47],
  ['##0.0E+0', 48],
  ['@', 49],
])
const STYLE_NUMBER_FORMAT_GENERAL = 'General'
const STYLE_ALIGNMENT_TAG = 'alignment'
const STYLE_PROTECTION_TAG = 'protection'
const SPREADSHEET_NAMESPACES: ReadonlySet<string> = new Set([
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://purl.oclc.org/ooxml/spreadsheetml/main',
])
/** Worksheet elements that carry a cell range another edit could invalidate. */
const WORKSHEET_RANGE_KINDS: Readonly<Record<string, { readonly attribute: string; readonly construct: string }>> = {
  mergeCell: { attribute: 'ref', construct: 'merged range' },
  conditionalFormatting: { attribute: 'sqref', construct: 'conditional formatting range' },
  dataValidation: { attribute: 'sqref', construct: 'data validation range' },
  autoFilter: { attribute: 'ref', construct: 'auto filter range' },
}
const WORKBOOK_TEXT_DECODER = new TextDecoder()
const WORKBOOK_TEXT_ENCODER = new TextEncoder()
const XML_ERROR_VALUES: ReadonlySet<string> = new Set([
  '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A',
  '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#UNKNOWN!',
  '#CONNECT!', '#BUSY!',
])
const WORKSHEET_CONTENT_TYPES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const STRICT_WORKSHEET_CONTENT_TYPES = 'http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet'
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
const BORDER_STYLE_NONE = 0 as BorderStyleTypes
const HORIZONTAL_UNSPECIFIED = 0 as HorizontalAlign
const VERTICAL_UNSPECIFIED = 0 as VerticalAlign
const BASELINE_NORMAL = 1 as BaselineOffset
const VERTICAL_ROTATION = 'vertical'
const DEFAULT_ZOOM_RATIO = 100
const DETERMINISTIC_DATE = new Date('2000-01-01T00:00:00.000Z')
const EMU_PER_PIXEL = 9525
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

/** One border side an ExcelJS style exposes. */
type ExcelBorderSide = 'top' | 'right' | 'bottom' | 'left'

/**
 * Style fields read from an ExcelJS cell, row, column, or rule. ExcelJS
 * declares each block non-null, but omits it as `null` at runtime.
 */
interface ExcelStyleSource {
  readonly font?: Partial<ExcelFont> | null
  readonly numFmt?: string
  readonly fill?: ExcelFill | null
  readonly border?: Partial<Record<ExcelBorderSide, Partial<ExcelBorder> | null>> | null
  readonly alignment?: Partial<ExcelAlignment> | null
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

/** A workbook snapshot that carries the Exchange layer's source sidecar. */
interface OfficeExchangeWorkbookData extends IWorkbookData {
  /** Source package state the editable workbook can address. */
  officeExchange?: {
    /** Anchors that bind the editable workbook to its source package. */
    readonly source?: OfficeSourceState
  }
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
  return readWorkbookData(bytes, config)
}

/**
 * Read `.xlsx` or `.xlsm` bytes into a workbook snapshot that also carries the
 * `data.officeExchange.source` sidecar when worksheet parts and cells resolve.
 *
 * Per-cell anchors are used while they fit `maxSourceAnchors` and
 * `maxSourceAnchorBytes`; a larger workbook stores one sheet summary per sheet.
 * A container or worksheet the scanner refuses leaves `source` absent.
 * @param bytes - Complete workbook bytes.
 * @param format - Source format recorded in the sidecar.
 * @param config - Bounds applied before parsing and when validating anchors.
 * @returns Workbook snapshot with source anchors, or without a sidecar when
 * anchors cannot be built.
 * @throws {OfficeExchangeError} `too-large` when a bound is exceeded,
 * `conversion-failed` when the input is not a readable workbook.
 */
export async function readWorkbookWithSource(
  bytes: Uint8Array,
  format: 'xlsx' | 'xlsm',
  config?: OfficeExchangeConfig,
): Promise<IWorkbookData> {
  const bounds = resolveConfig(config ?? {})
  const data = await readWorkbookData(bytes, bounds)
  const source = buildWorkbookSourceState(bytes, data, format, bounds)
  if (source !== undefined) {
    const carrier = data as OfficeExchangeWorkbookData
    carrier.officeExchange = { source }
  }
  return data
}

async function readWorkbookData(bytes: Uint8Array, config?: OfficeExchangeConfig): Promise<IWorkbookData> {
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

/** The sidecar-independent identity of one worksheet the model sheet order names. */
interface WorkbookSheetSource {
  /** Model sheet id. */
  readonly sheetId: string
  /** Resolved worksheet part name. */
  readonly part: string
  /** Cells the model read from that worksheet, keyed by zero-based row and column. */
  readonly cellData: Readonly<Record<number, Readonly<Record<number, ICellData>>>>
}

/** One worksheet declared by `xl/workbook.xml`. */
interface WorkbookSheetEntry {
  /** Sheet display name. */
  readonly name: string
  /** `r:id` that resolves the worksheet part. */
  readonly relationshipId: string
}

/** One relationship declared by `xl/_rels/workbook.xml.rels`. */
interface WorkbookRelationship {
  /** Relationship id. */
  readonly id: string
  /** Relationship type URI. */
  readonly type: string
  /** Target part name relative to `xl/`. */
  readonly target: string
}

const WORKBOOK_PARSER = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: true,
  parseTagValue: false,
  parseAttributeValue: false,
  isArray: (name: string) => name === 'sheet' || name === 'Relationship',
})

/**
 * Validate workbook anchors against the source package and return the sidecar.
 * @param bytes - Source package bytes.
 * @param data - Workbook snapshot read from those bytes.
 * @param format - Source format recorded in the sidecar.
 * @param bounds - Validated anchor count and payload limits.
 * @returns The validated sidecar, or `undefined` when worksheet parts or
 * cells cannot be resolved unambiguously.
 */
function buildWorkbookSourceState(
  bytes: Uint8Array,
  data: IWorkbookData,
  format: OfficeSourceFormat,
  bounds: OfficeExchangeConfig,
): OfficeSourceState | undefined {
  let packageFile: ReturnType<typeof open>
  try {
    packageFile = open(bytes, { maxEntries: MAX_SOURCE_ENTRIES })
  } catch (error) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (error instanceof ZipError) return undefined
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw error
  }
  const sheets = resolveWorkbookSheets(packageFile, data)
  if (sheets === undefined) return undefined

  const anchors: UnitAnchor[] = []
  let perUnit = true
  for (const sheet of sheets) {
    let partBytes: Uint8Array
    try {
      partBytes = packageFile.readPart(sheet.part)
    } catch (error) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (error instanceof ZipError) return undefined
      /* v8 ignore next -- caller contract makes this path unreachable */
      throw error
    }
    const scan = scanSpans(partBytes, sheet.part, ['sheetCell'])
    if (scan.failure !== undefined) return undefined
    const spans = new Map<string, { readonly start: number; readonly end: number }>()
    for (const span of scan.spans) {
      const identity = span.identity
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (identity?.kind !== 'sheetCell') continue
      if (spans.has(identity.reference)) {
        perUnit = false
        break
      }
      spans.set(identity.reference, { start: span.start, end: span.end })
    }
    if (!perUnit) continue
    for (const [rowKey, cells] of sortedNumericEntries(sheet.cellData)) {
      const row = Number(rowKey)
      for (const [columnKey, cell] of sortedNumericEntries(cells)) {
        const column = Number(columnKey)
        const span = spans.get(cellAddress(row + 1, column + 1))
        if (span === undefined) {
          perUnit = false
          break
        }
        anchors.push({
          kind: 'sheetCell',
          part: sheet.part,
          key: `${sheet.sheetId}!${String(row)}:${String(column)}`,
          start: span.start,
          end: span.end,
          digest: digestSheetCell(cell),
        })
      }
      if (!perUnit) break
    }
  }

  const sourceParts = [...new Set(sheets.map(sheet => sheet.part))]
  const partByteLengths = new Map<string, number>()
  for (const entry of packageFile.source.entries) {
    if (!entry.name.endsWith('/')) partByteLengths.set(entry.name, entry.uncompressedSize)
  }
  const context = {
    partByteLengths,
    maxSourceAnchors: bounds.maxSourceAnchors,
    maxSourceAnchorBytes: bounds.maxSourceAnchorBytes,
  }
  try {
    if (
      perUnit
      && anchors.length <= bounds.maxSourceAnchors
      && Buffer.byteLength(JSON.stringify(anchors), 'utf8') <= bounds.maxSourceAnchorBytes
    ) {
      return parseSourceState({
        schemaVersion: 1,
        format,
        packageDigest: packageFile.source.digest,
        byteLength: bytes.byteLength,
        anchorMode: 'per-unit',
        anchors,
        sourceParts,
      }, context)
    }
    const summaries: SheetSummary[] = sheets.map(sheet => ({
      sheetId: sheet.sheetId,
      part: sheet.part,
      digest: digestSheet(sheet.cellData),
      cellCount: countSheetCells(sheet.cellData),
    }))
    return parseSourceState({
      schemaVersion: 1,
      format,
      packageDigest: packageFile.source.digest,
      byteLength: bytes.byteLength,
      anchorMode: 'sheet-summary',
      sheets: summaries,
      sourceParts,
    }, context)
  } catch (error) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (error instanceof SourceStateError) return undefined
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw error
  }
}

/**
 * Resolve each model sheet to its worksheet part through the workbook part and
 * its relationships.
 * @param packageFile - Indexed source package.
 * @param data - Workbook snapshot whose sheet order must match the workbook part.
 * @returns One source per model sheet in workbook order, or `undefined` when the
 * workbook metadata disagrees with the model.
 */
function resolveWorkbookSheets(
  packageFile: ReturnType<typeof open>,
  data: IWorkbookData,
): readonly WorkbookSheetSource[] | undefined {
  const workbookBytes = readOptionalPart(packageFile, WORKBOOK_PART)
  const relationshipsBytes = readOptionalPart(packageFile, WORKBOOK_RELS_PART)
  if (workbookBytes === undefined || relationshipsBytes === undefined) return undefined
  const entries = parseWorkbookSheets(workbookBytes)
  const relationships = parseWorkbookRelationships(relationshipsBytes)
  if (entries === undefined || relationships === undefined) return undefined
  if (entries.length !== data.sheetOrder.length) return undefined
  const sheets: WorkbookSheetSource[] = []
  for (const [index, entry] of entries.entries()) {
    const sheetId = data.sheetOrder[index]
    /* v8 ignore next -- the loop indexes sheetOrder within its length. */
    if (sheetId === undefined) return undefined
    const sheet = data.sheets[sheetId]
    if (sheet === undefined || sheet.name !== entry.name) return undefined
    const relationship = relationships.get(entry.relationshipId)
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (relationship === undefined || !isWorksheetRelationship(relationship.type)) return undefined
    const part = resolveWorkbookTarget(relationship.target)
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (part === undefined || !hasSourcePart(packageFile, part)) return undefined
    sheets.push({ sheetId, part, cellData: sheet.cellData ?? {} })
  }
  return sheets
}

function readOptionalPart(packageFile: ReturnType<typeof open>, name: string): Uint8Array | undefined {
  try {
    return packageFile.readPart(name)
  } catch (error) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (error instanceof ZipError) return undefined
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw error
  }
}

function hasSourcePart(packageFile: ReturnType<typeof open>, name: string): boolean {
  return packageFile.source.entries.some(entry => entry.name === name && !entry.name.endsWith('/'))
}

function parseWorkbookSheets(bytes: Uint8Array): readonly WorkbookSheetEntry[] | undefined {
  let parsed: unknown
  try {
    parsed = WORKBOOK_PARSER.parse(new TextDecoder().decode(bytes))
  } catch {
    /* v8 ignore next -- caller contract makes this path unreachable */
    return undefined
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (!isObjectRecord(parsed)) return undefined
  const workbook = parsed['workbook']
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (!isObjectRecord(workbook)) return undefined
  const sheets = workbook['sheets']
  if (!isObjectRecord(sheets)) return undefined
  const entries = arrayOf(sheets['sheet'])
  const result: WorkbookSheetEntry[] = []
  for (const entry of entries) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (!isObjectRecord(entry)) return undefined
    const name = entry['@_name']
    const relationshipId = entry['@_id']
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (typeof name !== 'string' || typeof relationshipId !== 'string') return undefined
    result.push({ name, relationshipId })
  }
  return result.length === 0 ? undefined : result
}

function parseWorkbookRelationships(bytes: Uint8Array): ReadonlyMap<string, WorkbookRelationship> | undefined {
  let parsed: unknown
  try {
    parsed = WORKBOOK_PARSER.parse(new TextDecoder().decode(bytes))
  } catch {
    /* v8 ignore next -- caller contract makes this path unreachable */
    return undefined
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (!isObjectRecord(parsed)) return undefined
  const root = parsed['Relationships']
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (!isObjectRecord(root)) return undefined
  const relationships = new Map<string, WorkbookRelationship>()
  for (const entry of arrayOf(root['Relationship'])) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (!isObjectRecord(entry)) return undefined
    const id = entry['@_Id']
    const type = entry['@_Type']
    const target = entry['@_Target']
    if (typeof id !== 'string' || typeof type !== 'string' || typeof target !== 'string') return undefined
    relationships.set(id, { id, type, target })
  }
  return relationships
}

function isWorksheetRelationship(type: string): boolean {
  return type === WORKSHEET_CONTENT_TYPES || type === STRICT_WORKSHEET_CONTENT_TYPES
}

function resolveWorkbookTarget(target: string): string | undefined {
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (target.includes('://')) return undefined
  const part = target.startsWith('/') ? target.slice(1) : posix.normalize(`xl/${target}`)
  /* v8 ignore next -- caller contract makes this path unreachable */
  return part.startsWith('../') ? undefined : part
}

function arrayOf(value: unknown): readonly unknown[] {
  // The workbook parser forces arrays for every name passed here.
  return value === undefined ? [] : value as readonly unknown[]
}

function isObjectRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

/**
 * Digest a sheet's canonical cell projection.
 * @param cellData - Cells keyed by zero-based row and column.
 * @returns Lowercase SHA-256 hex of the sorted row, column, and cell digests.
 */
function digestSheet(cellData: Readonly<Record<number, Readonly<Record<number, ICellData>>>>): string {
  const cells: Array<readonly [number, number, string]> = []
  for (const [rowKey, row] of sortedNumericEntries(cellData)) {
    const rowIndex = Number(rowKey)
    for (const [columnKey, cell] of sortedNumericEntries(row)) {
      cells.push([rowIndex, Number(columnKey), digestSheetCell(cell)])
    }
  }
  return createHash('sha256').update(JSON.stringify({ cells }), 'utf8').digest('hex')
}

function countSheetCells(cellData: Readonly<Record<number, Readonly<Record<number, ICellData>>>>): number {
  let count = 0
  for (const [, row] of sortedNumericEntries(cellData)) count += sortedNumericEntries(row).length
  return count
}

function sortedNumericEntries<T>(record: Readonly<Record<number, T>>): Array<[string, T]> {
  return Object.entries(record).sort((left, right) => Number(left[0]) - Number(right[0]))
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
    {},
  )
  const dataValidation = resourceOf<Record<string, DataValidationResourceRule[]>>(
    data,
    DATA_VALIDATION_RESOURCE,
    {},
  )
  const drawings = resourceOf<Record<string, DrawingResourceSheet>>(data, DRAWING_RESOURCE, {})
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
  const defaultRowHeight = parsed(properties.defaultRowHeight)
  const rowHeightPoints = defaultRowHeight !== undefined && Number.isFinite(defaultRowHeight)
    ? defaultRowHeight
    : 15
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
    defaultRowHeight: pointsToPixels(rowHeightPoints),
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
  // ExcelJS only attaches a hyperlink value for a non-empty relationship target.
  const hyperlinkValue = value as { readonly hyperlink: string; readonly text?: string }
  const hyperlink = hyperlinkValue.hyperlink
  const text = typeof hyperlinkValue.text === 'string' ? hyperlinkValue.text : hyperlink
  return {
    id: `cell-link-${cell.row}-${cell.col}`,
    documentStyle: {},
    body: {
      dataStream: `${text}\r\n`,
      paragraphs: [{ startIndex: text.length, paragraphId: 'cell-paragraph' }],
      customRanges: [{
        startIndex: 0,
        endIndex: Math.max(0, text.length - 1),
        rangeId: `hyperlink-${cell.row}-${cell.col}`,
        rangeType: HYPERLINK_RANGE_TYPE,
        // ExcelJS drops cell hyperlink tooltips from its value model.
        properties: { url: hyperlink },
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
      // ExcelJS normalizes every text rule to a string operator.
      const operator = rule.operator as string
      const text = textOfRule(formula)
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
      const text = textOfRule(formula)
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
      // ExcelJS initializes `cfvo` for every color-scale rule it parses.
      const thresholds = rule.cfvo as readonly ExcelConditionalValue[]
      const config = thresholds.map((threshold, index) => ({
        index,
        color: colorStringOf(colors[index] ?? colors[0]),
        value: conditionalValueOf(threshold),
      }))
      return { type: CF_COLOR_SCALE, config }
    }
    case 'dataBar': {
      // The ExcelJS parser omits data-bar display attributes; its writer defaults them to true.
      const thresholds = rule.cfvo as readonly ExcelConditionalValue[]
      const color = colorStringOf(colorArrayOf(rule.color)[0])
      return {
        type: CF_DATA_BAR,
        isShowValue: true,
        config: {
          min: conditionalValueOf(thresholds[0]),
          max: conditionalValueOf(thresholds[1]),
          isGradient: true,
          positiveColor: color,
          nativeColor: color,
        },
      }
    }
    case 'iconSet': {
      const iconType = iconTypeOf(rule.iconSet)
      const thresholds = rule.cfvo as readonly ExcelConditionalValue[]
      const config = thresholds.map((threshold, index) => ({
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
  if (operator === undefined) return `${target}=${first}`
  return `${target}${operator}${first}`
}

/**
 * Extract the quoted search text a text rule formula carries.
 * @param formula - First formula of the rule.
 * @returns The text the rule matches, or `undefined` when it is not recoverable.
 */
function textOfRule(formula: string | undefined): string | undefined {
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
  return typeof value === 'string' ? value : undefined
}

/** @returns A rule value as a finite number, or `undefined` when it is not a number. */
function numberOf(value: unknown): number | undefined {
  if (typeof value !== 'string' || value.trim().length === 0) return undefined
  const numeric = Number(value)
  return Number.isFinite(numeric) ? numeric : undefined
}

/** @returns The relative A1 reference of a range's first cell. */
function relativeReferenceOf(ref: string): string {
  const trimmed = ref.trim()
  const whitespace = trimmed.search(/\s/u)
  const first = whitespace < 0 ? trimmed : trimmed.slice(0, whitespace)
  const colon = first.indexOf(':')
  const start = colon < 0 ? first : first.slice(0, colon)
  return start.replaceAll('$', '') || 'A1'
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
  return isColorArray(color) ? color : [color]
}

/** @returns Whether a rule color field carries several colors rather than one. */
function isColorArray(
  color: Partial<ExcelColor> | readonly Partial<ExcelColor>[],
): color is readonly Partial<ExcelColor>[] {
  return Array.isArray(color)
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
  // After the date and number cases, ExcelJS only produces string criteria.
  const text = formula as string
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
    const medium = workbook.getImage(Number(image.imageId)) as ExcelMedia
    const drawingId = `sheet-drawing-${String(sheetIndex + 1)}-${String(index + 1)}`
    const range = image.range as ExcelImageRange
    const from = drawingPositionOf(anchorModelOf(range.tl))
    const to = range.br === undefined
      ? advanceAnchor(from, range.ext.width, range.ext.height, worksheet)
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
        editAs: range.editAs,
        ...(range.ext === undefined ? {} : { ext: { width: range.ext.width, height: range.ext.height } }),
        ...hyperlink,
      },
    }
    order.push(drawingId)
  }
  return { data, order }
}

/** A drawing range as ExcelJS exposes it; its parser always supplies `editAs` and an anchor's own extent. */
type ExcelImageRange = {
  readonly tl: ExcelJS.Anchor
  readonly editAs: string
  readonly hyperlinks?: { readonly hyperlink: string; readonly tooltip?: string }
} & (
  | {
    readonly br: ExcelJS.Anchor
    readonly ext?: { readonly width: number; readonly height: number }
  }
  | {
    readonly br?: undefined
    readonly ext: { readonly width: number; readonly height: number }
  }
)

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
function dataUrlOf(medium: ExcelMedia): string {
  const bytes = new Uint8Array(medium.buffer)
  return `data:${mimeOfExtension(medium.extension)};base64,${Buffer.from(bytes).toString('base64')}`
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
  const medium = workbook.getImage(Number(imageId)) as ExcelMedia
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
    const startColumn = columnIndexOf(match[1] as string)
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
  readonly model: Record<string, ExcelDataValidation>
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
  // Both capture groups are mandatory; the payload may be empty.
  return { mime: match[1] as string, base64: match[2] as string }
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

function readStyle(source: ExcelStyleSource): IStyleData | undefined {
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

function readBorder(
  border: Partial<Record<ExcelBorderSide, Partial<ExcelBorder> | null>> | undefined,
): IBorderData | undefined {
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
  const url: unknown = range?.properties?.['url']
  if (typeof url !== 'string' || url.length === 0) return undefined
  const text = cellTextOf(data)
  const tooltip: unknown = range?.properties?.['tooltip']
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
 * @param fallback - Value returned when the snapshot carries no such resource.
 * @returns The parsed resource, or `fallback` when the snapshot carries none.
 * @throws {OfficeExchangeError} `conversion-failed` when the resource JSON is malformed.
 */
function resourceOf<T>(data: IWorkbookData, name: string, fallback: T): T {
  const entry = data.resources?.find(resource => resource.name === name)
  if (entry === undefined) return fallback
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

/* ------------------------------------------------------------------------- *
 * In-place XLSX save
 * ------------------------------------------------------------------------- */

/** One worksheet cell address, zero-based. */
interface WorksheetAddress {
  /** Zero-based row index. */
  readonly row: number
  /** Zero-based column index. */
  readonly column: number
}

/** One attribute of a worksheet start tag with its source byte range. */
interface WorksheetTagAttribute {
  /** Attribute name as written, without a namespace prefix. */
  readonly name: string
  /** Decoded attribute value. */
  readonly value: string
  /** Inclusive byte offset of the value's first character. */
  readonly valueStart: number
  /** Exclusive byte offset of the value's last character. */
  readonly valueEnd: number
}

/** One parsed worksheet start tag. */
interface WorksheetStartTag {
  /** Qualified name as written. */
  readonly qname: string
  /** Inclusive byte offset of the opening `<`. */
  readonly start: number
  /** Exclusive byte offset immediately after the opening tag's `>`. */
  readonly openEnd: number
  /** Whether the tag ends in `/>`. */
  readonly selfClosing: boolean
  /** Attributes in source order. */
  readonly attributes: readonly WorksheetTagAttribute[]
}

/** A formula element inside a source cell. */
interface WorksheetFormula {
  /** Formula text as stored without a leading `=`. */
  readonly text: string
  /** Formula type from the `t` attribute, when present. */
  readonly type: string | undefined
}

/** An inline string element inside a source cell. */
interface WorksheetInline {
  /** Concatenated text of the element's `t` children. */
  readonly text: string
  /** Whether the element carries rich-text runs. */
  readonly rich: boolean
}

/** One source worksheet cell and its representation. */
interface WorksheetSourceCell {
  /** Zero-based address derived from `r` or sequential position. */
  readonly address: WorksheetAddress
  /** Inclusive byte offset of the cell's opening `<`. */
  readonly start: number
  /** Exclusive byte offset after the cell's last `>`. */
  readonly end: number
  /** Qualified cell name as written. */
  readonly qname: string
  /** Whether the cell states its `r` reference. */
  readonly hasReference: boolean
  /** The `r` attribute, when present. */
  readonly referenceAttribute: WorksheetTagAttribute | undefined
  /** The `s` style index, when present. */
  readonly style: string | undefined
  /** The `t` value-representation type, when present. */
  readonly type: string | undefined
  /** Formula child, when present. */
  readonly formula: WorksheetFormula | undefined
  /** Cached `<v>` text, when present. */
  readonly value: string | undefined
  /** Inline string child, when present. */
  readonly inline: WorksheetInline | undefined
}

/** One `<row>` element in a source worksheet. */
interface WorksheetSourceRow {
  /** Zero-based row index. */
  readonly index: number
  /** Inclusive byte offset of the row's opening `<`. */
  readonly start: number
  /** Exclusive byte offset after the row's last `>`. */
  readonly end: number
  /** Exclusive byte offset immediately after the row's opening tag. */
  readonly openEnd: number
  /** Inclusive byte offset of `</row>`, or `undefined` for a self-closing row. */
  readonly closeStart: number | undefined
  /** The row's start tag. */
  readonly tag: WorksheetStartTag
  /** Cells in source order. */
  readonly cells: readonly WorksheetSourceCell[]
}

/** One `dimension` element in a source worksheet. */
interface WorksheetDimension {
  /** Inclusive byte offset of the element's opening `<`. */
  readonly start: number
  /** Exclusive byte offset after the element's last `>`. */
  readonly end: number
  /** The `ref` attribute, when present. */
  readonly referenceAttribute: WorksheetTagAttribute | undefined
}

/** An address-bearing worksheet range a structural shift cannot move. */
interface WorksheetRange {
  /** Named construct that owns the range. */
  readonly construct: string
  /** Range reference as written in the source. */
  readonly reference: string
}

/** A worksheet table part whose range a structural shift cannot keep current. */
interface WorksheetTableFailure {
  /** Relationship id named by the worksheet's `tablePart` element. */
  readonly relationshipId: string
  /** Table part the relationship resolves to inside the package, when it does. */
  readonly part: string | undefined
}

/** Resolvable table ranges and declared table parts that produced none. */
interface WorksheetTableScan {
  /** One range per table part that carries a readable `ref`. */
  readonly ranges: readonly WorksheetRange[]
  /** Declared table parts whose range could not be read. */
  readonly unresolved: readonly WorksheetTableFailure[]
}

/** A source worksheet scanned into rows, cells, and its used-range declaration. */
interface WorksheetScan {
  /** Namespace prefix bound to the SpreadsheetML main namespace. */
  readonly prefix: string
  /** Row elements in source order. */
  readonly rows: readonly WorksheetSourceRow[]
  /** Cells in source order. */
  readonly cells: readonly WorksheetSourceCell[]
  /** Used-range declaration, when present. */
  readonly dimension: WorksheetDimension | undefined
  /** Merged, conditional-formatting, data-validation, and auto-filter ranges in source order. */
  readonly ranges: readonly WorksheetRange[]
  /** Relationship ids named by `tablePart` elements in source order. */
  readonly tablePartIds: readonly string[]
  /** Cell spans reported by `xml-spans`. */
  readonly spans: readonly ElementSpan[]
}

/** One model cell with its canonical digest. */
interface WorksheetModelCell {
  /** Zero-based address. */
  readonly address: WorksheetAddress
  /** A1 reference. */
  readonly reference: string
  /** Editable cell data. */
  readonly data: ICellData
  /** Canonical cell digest. */
  readonly digest: string
}

/** One model row and its cells. */
interface WorksheetModelRow {
  /** Zero-based row index. */
  readonly index: number
  /** Cells in column order. */
  readonly cells: readonly WorksheetModelCell[]
}

/** A source row whose content can be compared with the model. */
interface WorksheetSourceRowEntry {
  /** Zero-based row index. */
  readonly index: number
  /** Source row element. */
  readonly row: WorksheetSourceRow
  /** Column and digest pairs for modeled cells. */
  readonly cells: readonly (readonly [number, string])[]
  /** Digest of the cells alone, independent of the row index. */
  readonly signature: string
  /** Digest of the row index and cells. */
  readonly digest: string
}

/** A model row whose content can be compared with the source. */
interface WorksheetModelRowEntry {
  /** Zero-based row index. */
  readonly index: number
  /** Model row. */
  readonly row: WorksheetModelRow
  /** Column and digest pairs. */
  readonly cells: readonly (readonly [number, string])[]
  /** Digest of the cells alone, independent of the row index. */
  readonly signature: string
  /** Digest of the row index and cells. */
  readonly digest: string
}

/** Source-to-model row alignment for one worksheet. */
interface WorksheetRowAlignment {
  /** Source rows paired with their model row. */
  readonly pairs: readonly {
    readonly source: WorksheetSourceRowEntry
    readonly model: WorksheetModelRowEntry
  }[]
  /** Source rows the model no longer carries. */
  readonly deleted: readonly WorksheetSourceRowEntry[]
  /** Model rows the source does not carry. */
  readonly inserted: readonly WorksheetModelRowEntry[]
}

/** Everything one worksheet part needs across an in-place save. */
interface WorksheetPartState {
  /** Model sheet id. */
  readonly sheetId: string
  /** Resolved decoded part name. */
  readonly part: string
  /** Uncompressed source bytes. */
  readonly bytes: Uint8Array
  /** SpreadsheetML namespace prefix used by the part. */
  readonly prefix: string
  /** Scanned rows. */
  readonly rows: readonly WorksheetSourceRow[]
  /** Scanned cells. */
  readonly cells: readonly WorksheetSourceCell[]
  /** Scanned used-range declaration. */
  readonly dimension: WorksheetDimension | undefined
  /** Address-bearing ranges the splice cannot shift. */
  readonly ranges: readonly WorksheetRange[]
  /** Cell spans from `xml-spans`. */
  readonly spans: readonly ElementSpan[]
  /** Source cells the reader mapped, keyed by `row:column`. */
  readonly sourceCells: ReadonlyMap<string, ICellData>
  /** Source rows with modeled cells, in row order. */
  readonly sourceRows: readonly WorksheetSourceRowEntry[]
  /** Model rows with cells, in row order. */
  readonly modelRows: readonly WorksheetModelRow[]
  /** Model cells in row then column order. */
  readonly modelCells: readonly WorksheetModelCell[]
  /** Whether address shifts require row-level span operations. */
  readonly rowLevel: boolean
  /** Plan anchors for this part. */
  readonly anchors: readonly UnitAnchor[]
  /** Source-to-model row alignment for a row-level part. */
  readonly alignment: WorksheetRowAlignment | undefined
}

/** Shared state for one in-place workbook save. */
interface WorksheetSpliceState {
  /** Opened source package. */
  readonly packageFile: OfficePackage
  /** Source package as opened. */
  readonly source: SourcePackage
  /** Validated sidecar. */
  readonly sourceState: OfficeSourceState
  /** Validated bounds. */
  readonly bounds: OfficeExchangeConfig
  /** Worksheet parts in model sheet order. */
  readonly sheets: readonly WorksheetPartState[]
  /** Worksheet parts keyed by decoded part name. */
  readonly partsByName: ReadonlyMap<string, WorksheetPartState>
  /** Shared-string entries, or `undefined` when the package has none or they cannot be read. */
  readonly sharedStrings: readonly SharedStringEntry[] | undefined
  /** Workbook style table referenced by string style ids in the snapshot. */
  readonly workbookStyles: IWorkbookData['styles']
  /** Lazily loaded `xl/styles.xml` edit plan. */
  readonly styles: WorksheetStylesHolder
}

/** Lazy holder for the source style table, which is parsed only when a style changes. */
interface WorksheetStylesHolder {
  /** Whether loading was attempted. */
  loaded: boolean
  /** Parsed table and pending appends, or `undefined` when the package has no usable table. */
  value: WorksheetStylesState | undefined
}

/** Whether one shared-string entry uses rich runs. */
interface SharedStringEntry {
  /** True when the entry carries `r` runs instead of one plain `t`. */
  readonly rich: boolean
}

/**
 * Write a workbook snapshot back into the package it was read from by splicing
 * only the cells the edit changed.
 *
 * Cells whose digest still matches the sidecar are copied from the source
 * bytes. A changed number, boolean, error, or formula keeps its existing
 * representation and the cell's `s` style index; a changed string is written as
 * an inline string so `xl/sharedStrings.xml` stays untouched. Row and column
 * insert or delete shifts are expressed as row or cell span operations, and a
 * changed used range updates the worksheet's `dimension` reference in place.
 * A shift that would move a merged, conditional-formatting, data-validation,
 * auto-filter, or table range refuses instead of leaving the range at stale
 * coordinates. A declared table part whose range cannot be read refuses for
 * the same reason.
 *
 * A rich-text shared string, a shared or array formula, a cell style change, or
 * an unmapped cell attribute the serializer cannot reproduce raises a typed
 * refusal and produces no bytes.
 * @param snapshot - Editable workbook snapshot carrying `officeExchange.source`.
 * @param source - Complete source package bytes read at the caller's expected version.
 * @param config - Bounds validated before splicing; omitted fields keep their defaults.
 * @param signal - Optional cancellation checked before, during, and after splicing.
 * @returns Complete output bytes and the per-entry save disposition.
 * @throws {OfficeExchangeError} `too-large` when the source, model cell count,
 * or output exceeds a byte or cell bound; `unsupported-feature` when a coupled
 * construct overlaps an edited cell, a cell style would change, or a structural
 * shift would move an address-bearing range or meets a declared table part
 * with no readable range; `conversion-failed` when the
 * snapshot has no usable source state, the source bytes do not match it, a span
 * cannot be located, or a spliced part is not well-formed.
 */
export function writeWorkbookFromSource(
  snapshot: IWorkbookData,
  source: Uint8Array,
  config?: OfficeExchangeConfig,
  signal?: AbortSignal,
): Promise<OfficeWriteResult> {
  return Promise.resolve().then(async () => {
    const bounds = resolveConfig(config ?? {})
    return writeWorkbookSourceBytes(snapshot, source, bounds, signal)
  })
}

async function writeWorkbookSourceBytes(
  data: IWorkbookData,
  source: Uint8Array,
  bounds: OfficeExchangeConfig,
  signal: AbortSignal | undefined,
): Promise<OfficeWriteResult> {
  signal?.throwIfAborted()
  if (source.byteLength > bounds.maxFileBytes) {
    throw new OfficeExchangeError(
      'too-large',
      `Source workbook is ${String(source.byteLength)} bytes; the limit is ${String(bounds.maxFileBytes)}`,
    )
  }
  const state = await createWorksheetSpliceState(data, source, bounds, signal)
  signal?.throwIfAborted()

  const plan = buildSplicePlan({
    source: state.source,
    sourceState: state.sourceState,
    anchors: state.sheets.flatMap(sheet => sheet.anchors),
    units: state.sheets.flatMap(sheet => (
      sheet.rowLevel ? worksheetRowUnits(sheet) : worksheetCellUnits(sheet)
    )),
    spans: state.sheets.flatMap(sheet => (
      sheet.rowLevel
        ? sheet.rows.map(row => ({
          kind: 'sheetCell' as const,
          part: sheet.part,
          start: row.start,
          end: row.end,
        }))
        : sheet.spans.map(span => ({
          kind: 'sheetCell' as const,
          part: sheet.part,
          start: span.start,
          end: span.end,
        }))
    )),
    renderUnit: (unit) => {
      const sheet = state.partsByName.get(unit.part)
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (sheet === undefined) {
        /* v8 ignore next -- caller contract makes this path unreachable */
        throw new OfficeExchangeError('conversion-failed', `Unit ${unit.key} names unknown part ${unit.part}`)
      }
      return sheet.rowLevel
        ? renderWorksheetRowUnit(state, sheet, unit)
        : renderWorksheetCellUnit(state, sheet, unit)
    },
  })
  if (plan.refusal !== undefined) {
    throw new OfficeExchangeError('conversion-failed', plan.refusal.detail)
  }
  const applied = applyWorksheetOperations(
    state,
    sortWorksheetOperations([
      ...worksheetMeaningfulOperations(state, plan.operations),
      ...dimensionOperations(state),
      ...worksheetStyleOperations(state.styles),
    ]),
    signal,
  )
  return {
    bytes: applied.bytes,
    disposition: buildSaveDisposition(sourceDispositions(state.source, applied.operations)),
  }
}

function worksheetCellUnits(sheet: WorksheetPartState): readonly CurrentUnit[] {
  return sheet.modelCells.map(cell => ({
    kind: 'sheetCell',
    part: sheet.part,
    key: `${sheet.sheetId}!${String(cell.address.row)}:${String(cell.address.column)}`,
    digest: cell.digest,
  }))
}

function worksheetRowUnits(sheet: WorksheetPartState): readonly CurrentUnit[] {
  const alignment = sheet.alignment
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (alignment === undefined) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw new OfficeExchangeError('conversion-failed', `Row-level state for ${sheet.part} has no row alignment`)
  }
  return sheet.modelRows.map(row => ({
    kind: 'sheetCell',
    part: sheet.part,
    key: worksheetRowKey(sheet.sheetId, row.index),
    digest: worksheetRowDigest(row.index, row.cells.map(cell => [cell.address.column, cell.digest] as const)),
  }))
}

function worksheetRowKey(sheetId: string, index: number): string {
  return `${sheetId}!${String(index)}`
}

/**
 * Report whether a sidecar describes another package than the one being saved.
 *
 * A stale sidecar keeps offsets and digests from the package the snapshot was
 * read from, so its spans must not be bounded by the current part lengths.
 * @param value - Untrusted `data.officeExchange.source` value.
 * @param source - Package opened for this save.
 * @returns `true` when the recorded package digest or byte length differs.
 */
function worksheetSnapshotStateIsStale(value: unknown, source: SourcePackage): boolean {
  if (!isObjectRecord(value)) return false
  return value['packageDigest'] !== source.digest || value['byteLength'] !== source.bytes.byteLength
}

/**
 * Replace every part length with an unbounded value while keeping the part-name
 * set, so a stale sidecar's spans can be validated structurally but not
 * against the current package's byte lengths.
 * @param source - Package opened for this save.
 * @returns A package whose recorded part lengths accept any sidecar span.
 */
function unboundedWorksheetSource(source: SourcePackage): SourcePackage {
  return {
    ...source,
    entries: source.entries.map(entry => ({ ...entry, uncompressedSize: Number.MAX_SAFE_INTEGER })),
  }
}

async function createWorksheetSpliceState(
  data: IWorkbookData,
  source: Uint8Array,
  bounds: OfficeExchangeConfig,
  signal: AbortSignal | undefined,
): Promise<WorksheetSpliceState> {
  const packageFile = openWorksheetPackage(source)
  const value = workbookSourceStateValue(data)
  if (value === undefined) {
    throw new OfficeExchangeError(
      'conversion-failed',
      'Workbook snapshot carries no source anchors; use the regeneration writer',
    )
  }
  const stale = worksheetSnapshotStateIsStale(value, packageFile.source)
  let sourceState = parseSnapshotSourceState(
    value,
    stale ? unboundedWorksheetSource(packageFile.source) : packageFile.source,
    bounds,
    'Workbook',
  )
  if (sourceState.format !== 'xlsx' && sourceState.format !== 'xlsm') {
    throw new OfficeExchangeError(
      'unsupported-format',
      `Source state names a ${sourceState.format} package, not a workbook`,
    )
  }
  signal?.throwIfAborted()
  const sourceWorkbook = await readWorkbookData(source, bounds)
  if (stale) {
    const rebuilt = buildWorkbookSourceState(source, sourceWorkbook, sourceState.format, bounds)
    if (rebuilt === undefined) {
      throw new OfficeExchangeError(
        'conversion-failed',
        'The current source cannot be re-anchored for an in-place save',
      )
    }
    sourceState = rebuilt
  }
  signal?.throwIfAborted()
  const sources = resolveWorkbookSheets(packageFile, data)
  if (sources === undefined) {
    throw new OfficeExchangeError(
      'conversion-failed',
      'The workbook part does not declare the sheets the snapshot names',
    )
  }
  if (sourceWorkbook.sheetOrder.length !== sources.length) {
    throw new OfficeExchangeError('conversion-failed', 'The source workbook sheet count changed after it was read')
  }
  const sourceParts = new Set(sourceState.sourceParts)
  const anchorsByPart = new Map<string, UnitAnchor[]>()
  if (sourceState.anchorMode === 'per-unit') {
    for (const anchor of sourceState.anchors) {
      if (anchor.kind !== 'sheetCell') {
        throw new OfficeExchangeError(
          'conversion-failed',
          `Workbook anchor ${anchor.key} names a ${anchor.kind} in ${anchor.part}`,
        )
      }
      const list = anchorsByPart.get(anchor.part)
      if (list === undefined) anchorsByPart.set(anchor.part, [anchor])
      else list.push(anchor)
    }
  }
  const sheets: WorksheetPartState[] = []
  const seenParts = new Set<string>()
  for (const [index, resolved] of sources.entries()) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (seenParts.has(resolved.part)) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      throw new OfficeExchangeError('conversion-failed', `Two sheets resolve to the same part ${resolved.part}`)
    }
    seenParts.add(resolved.part)
    if (!sourceParts.has(resolved.part)) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `Source state does not name worksheet part ${resolved.part}`,
      )
    }
    const sourceSheetId = sourceWorkbook.sheetOrder[index]
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (sourceSheetId === undefined) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      throw new OfficeExchangeError('conversion-failed', 'The source workbook has fewer sheets than the snapshot')
    }
    const sourceSheet = sourceWorkbook.sheets[sourceSheetId]
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (sourceSheet === undefined) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      throw new OfficeExchangeError('conversion-failed', `The source workbook has no sheet ${sourceSheetId}`)
    }
    const bytes = readWorksheetPart(packageFile, resolved.part)
    const scan = scanWorksheetPart(bytes, resolved.part)
    const sourceCells = worksheetSourceCellMap(sourceSheet)
    if (sourceState.anchorMode === 'sheet-summary') {
      verifyWorksheetSummary(sourceState.sheets, resolved.sheetId, resolved.part, sourceSheet)
    }
    const modelRows = worksheetModelRows(resolved.sheetId, data, bounds)
    const modelCells = modelRows.flatMap(row => row.cells)
    const sourceRows = worksheetSourceRows(scan, sourceCells)
    const rowLevel = worksheetNeedsRowLevel(sourceRows, modelRows)
    const alignment = rowLevel
      ? alignWorksheetRows(sourceRows, modelRows)
      : undefined
    const tableScan = worksheetTableRanges(packageFile, resolved.part, scan.tablePartIds)
    const ranges = [...scan.ranges, ...tableScan.ranges]
    const range = ranges[0]
    if (alignment !== undefined) {
      const shiftAlignment = alignWorksheetRows(
        sourceRows,
        worksheetModelRowsWithoutMergeSlaves(modelRows, worksheetMergeSlaves(data, resolved.sheetId)),
      )
      if (worksheetShiftsAddresses(shiftAlignment)) {
        const unresolved = tableScan.unresolved[0]
        if (unresolved !== undefined) {
          const target = unresolved.part === undefined ? '' : ` (${unresolved.part})`
          throw new OfficeExchangeError(
            'unsupported-feature',
            `table relationship ${unresolved.relationshipId}${target} in ${resolved.part}`
            + ' has no resolvable table range; the worksheet cannot follow an insert, delete, or move of rows or cells',
          )
        }
        if (range !== undefined) {
          throw new OfficeExchangeError(
            'unsupported-feature',
            `${range.construct} ${range.reference} in ${resolved.part} cannot follow an insert, delete, or move of rows or cells`,
          )
        }
      }
    }
    const anchors = alignment !== undefined
      ? worksheetRowAnchors(resolved.sheetId, resolved.part, alignment)
      : worksheetCellAnchors(
        resolved.sheetId,
        resolved.part,
        scan,
        sourceCells,
        anchorsByPart.get(resolved.part) ?? [],
        sourceState.anchorMode,
      )
    sheets.push({
      sheetId: resolved.sheetId,
      part: resolved.part,
      bytes,
      prefix: scan.prefix,
      rows: scan.rows,
      cells: scan.cells,
      dimension: scan.dimension,
      ranges,
      spans: scan.spans,
      sourceCells,
      sourceRows,
      modelRows,
      modelCells,
      rowLevel,
      anchors,
      alignment,
    })
  }
  signal?.throwIfAborted()
  return {
    packageFile,
    source: packageFile.source,
    sourceState,
    bounds,
    sheets,
    partsByName: new Map(sheets.map(sheet => [sheet.part, sheet])),
    sharedStrings: loadWorksheetSharedStrings(packageFile),
    workbookStyles: data.styles,
    styles: { loaded: false, value: undefined },
  }
}

function workbookSourceStateValue(data: IWorkbookData): unknown {
  const extension: unknown = (data as OfficeExchangeWorkbookData).officeExchange
  if (extension === undefined) return undefined
  if (!isObjectRecord(extension)) {
    throw new OfficeExchangeError('conversion-failed', 'Workbook snapshot officeExchange state is malformed')
  }
  return extension['source']
}

function openWorksheetPackage(source: Uint8Array): OfficePackage {
  try {
    return open(source, { maxEntries: MAX_SOURCE_ENTRIES })
  } catch (error) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (error instanceof ZipError) {
      throw new OfficeExchangeError('conversion-failed', `Cannot open the source package: ${messageOf(error)}`)
    }
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw error
  }
}

function readWorksheetPart(packageFile: OfficePackage, part: string): Uint8Array {
  try {
    return packageFile.readPart(part)
  } catch (error) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (error instanceof ZipError) {
      throw new OfficeExchangeError('conversion-failed', `Cannot read worksheet part ${part}: ${messageOf(error)}`)
    }
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw error
  }
}

function verifyWorksheetSummary(
  summaries: readonly SheetSummary[],
  sheetId: string,
  part: string,
  sourceSheet: Partial<IWorksheetData>,
): void {
  const summary = summaries.find(candidate => candidate.sheetId === sheetId)
  if (summary === undefined) {
    throw new OfficeExchangeError('conversion-failed', `Source state has no summary for sheet ${sheetId}`)
  }
  if (summary.part !== part) {
    throw new OfficeExchangeError(
      'conversion-failed',
      `Sheet ${sheetId} summary names ${summary.part}, but the workbook declares ${part}`,
    )
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  const cellData = sourceSheet.cellData ?? {}
  if (summary.digest !== digestSheet(cellData) || summary.cellCount !== countSheetCells(cellData)) {
    throw new OfficeExchangeError(
      'conversion-failed',
      `The source worksheet ${part} does not match its recorded summary; the file changed after it was read`,
    )
  }
}

function worksheetSourceCellMap(sheet: Partial<IWorksheetData>): ReadonlyMap<string, ICellData> {
  const cells = new Map<string, ICellData>()
  /* v8 ignore next -- caller contract makes this path unreachable */
  for (const [rowKey, row] of sortedNumericEntries(sheet.cellData ?? {})) {
    for (const [columnKey, cell] of sortedNumericEntries(row)) {
      cells.set(`${rowKey}:${columnKey}`, cell)
    }
  }
  return cells
}

function worksheetModelRows(
  sheetId: string,
  data: IWorkbookData,
  bounds: OfficeExchangeConfig,
): readonly WorksheetModelRow[] {
  const sheet = data.sheets[sheetId]
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (sheet === undefined) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw new OfficeExchangeError('conversion-failed', `Workbook snapshot is missing sheet ${sheetId}`)
  }
  const rows = new Map<number, WorksheetModelCell[]>()
  let count = 0
  for (const [rowKey, row] of sortedNumericEntries(sheet.cellData ?? {})) {
    const rowIndex = Number(rowKey)
    for (const [columnKey, cell] of sortedNumericEntries(row)) {
      if (worksheetCellIsEmpty(cell)) continue
      count += 1
      if (count > bounds.maxCells) {
        throw new OfficeExchangeError('too-large', `Workbook has more than ${bounds.maxCells} cells`)
      }
      const column = Number(columnKey)
      const list = rows.get(rowIndex)
      const modelCell: WorksheetModelCell = {
        address: { row: rowIndex, column },
        reference: cellAddress(rowIndex + 1, column + 1),
        data: cell,
        digest: digestSheetCell(cell),
      }
      if (list === undefined) rows.set(rowIndex, [modelCell])
      else list.push(modelCell)
    }
  }
  return [...rows.entries()]
    .sort((left, right) => left[0] - right[0])
    .map(([index, cells]) => ({
      index,
      cells: cells.sort((left, right) => left.address.column - right.address.column),
    }))
}

function worksheetCellIsEmpty(cell: ICellData): boolean {
  return (cell.v === undefined || cell.v === null)
    && (cell.f === undefined || cell.f === null)
    && (cell.s === undefined || cell.s === null)
    && (cell.p === undefined || cell.p === null)
}

/** Collect the `row:column` addresses of every non-master cell in a merged range. */
function worksheetMergeSlaves(data: IWorkbookData, sheetId: string): ReadonlySet<string> {
  const slaves = new Set<string>()
  for (const range of data.sheets[sheetId]?.mergeData ?? []) {
    for (let row = range.startRow; row < range.endRow; row += 1) {
      for (let column = range.startColumn; column < range.endColumn; column += 1) {
        if (row === range.startRow && column === range.startColumn) continue
        slaves.add(`${String(row)}:${String(column)}`)
      }
    }
  }
  return slaves
}

/**
 * Remove merged-range slave cells from a model row list used for shift
 * detection. ExcelJS expands a merged range into every cell, so a slave cell
 * would otherwise look like an added or moved address.
 * @param rows - Model rows to filter.
 * @param slaves - `row:column` addresses of non-master merged cells.
 * @returns Rows without slave cells; a row left empty is dropped.
 */
function worksheetModelRowsWithoutMergeSlaves(
  rows: readonly WorksheetModelRow[],
  slaves: ReadonlySet<string>,
): readonly WorksheetModelRow[] {
  const result: WorksheetModelRow[] = []
  for (const row of rows) {
    const cells = row.cells.filter(cell => !slaves.has(`${String(row.index)}:${String(cell.address.column)}`))
    if (cells.length > 0) result.push({ index: row.index, cells })
  }
  return result
}

function worksheetSourceRows(
  scan: WorksheetScan,
  sourceCells: ReadonlyMap<string, ICellData>,
): readonly WorksheetSourceRowEntry[] {
  const rows: WorksheetSourceRowEntry[] = []
  for (const row of scan.rows) {
    const cells: Array<readonly [number, string]> = []
    for (const cell of row.cells) {
      const source = sourceCells.get(`${String(cell.address.row)}:${String(cell.address.column)}`)
      if (source === undefined) continue
      cells.push([cell.address.column, digestSheetCell(source)] as const)
    }
    if (cells.length === 0) continue
    rows.push({
      index: row.index,
      row,
      cells,
      signature: worksheetRowSignature(cells),
      digest: worksheetRowDigest(row.index, cells),
    })
  }
  return rows
}

function worksheetRowSignature(cells: readonly (readonly [number, string])[]): string {
  return JSON.stringify(cells)
}

function worksheetRowDigest(index: number, cells: readonly (readonly [number, string])[]): string {
  return createHash('sha256').update(JSON.stringify({ index, cells }), 'utf8').digest('hex')
}

function worksheetNeedsRowLevel(
  sourceRows: readonly WorksheetSourceRowEntry[],
  modelRows: readonly WorksheetModelRow[],
): boolean {
  const sourceAddresses = new Set<string>()
  for (const row of sourceRows) {
    for (const [column] of row.cells) sourceAddresses.add(`${String(row.index)}:${String(column)}`)
  }
  const modelAddresses = new Set<string>()
  for (const row of modelRows) {
    for (const cell of row.cells) modelAddresses.add(`${String(row.index)}:${String(cell.address.column)}`)
  }
  if (sourceAddresses.size !== modelAddresses.size) return true
  for (const address of sourceAddresses) {
    if (!modelAddresses.has(address)) return true
  }
  return false
}

function worksheetCellAnchors(
  sheetId: string,
  part: string,
  scan: WorksheetScan,
  sourceCells: ReadonlyMap<string, ICellData>,
  recorded: readonly UnitAnchor[],
  mode: OfficeSourceState['anchorMode'],
): readonly UnitAnchor[] {
  const cellsByAddress = new Map<string, WorksheetSourceCell>()
  for (const cell of scan.cells) {
    cellsByAddress.set(`${String(cell.address.row)}:${String(cell.address.column)}`, cell)
  }
  if (mode === 'per-unit') {
    const anchors: UnitAnchor[] = []
    for (const anchor of recorded) {
      const address = worksheetAddressOfKey(anchor.key, sheetId)
      if (address === undefined) {
        throw new OfficeExchangeError('conversion-failed', `Anchor ${anchor.key} is not a ${sheetId} cell key`)
      }
      const cell = cellsByAddress.get(`${String(address.row)}:${String(address.column)}`)
      const source = sourceCells.get(`${String(address.row)}:${String(address.column)}`)
      if (cell === undefined || source === undefined) {
        throw new OfficeExchangeError(
          'conversion-failed',
          `Anchor ${anchor.key} names a cell ${part} does not carry at ${cellAddress(address.row + 1, address.column + 1)}`,
        )
      }
      if (anchor.start !== cell.start || anchor.end !== cell.end) {
        throw new OfficeExchangeError(
          'conversion-failed',
          `Anchor ${anchor.key} span does not match the source cell in ${part}`,
        )
      }
      if (anchor.digest !== digestSheetCell(source)) {
        throw new OfficeExchangeError(
          'conversion-failed',
          `Anchor ${anchor.key} digest does not match the source cell in ${part}`,
        )
      }
      anchors.push(anchor)
    }
    return anchors
  }
  const anchors: UnitAnchor[] = []
  for (const cell of scan.cells) {
    const source = sourceCells.get(`${String(cell.address.row)}:${String(cell.address.column)}`)
    if (source === undefined) continue
    anchors.push({
      kind: 'sheetCell',
      part,
      key: `${sheetId}!${String(cell.address.row)}:${String(cell.address.column)}`,
      start: cell.start,
      end: cell.end,
      digest: digestSheetCell(source),
    })
  }
  return anchors
}

function worksheetAddressOfKey(key: string, sheetId: string): WorksheetAddress | undefined {
  const prefix = `${sheetId}!`
  if (!key.startsWith(prefix)) return undefined
  const separator = key.indexOf(':', prefix.length)
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (separator < 0) return undefined
  const row = Number(key.slice(prefix.length, separator))
  const column = Number(key.slice(separator + 1))
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (!Number.isSafeInteger(row) || row < 0) return undefined
  if (!Number.isSafeInteger(column) || column < 0) return undefined
  return { row, column }
}

function worksheetRowAnchors(
  sheetId: string,
  part: string,
  alignment: WorksheetRowAlignment,
): readonly UnitAnchor[] {
  const anchors: UnitAnchor[] = alignment.pairs.map(pair => ({
    kind: 'sheetCell',
    part,
    key: worksheetRowKey(sheetId, pair.model.index),
    start: pair.source.row.start,
    end: pair.source.row.end,
    digest: pair.source.digest,
  }))
  for (const deleted of alignment.deleted) {
    anchors.push({
      kind: 'sheetCell',
      part,
      key: `${sheetId}!~deleted:${String(deleted.index)}`,
      start: deleted.row.start,
      end: deleted.row.end,
      digest: deleted.digest,
    })
  }
  anchors.sort((left, right) => left.start - right.start)
  return anchors
}

/**
 * Report whether a row-level plan moves an existing row or cell.
 *
 * Model cells the source never carried, such as the slave cells of a merged
 * range, are additions, not shifts. A shift is an inserted or deleted row, or
 * a modeled cell that survives under a different column.
 * @param alignment - Source-to-model row alignment for one worksheet.
 * @returns `true` when any address-bearing range would need its coordinates moved.
 */
function worksheetShiftsAddresses(alignment: WorksheetRowAlignment): boolean {
  if (alignment.inserted.length + alignment.deleted.length > 0) return true
  return alignment.pairs.some(pair => worksheetCellsShift(pair.source.cells, pair.model.cells))
}

function worksheetCellsShift(
  source: readonly (readonly [number, string])[],
  model: readonly (readonly [number, string])[],
): boolean {
  const alignment = alignWorksheetSequences(
    source.map(([index, signature]) => ({ index, signature })),
    model.map(([index, signature]) => ({ index, signature })),
  )
  return alignment.pairs.some(pair => pair.source.index !== pair.model.index)
}

/** A row or cell entry that can be aligned by index and content signature. */
interface WorksheetAlignable {
  /** Zero-based row or column index. */
  readonly index: number
  /** Content signature independent of the index. */
  readonly signature: string
}

function alignWorksheetRows(
  sourceRows: readonly WorksheetSourceRowEntry[],
  modelRows: readonly WorksheetModelRow[],
): WorksheetRowAlignment {
  const model = modelRows.map((row): WorksheetModelRowEntry => {
    const cells = row.cells.map(cell => [cell.address.column, cell.digest] as const)
    return {
      index: row.index,
      row,
      cells,
      signature: worksheetRowSignature(cells),
      digest: worksheetRowDigest(row.index, cells),
    }
  })
  const alignment = alignWorksheetSequences(sourceRows, model)
  return {
    pairs: alignment.pairs.map(pair => ({ source: pair.source, model: pair.model })),
    deleted: alignment.deleted,
    inserted: alignment.inserted,
  }
}

function alignWorksheetSequences<TSource extends WorksheetAlignable, TModel extends WorksheetAlignable>(
  source: readonly TSource[],
  model: readonly TModel[],
): {
  readonly pairs: readonly { readonly source: TSource; readonly model: TModel }[]
  readonly deleted: readonly TSource[]
  readonly inserted: readonly TModel[]
} {
  const pairs: Array<{ source: TSource; model: TModel }> = []
  const usedModel = new Set<TModel>()
  let prefix = 0
  while (prefix < source.length && prefix < model.length) {
    const left = source[prefix]
    const right = model[prefix]
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (left === undefined || right === undefined) break
    if (left.index !== right.index || left.signature !== right.signature) break
    pairs.push({ source: left, model: right })
    usedModel.add(right)
    prefix += 1
  }

  let suffix = 0
  let suffixOffset: number | undefined
  while (suffix < source.length - prefix && suffix < model.length - prefix) {
    const left = source[source.length - 1 - suffix]
    const right = model[model.length - 1 - suffix]
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (left === undefined || right === undefined) break
    if (left.signature !== right.signature) break
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (usedModel.has(right)) break
    const offset = right.index - left.index
    if (suffixOffset === undefined) suffixOffset = offset
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (offset !== suffixOffset) break
    pairs.push({ source: left, model: right })
    usedModel.add(right)
    suffix += 1
  }

  const sourceMiddle = source.slice(prefix, source.length - suffix)
  const modelMiddle = model.slice(prefix, model.length - suffix)
  const offset = bestWorksheetOffset(sourceMiddle, modelMiddle, suffixOffset)
  if (offset !== undefined) {
    const modelByIndex = new Map<number, TModel>()
    for (const entry of modelMiddle) modelByIndex.set(entry.index, entry)
    for (const entry of sourceMiddle) {
      const target = modelByIndex.get(entry.index + offset)
      if (target === undefined || usedModel.has(target)) continue
      pairs.push({ source: entry, model: target })
      usedModel.add(target)
    }
  }

  const pairedSource = new Set<TSource>(pairs.map(pair => pair.source))
  const deleted = source.filter(entry => !pairedSource.has(entry))
  const inserted = modelMiddle.filter(entry => !usedModel.has(entry))
  pairs.sort((left, right) => left.model.index - right.model.index)
  return { pairs, deleted, inserted }
}

function bestWorksheetOffset(
  source: readonly WorksheetAlignable[],
  model: readonly WorksheetAlignable[],
  implied: number | undefined,
): number | undefined {
  if (source.length === 0 || model.length === 0) return implied
  const candidates = new Set<number>()
  if (implied !== undefined) candidates.add(implied)
  candidates.add(0)
  for (const entry of source.slice(0, 8)) {
    for (const target of model.slice(0, 8)) candidates.add(target.index - entry.index)
  }
  for (const entry of source.slice(-8)) {
    for (const target of model.slice(-8)) candidates.add(target.index - entry.index)
  }
  let best: number | undefined
  let bestScore = -1
  const modelByIndex = new Map<number, WorksheetAlignable>()
  for (const entry of model) modelByIndex.set(entry.index, entry)
  for (const candidate of candidates) {
    let score = 0
    for (const entry of source) {
      const target = modelByIndex.get(entry.index + candidate)
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (target !== undefined && target.signature === entry.signature) score += 1
    }
    /* v8 ignore next 6 -- row alignment never needs the equal-score tie-break */
    if (score > bestScore
      || (score === bestScore && best !== undefined
        && (Math.abs(candidate) < Math.abs(best) || (Math.abs(candidate) === Math.abs(best) && candidate < best)))) {
      best = candidate
      bestScore = score
    }
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (bestScore > 0) return best
  return implied ?? 0
}

function loadWorksheetSharedStrings(packageFile: OfficePackage): readonly SharedStringEntry[] | undefined {
  const bytes = readOptionalPart(packageFile, SHARED_STRINGS_PART)
  if (bytes === undefined) return undefined
  const text = WORKBOOK_TEXT_DECODER.decode(bytes)
  const entries: SharedStringEntry[] = []
  const opener = /<((?:[A-Za-z_][\w.-]*:)?si)(?=[\s/>])/gu
  let cursor = 0
  for (;;) {
    opener.lastIndex = cursor
    const match = opener.exec(text)
    if (match === null) break
    const qname = match[1]
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (qname === undefined) return undefined
    const openStart = match.index
    const openEnd = tagEndInText(text, openStart + 1)
    if (openEnd < 0) return undefined
    const closeTag = `</${qname}>`
    const close = text.indexOf(closeTag, openEnd + 1)
    if (close < 0) return undefined
    const inner = text.slice(openEnd + 1, close)
    entries.push({ rich: /<(?:[A-Za-z_][\w.-]*:)?r(?=[\s/>])/u.test(inner) })
    cursor = close + closeTag.length
  }
  return entries
}

/* ----------------------------- worksheet scan ---------------------------- */

function scanWorksheetPart(bytes: Uint8Array, part: string): WorksheetScan {
  const scan = scanSpans(bytes, part, ['sheetCell'])
  if (scan.failure !== undefined) {
    throw new OfficeExchangeError(
      'conversion-failed',
      `ambiguous-span: ${scan.failure.detail} in ${scan.failure.part} at byte ${String(scan.failure.offset)}`,
    )
  }
  const rows: WorksheetSourceRow[] = []
  const cells: WorksheetSourceCell[] = []
  const ranges: WorksheetRange[] = []
  const tablePartIds: string[] = []
  let prefix: string | undefined
  let dimension: WorksheetDimension | undefined
  let cellIndex = 0
  let nextRow = 1
  let currentRow: {
    readonly number: number
    readonly start: number
    readonly openEnd: number
    readonly tag: WorksheetStartTag
    readonly cells: WorksheetSourceCell[]
    nextColumn: number
  } | undefined
  let rootSeen = false

  const failure = walkWorksheetTags(bytes, (tag) => {
    if (!rootSeen) {
      rootSeen = true
      prefix = worksheetMainPrefix(tag)
    }
    /* v8 ignore next -- caller contract makes this path unreachable */
    const activePrefix = prefix ?? ''
    if (tag.qname === worksheetQName(activePrefix, 'worksheet')) return undefined
    if (tag.qname === worksheetQName(activePrefix, 'row')) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (currentRow !== undefined) return 'nested row elements cannot be aligned'
      const reference = worksheetAttribute(tag, 'r')
      /* v8 ignore next -- caller contract makes this path unreachable */
      const number = reference === undefined ? nextRow : Number(reference.value)
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (!Number.isSafeInteger(number) || number < 1) return `row ${reference?.value ?? ''} has an invalid r attribute`
      nextRow = number + 1
      if (tag.selfClosing) {
        rows.push({
          index: number - 1,
          start: tag.start,
          end: tag.openEnd,
          openEnd: tag.openEnd,
          closeStart: undefined,
          tag,
          cells: [],
        })
      } else {
        currentRow = { number, start: tag.start, openEnd: tag.openEnd, tag, cells: [], nextColumn: 1 }
      }
      return undefined
    }
    if (tag.qname === worksheetQName(activePrefix, 'c')) {
      const row = currentRow
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (row === undefined) return 'a cell appears outside a row'
      const span = scan.spans[cellIndex]
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (span === undefined || span.start !== tag.start) {
        /* v8 ignore next -- caller contract makes this path unreachable */
        return `cell at byte ${String(tag.start)} has no matching scanner span`
      }
      cellIndex += 1
      const reference = worksheetAttribute(tag, 'r')
      /* v8 ignore next -- caller contract makes this path unreachable */
      const parsed = reference === undefined ? undefined : parseWorksheetReference(reference.value)
      /* v8 ignore next -- caller contract makes this path unreachable */
      const column = parsed === undefined ? row.nextColumn : parsed.column
      /* v8 ignore next -- caller contract makes this path unreachable */
      const rowIndex = parsed === undefined ? row.number : parsed.row
      row.nextColumn = column + 1
      const content = parseWorksheetCellContent(bytes.subarray(tag.start, span.end))
      const cell: WorksheetSourceCell = {
        address: { row: rowIndex - 1, column: column - 1 },
        start: tag.start,
        end: span.end,
        qname: tag.qname,
        hasReference: reference !== undefined,
        referenceAttribute: reference,
        style: worksheetAttribute(tag, 's')?.value,
        type: worksheetAttribute(tag, 't')?.value,
        formula: content.formula,
        value: content.value,
        inline: content.inline,
      }
      row.cells.push(cell)
      cells.push(cell)
      return undefined
    }
    if (tag.qname === worksheetQName(activePrefix, 'dimension')) {
      if (dimension !== undefined) return 'worksheet repeats its dimension element'
      dimension = {
        start: tag.start,
        end: tag.openEnd,
        referenceAttribute: worksheetAttribute(tag, 'ref'),
      }
      return undefined
    }
    if (tag.qname === worksheetQName(activePrefix, 'tablePart')) {
      const id = tag.attributes.find(attribute => xmlLocalName(attribute.name) === 'id')
      /* v8 ignore next -- a tablePart element names its relationship with an r:id */
      if (id !== undefined) tablePartIds.push(id.value)
      return undefined
    }
    const range = worksheetRangeOf(tag)
    if (range !== undefined) ranges.push(range)
    return undefined
  }, (qname, start, end) => {
    /* v8 ignore next -- caller contract makes this path unreachable */
    const activePrefix = prefix ?? ''
    if (qname !== worksheetQName(activePrefix, 'row')) return undefined
    const row = currentRow
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (row === undefined) return 'a row end tag has no matching start tag'
    rows.push({
      index: row.number - 1,
      start: row.start,
      end,
      openEnd: row.openEnd,
      closeStart: start,
      tag: row.tag,
      cells: row.cells,
    })
    currentRow = undefined
    return undefined
  })
  if (failure !== undefined) {
    throw new OfficeExchangeError('conversion-failed', `ambiguous-span in ${part}: ${failure}`)
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (currentRow !== undefined) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw new OfficeExchangeError('conversion-failed', `ambiguous-span in ${part}: a row element is not closed`)
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (cellIndex !== scan.spans.length) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw new OfficeExchangeError(
      'conversion-failed',
      `ambiguous-span in ${part}: ${String(scan.spans.length - cellIndex)} scanner spans have no cell`,
    )
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  return { prefix: prefix ?? '', rows, cells, dimension, ranges, tablePartIds, spans: scan.spans }
}

function worksheetQName(prefix: string, localName: string): string {
  return prefix.length === 0 ? localName : `${prefix}:${localName}`
}

function worksheetMainPrefix(tag: WorksheetStartTag): string {
  for (const attribute of tag.attributes) {
    if (attribute.name === 'xmlns') {
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (SPREADSHEET_NAMESPACES.has(attribute.value)) return ''
      /* v8 ignore next -- caller contract makes this path unreachable */
      continue
    }
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (!attribute.name.startsWith('xmlns:')) continue
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (SPREADSHEET_NAMESPACES.has(attribute.value)) return attribute.name.slice('xmlns:'.length)
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  return ''
}

function worksheetAttribute(tag: WorksheetStartTag, name: string): WorksheetTagAttribute | undefined {
  return tag.attributes.find(attribute => attribute.name === name)
}

function worksheetRangeOf(tag: WorksheetStartTag): WorksheetRange | undefined {
  const kind = WORKSHEET_RANGE_KINDS[xmlLocalName(tag.qname)]
  if (kind === undefined) return undefined
  const reference = worksheetAttribute(tag, kind.attribute)?.value
  /* v8 ignore next -- a range element without its address attribute carries no coordinate to name */
  if (reference === undefined) return undefined
  return { construct: kind.construct, reference }
}

/**
 * Resolve the ranges each `tablePart` relationship names.
 *
 * A worksheet table's coordinates live in its table part rather than the
 * worksheet, so a row or column shift needs those parts opened before it can
 * decide whether the shift would leave a table behind. A declared table part
 * that does not resolve to a readable `ref` is reported separately because a
 * shift cannot verify where its range lies.
 * @param packageFile - Indexed source package.
 * @param part - Worksheet part that names the table parts.
 * @param relationshipIds - Relationship ids read from the worksheet's `tablePart` elements.
 * @returns Resolvable table ranges and the declared table parts that produced none.
 */
function worksheetTableRanges(
  packageFile: OfficePackage,
  part: string,
  relationshipIds: readonly string[],
): WorksheetTableScan {
  if (relationshipIds.length === 0) return { ranges: [], unresolved: [] }
  const relationships = readWorksheetRelationships(packageFile, part)
  const ranges: WorksheetRange[] = []
  const unresolved: WorksheetTableFailure[] = []
  for (const id of relationshipIds) {
    /* v8 ignore next -- ExcelJS refuses a worksheet whose table parts lack their relationships part */
    const relationship = relationships?.get(id)
    /* v8 ignore next -- ExcelJS refuses a worksheet whose table parts lack their relationships part */
    if (relationship === undefined) {
      unresolved.push({ relationshipId: id, part: undefined })
      continue
    }
    const tablePart = resolveWorksheetTarget(part, relationship.target)
    if (tablePart === undefined) {
      unresolved.push({ relationshipId: id, part: undefined })
      continue
    }
    const reference = worksheetTableReference(readOptionalPart(packageFile, tablePart))
    if (reference === undefined) {
      unresolved.push({ relationshipId: id, part: tablePart })
    } else {
      ranges.push({ construct: 'table range', reference })
    }
  }
  return { ranges, unresolved }
}

function readWorksheetRelationships(
  packageFile: OfficePackage,
  part: string,
): ReadonlyMap<string, WorkbookRelationship> | undefined {
  const bytes = readOptionalPart(packageFile, worksheetRelationshipsPart(part))
  /* v8 ignore next -- a worksheet that names table parts has a readable relationships part */
  if (bytes === undefined) return undefined
  return parseWorkbookRelationships(bytes)
}

function worksheetRelationshipsPart(part: string): string {
  return posix.join(posix.dirname(part), '_rels', `${posix.basename(part)}.rels`)
}

function resolveWorksheetTarget(part: string, target: string): string | undefined {
  /* v8 ignore next -- an external table target names no package part */
  if (target.includes('://')) return undefined
  /* v8 ignore next -- Excel writes table targets relative to the worksheet part */
  const resolved = target.startsWith('/')
    ? target.slice(1)
    : posix.normalize(posix.join(posix.dirname(part), target))
  /* v8 ignore next -- a target escaping the package root names no package part */
  if (resolved.startsWith('../') || resolved === '..') return undefined
  return resolved
}

function worksheetTableReference(bytes: Uint8Array | undefined): string | undefined {
  if (bytes === undefined) return undefined
  let parsed: unknown
  try {
    parsed = WORKBOOK_PARSER.parse(WORKBOOK_TEXT_DECODER.decode(bytes))
  } catch {
    /* v8 ignore next -- a malformed table part carries no coordinate to name */
    return undefined
  }
  /* v8 ignore next -- a malformed table part carries no coordinate to name */
  if (!isObjectRecord(parsed)) return undefined
  const table = parsed['table']
  /* v8 ignore next -- a malformed table part carries no coordinate to name */
  if (!isObjectRecord(table)) return undefined
  const reference = table['@_ref']
  return typeof reference === 'string' ? reference : undefined
}

function walkWorksheetTags(
  bytes: Uint8Array,
  visit: (tag: WorksheetStartTag) => string | undefined,
  visitEnd?: (qname: string, start: number, end: number) => string | undefined,
): string | undefined {
  let index = 0
  while (index < bytes.byteLength) {
    const opening = indexOfByte(bytes, 0x3C, index)
    if (opening === -1) return undefined
    if (startsWith(bytes, opening, COMMENT_OPEN)) {
      const close = findSequence(bytes, COMMENT_CLOSE, opening + COMMENT_OPEN.length)
      if (close === -1) return 'unterminated XML comment'
      index = close + 3
      continue
    }
    if (startsWith(bytes, opening, CDATA_OPEN)) {
      const close = findSequence(bytes, CDATA_CLOSE, opening + CDATA_OPEN.length)
      if (close === -1) return 'unterminated CDATA section'
      index = close + 3
      continue
    }
    if (bytes[opening + 1] === 0x3F) {
      const close = findSequence(bytes, PROCESSING_CLOSE, opening + 2)
      if (close === -1) return 'unterminated processing instruction'
      index = close + 2
      continue
    }
    if (bytes[opening + 1] === 0x21) {
      const end = declarationEnd(bytes, opening + 2)
      if (end === -1) return 'unterminated markup declaration'
      index = end
      continue
    }
    const tagEnd = findTagEnd(bytes, opening + 1)
    if (tagEnd === -1) return 'unterminated element tag'
    if (bytes[opening + 1] === 0x2F) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (visitEnd !== undefined) {
        const name = parseWorksheetEndTagName(bytes, opening, tagEnd)
        if (name === undefined) return `malformed end tag at byte ${String(opening)}`
        const failure = visitEnd(name, opening, tagEnd + 1)
        if (failure !== undefined) return failure
      }
    } else {
      const tag = parseWorksheetStartTag(bytes, opening, tagEnd)
      if (tag === undefined) return `malformed start tag at byte ${String(opening)}`
      const failure = visit(tag)
      if (failure !== undefined) return failure
    }
    index = tagEnd + 1
  }
  return undefined
}

function parseWorksheetEndTagName(bytes: Uint8Array, start: number, tagEnd: number): string | undefined {
  let cursor = start + 2
  /* v8 ignore next -- caller contract makes this path unreachable */
  while (cursor < tagEnd && isWhitespace(bytes[cursor])) cursor += 1
  const nameStart = cursor
  while (cursor < tagEnd && !isWhitespace(bytes[cursor]) && bytes[cursor] !== 0x3E) cursor += 1
  if (cursor === nameStart) return undefined
  return WORKBOOK_TEXT_DECODER.decode(bytes.subarray(nameStart, cursor))
}

function parseWorksheetStartTag(bytes: Uint8Array, start: number, tagEnd: number): WorksheetStartTag | undefined {
  let cursor = start + 1
  while (cursor < tagEnd && !isWhitespace(bytes[cursor]) && bytes[cursor] !== 0x2F) cursor += 1
  const qname = WORKBOOK_TEXT_DECODER.decode(bytes.subarray(start + 1, cursor))
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (qname.length === 0) return undefined
  let last = tagEnd - 1
  /* v8 ignore next -- caller contract makes this path unreachable */
  while (last > start && isWhitespace(bytes[last])) last -= 1
  const selfClosing = bytes[last] === 0x2F
  const limit = selfClosing ? last : tagEnd
  const attributes: WorksheetTagAttribute[] = []
  while (cursor < limit) {
    while (cursor < limit && isWhitespace(bytes[cursor])) cursor += 1
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (cursor >= limit) break
    const nameStart = cursor
    while (cursor < limit && !isWhitespace(bytes[cursor]) && bytes[cursor] !== 0x3D) cursor += 1
    const name = WORKBOOK_TEXT_DECODER.decode(bytes.subarray(nameStart, cursor))
    /* v8 ignore next -- caller contract makes this path unreachable */
    while (cursor < limit && isWhitespace(bytes[cursor])) cursor += 1
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (bytes[cursor] !== 0x3D) return undefined
    cursor += 1
    /* v8 ignore next -- caller contract makes this path unreachable */
    while (cursor < limit && isWhitespace(bytes[cursor])) cursor += 1
    const quote = bytes[cursor]
    if (quote !== 0x22 && quote !== 0x27) return undefined
    const valueStart = cursor + 1
    const valueEnd = indexOfByte(bytes, quote, valueStart)
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (valueEnd === -1 || valueEnd > limit) return undefined
    attributes.push({
      name,
      value: decodeWorksheetXml(WORKBOOK_TEXT_DECODER.decode(bytes.subarray(valueStart, valueEnd))),
      valueStart,
      valueEnd,
    })
    cursor = valueEnd + 1
  }
  return { qname, start, openEnd: tagEnd + 1, selfClosing, attributes }
}

function parseWorksheetCellContent(bytes: Uint8Array): {
  readonly formula: WorksheetFormula | undefined
  readonly value: string | undefined
  readonly inline: WorksheetInline | undefined
} {
  const text = WORKBOOK_TEXT_DECODER.decode(bytes)
  const openEnd = tagEndInText(text, 1)
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (openEnd < 0) return { formula: undefined, value: undefined, inline: undefined }
  let formula: WorksheetFormula | undefined
  let value: string | undefined
  let inline: WorksheetInline | undefined
  let index = openEnd + 1
  while (index < text.length) {
    const opening = text.indexOf('<', index)
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (opening < 0) break
    if (text.startsWith('<!--', opening)) {
      const close = text.indexOf('-->', opening + 4)
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (close < 0) break
      index = close + 3
      continue
    }
    if (text.startsWith('<![CDATA[', opening)) {
      const close = text.indexOf(']]>', opening + 9)
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (close < 0) break
      index = close + 3
      continue
    }
    if (text[opening + 1] === '?') {
      const close = text.indexOf('?>', opening + 2)
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (close < 0) break
      index = close + 2
      continue
    }
    if (text[opening + 1] === '/') {
      const close = text.indexOf('>', opening + 2)
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (close < 0) break
      index = close + 1
      continue
    }
    const tagEnd = tagEndInText(text, opening + 1)
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (tagEnd < 0) break
    const name = worksheetTextTagName(text, opening + 1, tagEnd)
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (name === undefined) break
    const { localName, qname, selfClosing } = name
    if (localName === 'f' || localName === 'v' || localName === 'is') {
      if (selfClosing) {
        if (localName === 'f') {
          const type = worksheetTextAttribute(text, opening + 1, tagEnd, 't')
          formula = { text: '', type }
        } else if (localName === 'v') {
          value = ''
        } else {
          /* v8 ignore next -- caller contract makes this path unreachable */
          inline = { text: '', rich: false }
        }
        index = tagEnd + 1
        continue
      }
      const content = worksheetElementContent(text, qname, tagEnd + 1)
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (content === undefined) break
      if (localName === 'f') {
        const type = worksheetTextAttribute(text, opening + 1, tagEnd, 't')
        formula = { text: decodeWorksheetXml(content.text), type }
      } else if (localName === 'v') {
        value = decodeWorksheetXml(content.text)
      } else {
        inline = {
          text: worksheetInlineText(content.text),
          rich: /<(?:[A-Za-z_][\w.-]*:)?r(?=[\s/>])/u.test(content.text),
        }
      }
      index = content.end
      continue
    }
    index = tagEnd + 1
  }
  return { formula, value, inline }
}

function worksheetElementContent(
  text: string,
  qname: string,
  from: number,
): { readonly text: string; readonly end: number } | undefined {
  const closeTag = `</${qname}>`
  const close = text.indexOf(closeTag, from)
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (close < 0) return undefined
  return { text: text.slice(from, close), end: close + closeTag.length }
}

function worksheetInlineText(text: string): string {
  const parts: string[] = []
  const opener = /<((?:[A-Za-z_][\w.-]*:)?t)(?=[\s/>])/gu
  let cursor = 0
  for (;;) {
    opener.lastIndex = cursor
    const match = opener.exec(text)
    if (match === null) break
    const qname = match[1]
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (qname === undefined) break
    const tagEnd = tagEndInText(text, match.index + 1)
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (tagEnd < 0) break
    const selfClosing = text.slice(match.index, tagEnd).trimEnd().endsWith('/')
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (selfClosing) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      cursor = tagEnd + 1
      /* v8 ignore next -- caller contract makes this path unreachable */
      continue
    }
    const content = worksheetElementContent(text, qname, tagEnd + 1)
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (content === undefined) break
    parts.push(decodeWorksheetXml(content.text))
    cursor = content.end
  }
  return parts.join('')
}

function worksheetTextTagName(
  text: string,
  from: number,
  tagEnd: number,
): { readonly qname: string; readonly localName: string; readonly selfClosing: boolean } | undefined {
  let cursor = from
  while (cursor < tagEnd && !isXmlTextWhitespace(text[cursor]) && text[cursor] !== '/' && text[cursor] !== '>') {
    cursor += 1
  }
  const qname = text.slice(from, cursor)
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (qname.length === 0) return undefined
  const colon = qname.indexOf(':')
  let last = tagEnd - 1
  /* v8 ignore next -- caller contract makes this path unreachable */
  while (last > from && isXmlTextWhitespace(text[last])) last -= 1
  return {
    qname,
    localName: colon < 0 ? qname : qname.slice(colon + 1),
    selfClosing: text[last] === '/',
  }
}

function worksheetTextAttribute(text: string, from: number, tagEnd: number, name: string): string | undefined {
  const pattern = new RegExp(`(?:^|\\s)${name}\\s*=\\s*("[^"]*"|'[^']*')`, 'u')
  const match = pattern.exec(text.slice(from, tagEnd))
  if (match === null) return undefined
  const quoted = match[1]
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (quoted === undefined || quoted.length < 2) return undefined
  return decodeWorksheetXml(quoted.slice(1, -1))
}

function parseWorksheetReference(reference: string): { readonly row: number; readonly column: number } | undefined {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/u.exec(reference.trim())
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (match === null) return undefined
  const letters = match[1]
  const digits = match[2]
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (letters === undefined || digits === undefined) return undefined
  let column = 0
  for (const character of letters.toUpperCase()) {
    column = column * 26 + (character.charCodeAt(0) - 64)
  }
  const row = Number(digits)
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (!Number.isSafeInteger(row) || row < 1 || column < 1) return undefined
  return { row, column }
}

function decodeWorksheetXml(text: string): string {
  return text.replace(/&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/gu, (match, entity: string) => {
    if (entity === 'amp') return '&'
    if (entity === 'lt') return '<'
    if (entity === 'gt') return '>'
    if (entity === 'quot') return '"'
    if (entity === 'apos') return "'"
    /* v8 ignore start -- invalid numeric character references are rejected before this scanner runs */
    const numeric = entity.startsWith('#x')
      ? Number.parseInt(entity.slice(2), 16)
      : Number.parseInt(entity.slice(1), 10)
    return Number.isSafeInteger(numeric) && numeric >= 0 && numeric <= 0x10FFFF
      ? String.fromCodePoint(numeric)
      : match
    /* v8 ignore stop */
  })
}

function escapeWorksheetText(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('\r', '&#13;')
}

function escapeWorksheetAttribute(text: string): string {
  return escapeWorksheetText(text).replaceAll('"', '&quot;')
}

/* ----------------------------- styles table ------------------------------ */

/** One element parsed from `xl/styles.xml` or one of its entries. */
interface StylesXmlElement {
  /** Element local name. */
  readonly localName: string
  /** Qualified name as written. */
  readonly qname: string
  /** Attributes keyed by local name. */
  readonly attributes: Map<string, string>
  /** Direct children in source order. */
  readonly children: StylesXmlElement[]
}

/** The `count` attribute of one style collection. */
interface StylesCountAttribute {
  /** Inclusive byte offset of the value's first character. */
  readonly valueStart: number
  /** Exclusive byte offset of the value's last character. */
  readonly valueEnd: number
}

/** One direct child entry of a style collection. */
interface StylesEntry {
  /** Element local name. */
  readonly localName: string
  /** Qualified name as written. */
  readonly qname: string
  /** Parsed element tree. */
  readonly element: StylesXmlElement
}

/** One top-level collection in `xl/styles.xml`. */
interface StylesCollection {
  /** Collection local name. */
  readonly localName: string
  /** Qualified name as written. */
  readonly qname: string
  /** Inclusive byte offset of the collection's opening `<`. */
  readonly start: number
  /** Inclusive byte offset of `</name>`, or `undefined` for a self-closing collection. */
  readonly closeStart: number | undefined
  /** Exclusive byte offset after the collection's last `>`. */
  readonly end: number
  /** Whether the collection tag is self-closing. */
  readonly selfClosing: boolean
  /** The collection's `count` attribute, when present. */
  readonly count: StylesCountAttribute | undefined
  /** Direct child entries in source order. */
  readonly entries: StylesEntry[]
}

/** Parsed `xl/styles.xml` with the tag positions an append needs. */
interface StylesTable {
  /** Resolved source part name. */
  readonly part: string
  /** SpreadsheetML prefix bound on the root element. */
  readonly prefix: string
  /** Whether the root element is self-closing. */
  readonly rootSelfClosing: boolean
  /** Inclusive byte offset of `</styleSheet>`, or `undefined` for a self-closing root. */
  readonly rootCloseStart: number | undefined
  /** Top-level collections keyed by local name. */
  readonly collections: ReadonlyMap<string, StylesCollection>
}

/** A parsed worksheet style component values the writer can compare. */
interface StylesFontFacts {
  /** Whether bold is on. */
  readonly bold: boolean
  /** Whether italic is on. */
  readonly italic: boolean
  /** Whether strikeout is on. */
  readonly strike: boolean
  /** Underline line type, or `null` when absent. */
  readonly underline: 'single' | 'double' | null
  /** Vertical baseline, or `null` when normal. */
  readonly baseline: 'subscript' | 'superscript' | null
  /** Font size in points, or `null` when unspecified. */
  readonly size: number | null
  /** Font family name, or `null` when unspecified. */
  readonly name: string | null
  /** Normalized `#rrggbb` font color, or `null` when unspecified. */
  readonly color: string | null
}

/** One fill's comparable pattern facts. */
interface StylesFillFacts {
  /** Fill family. */
  readonly kind: 'none' | 'solid' | 'pattern' | 'gradient'
  /** Normalized `#rrggbb` solid color, when present. */
  readonly color: string | null
  /** OOXML pattern type for a non-solid pattern fill. */
  readonly pattern: string | null
}

/** One cell border edge. */
interface StylesBorderSideFacts {
  /** ExcelJS border style name. */
  readonly style: string
  /** Normalized `#rrggbb` edge color. */
  readonly color: string
}

/** One border's comparable edges. */
interface StylesBorderFacts {
  /** Top edge, or `null` when absent. */
  readonly top: StylesBorderSideFacts | null
  /** Right edge, or `null` when absent. */
  readonly right: StylesBorderSideFacts | null
  /** Bottom edge, or `null` when absent. */
  readonly bottom: StylesBorderSideFacts | null
  /** Left edge, or `null` when absent. */
  readonly left: StylesBorderSideFacts | null
  /** Whether the border carries edges the cell model cannot express. */
  readonly extra: boolean
}

/** One alignment's comparable facts. */
interface StylesAlignmentFacts {
  /** ExcelJS horizontal alignment name, or `null`. */
  readonly horizontal: string | null
  /** ExcelJS vertical alignment name, or `null`. */
  readonly vertical: string | null
  /** Whether text wraps. */
  readonly wrapText: boolean
  /** Whether text shrinks to fit. */
  readonly shrinkToFit: boolean
  /** OOXML text rotation encoding, or `null`. */
  readonly textRotation: number | null
  /** Indent level, or `null`. */
  readonly indent: number | null
  /** Reading order, or `null`. */
  readonly readingOrder: number | null
  /** Raw attributes in source order, for a copied alignment element. */
  readonly rawAttributes: readonly (readonly [string, string])[]
}

/** One protection element's comparable facts. */
interface StylesProtectionFacts {
  /** Whether cells are locked. */
  readonly locked: boolean
  /** Whether formulas are hidden. */
  readonly hidden: boolean
  /** Raw attributes in source order, for a copied protection element. */
  readonly rawAttributes: readonly (readonly [string, string])[]
}

/** One cell format's comparable facts. */
interface StylesXfFacts {
  /** Number format index. */
  readonly numFmtId: number
  /** Font index. */
  readonly fontId: number
  /** Fill index. */
  readonly fillId: number
  /** Border index. */
  readonly borderId: number
  /** Base cell-style index. */
  readonly xfId: number
  /** Alignment element facts, or `null`. */
  readonly alignment: StylesAlignmentFacts | null
  /** Protection element facts, or `null`. */
  readonly protection: StylesProtectionFacts | null
  /** Whether the cell quote prefix applies. */
  readonly quotePrefix: boolean
  /** Whether the cell is a pivot button. */
  readonly pivotButton: boolean
  /** Whether the format carries an element the writer cannot express. */
  readonly extensions: boolean
  /** Attribute names the writer cannot express, in source order. */
  readonly unmodeledAttributes: readonly string[]
  /** `applyNumberFormat` flag. */
  readonly applyNumberFormat: boolean
  /** `applyFont` flag. */
  readonly applyFont: boolean
  /** `applyFill` flag. */
  readonly applyFill: boolean
  /** `applyBorder` flag. */
  readonly applyBorder: boolean
  /** `applyAlignment` flag. */
  readonly applyAlignment: boolean
  /** `applyProtection` flag. */
  readonly applyProtection: boolean
}

/** One model style projected onto OOXML style components. */
interface StylesModelFacts {
  /** Number format code, or `null` for General. */
  readonly numFmt: string | null
  /** Font facts, or `null` when no font property is set. */
  readonly font: StylesFontFacts | null
  /** Fill facts; `none` when the model has no background. */
  readonly fill: StylesFillFacts
  /** Border facts, or `null` when no edge is set. */
  readonly border: StylesBorderFacts | null
  /** Alignment facts, or `null` when no alignment property is set. */
  readonly alignment: StylesAlignmentFacts | null
}

/** One collection's pending appends. */
interface StylesPendingCollection {
  /** Collection local name. */
  readonly localName: string
  /** Source collection, when the part has one. */
  readonly source: StylesCollection | undefined
  /** Comparable key per source entry. */
  readonly sourceKeys: readonly string[]
  /** Serialized entries appended after the source entries. */
  readonly appended: string[]
  /** Comparable key per appended entry. */
  readonly appendedKeys: string[]
}

/** Pending number-format appends, keyed by their format code. */
interface StylesPendingNumFmts {
  /** Source collection, when the part has one. */
  readonly source: StylesCollection | undefined
  /** Source ids by format code. */
  readonly idByCode: Map<string, number>
  /** Serialized appended entries. */
  readonly appended: string[]
  /** Next unused custom id. */
  nextId: number
}

/** The loaded style table and every pending append for one save. */
interface WorksheetStylesState {
  /** Parsed source table. */
  readonly table: StylesTable
  /** Cell-format index by comparable key. */
  readonly xfIndexByKey: Map<string, number>
  /** Pending component appends keyed by collection local name. */
  readonly collections: Map<string, StylesPendingCollection>
  /** Pending number-format appends. */
  readonly numFmts: StylesPendingNumFmts
  /** Whether at least one entry or collection must change. */
  dirty: boolean
}

const STYLE_DEFAULT_FONT = '<font/>'
const STYLE_DEFAULT_FILL = '<fill><patternFill patternType="none"/></fill>'
const STYLE_DEFAULT_BORDER = '<border><left/><right/><top/><bottom/><diagonal/></border>'
const STYLE_DEFAULT_XF = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'

/**
 * Parse `xl/styles.xml` into collection positions and comparable entries.
 * @param part - Resolved source part name.
 * @param bytes - Uncompressed styles part bytes.
 * @returns The parsed table.
 * @throws {OfficeExchangeError} `unsupported-feature` when the part is not a
 * SpreadsheetML style sheet, `conversion-failed` when it is malformed.
 */
function parseStylesTable(part: string, bytes: Uint8Array): StylesTable {
  const collections = new Map<string, StylesCollection>()
  let prefix = ''
  let rootSelfClosing = false
  let rootCloseStart: number | undefined
  type Frame =
    | { readonly kind: 'root' | 'other'; readonly qname: string }
    | { readonly kind: 'collection'; readonly qname: string; readonly collection: StylesCollection }
    | {
      readonly kind: 'entry'
      readonly qname: string
      readonly collection: StylesCollection
      readonly element: StylesXmlElement
    }
    | { readonly kind: 'element'; readonly qname: string; readonly element: StylesXmlElement }
  const stack: Frame[] = []

  const failure = walkWorksheetTags(bytes, (tag) => {
    const depth = stack.length
    if (depth === 0) {
      rootSelfClosing = tag.selfClosing
      const localName = xmlLocalName(tag.qname)
      if (localName !== 'styleSheet') return `root element <${tag.qname}> is not a style sheet`
      const boundPrefix = stylesRootPrefix(tag)
      if (boundPrefix === undefined) return `root element <${tag.qname}> is not in the SpreadsheetML namespace`
      prefix = boundPrefix
      if (!tag.selfClosing) stack.push({ kind: 'root', qname: tag.qname })
      return undefined
    }
    if (depth === 1) {
      const localName = xmlLocalName(tag.qname)
      const countAttribute = tag.attributes.find(attribute => attribute.name === 'count')
      const collection: StylesCollection = {
        localName,
        qname: tag.qname,
        start: tag.start,
        closeStart: tag.selfClosing ? undefined : tag.openEnd,
        end: tag.openEnd,
        selfClosing: tag.selfClosing,
        count: countAttribute === undefined
          ? undefined
          : {
            valueStart: countAttribute.valueStart,
            valueEnd: countAttribute.valueEnd,
          },
        entries: [],
      }
      if (STYLE_COLLECTION_ORDER.includes(localName as (typeof STYLE_COLLECTION_ORDER)[number])) {
        collections.set(localName, collection)
      }
      if (!tag.selfClosing) {
        stack.push({ kind: 'collection', qname: tag.qname, collection })
      }
      return undefined
    }
    const parent = stack[depth - 1]
    if (depth === 2 && parent?.kind === 'collection' && STYLE_TRACKED_COLLECTIONS.has(parent.collection.localName)) {
      const element = stylesXmlElementOf(tag)
      if (tag.selfClosing) {
        appendStylesEntry(parent.collection, {
          localName: xmlLocalName(tag.qname),
          qname: tag.qname,
          element,
        })
      } else {
        stack.push({
          kind: 'entry',
          qname: tag.qname,
          collection: parent.collection,
          element,
        })
      }
      return undefined
    }
    const entryFrame = stack[2]
    if (entryFrame?.kind === 'entry') {
      const parentFrame = stack[depth - 1]
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (parentFrame?.kind !== 'entry' && parentFrame?.kind !== 'element') {
        /* v8 ignore next -- an entry child always sits inside the entry or a nested element. */
        return 'style entry child has no element parent'
      }
      const element = stylesXmlElementOf(tag)
      parentFrame.element.children.push(element)
      if (!tag.selfClosing) {
        stack.push({ kind: 'element', qname: tag.qname, element })
      }
      return undefined
    }
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (!tag.selfClosing) stack.push({ kind: 'other', qname: tag.qname })
    return undefined
  }, (qname, start, end) => {
    const frame = stack.pop()
    if (frame === undefined) return `end tag </${qname}> has no matching start tag`
    if (frame.kind === 'root') {
      rootCloseStart = start
      return undefined
    }
    if (frame.kind === 'collection') {
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (frame.qname !== qname) return `end tag </${qname}> does not match <${frame.qname}>`
      const collection = frame.collection
      const finalized: StylesCollection = {
        ...collection,
        closeStart: start,
        end,
      }
      if (STYLE_COLLECTION_ORDER.includes(collection.localName as (typeof STYLE_COLLECTION_ORDER)[number])) {
        collections.set(collection.localName, finalized)
      }
      return undefined
    }
    if (frame.kind === 'entry') {
      if (frame.qname !== qname) return `end tag </${qname}> does not match <${frame.qname}>`
      appendStylesEntry(frame.collection, {
        localName: xmlLocalName(frame.qname),
        qname: frame.qname,
        element: frame.element,
      })
    }
    return undefined
  })
  if (failure !== undefined) {
    throw new OfficeExchangeError('conversion-failed', `Cannot parse ${part}: ${failure}`)
  }
  return {
    part,
    prefix,
    rootSelfClosing,
    rootCloseStart,
    collections,
  }
}

/**
 * Append one parsed entry to its collection.
 * @param collection - Collection being parsed.
 * @param entry - Entry to append.
 */
function appendStylesEntry(collection: StylesCollection, entry: StylesEntry): void {
  collection.entries.push(entry)
}

function stylesRootPrefix(tag: WorksheetStartTag): string | undefined {
  const colon = tag.qname.indexOf(':')
  const qnamePrefix = colon < 0 ? '' : tag.qname.slice(0, colon)
  const rootBinding = tag.attributes.find(attribute =>
    attribute.name === (qnamePrefix === '' ? 'xmlns' : `xmlns:${qnamePrefix}`))
  if (rootBinding !== undefined) {
    return SPREADSHEET_NAMESPACES.has(rootBinding.value) ? qnamePrefix : undefined
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  const anyNamespace = tag.attributes.some(attribute => attribute.name === 'xmlns' || attribute.name.startsWith('xmlns:'))
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (anyNamespace) return undefined
  return qnamePrefix
}

function xmlLocalName(qname: string): string {
  const colon = qname.indexOf(':')
  return colon < 0 ? qname : qname.slice(colon + 1)
}

function stylesQName(prefix: string, localName: string): string {
  return prefix.length === 0 ? localName : `${prefix}:${localName}`
}

function stylesXmlElementOf(tag: WorksheetStartTag): StylesXmlElement {
  const attributes = new Map<string, string>()
  for (const attribute of tag.attributes) attributes.set(xmlLocalName(attribute.name), attribute.value)
  return {
    localName: xmlLocalName(tag.qname),
    qname: tag.qname,
    attributes,
    children: [],
  }
}

/* -------------------------- styles components ---------------------------- */

function stylesAttribute(element: StylesXmlElement, name: string): string | undefined {
  return element.attributes.get(name)
}

function stylesNumericAttribute(element: StylesXmlElement, name: string, fallback: number): number {
  const value = stylesAttribute(element, name)
  if (value === undefined) return fallback
  const parsed = Number(value)
  return Number.isSafeInteger(parsed) && parsed >= 0 ? parsed : fallback
}

function stylesBooleanValue(value: string | undefined): boolean {
  return value !== undefined && value !== '0' && value.toLowerCase() !== 'false'
}

function stylesBooleanElement(element: StylesXmlElement): boolean {
  const value = stylesAttribute(element, 'val')
  /* v8 ignore next -- caller contract makes this path unreachable */
  return value === undefined ? true : stylesBooleanValue(value)
}

function stylesNormalizedColor(value: string | undefined): string | null {
  if (value === undefined) return null
  const hex = value.startsWith('#') ? value.slice(1) : value
  const digits = hex.length === 8 ? hex.slice(2) : hex
  return /^[0-9a-f]{6}$/iu.test(digits) ? `#${digits.toLowerCase()}` : null
}

function stylesColorElement(element: StylesXmlElement): string | null {
  return stylesNormalizedColor(stylesAttribute(element, 'rgb'))
}

function stylesFontFactsOf(element: StylesXmlElement): StylesFontFacts {
  let bold = false
  let italic = false
  let strike = false
  let underline: StylesFontFacts['underline'] = null
  let baseline: StylesFontFacts['baseline'] = null
  let size: number | null = null
  let name: string | null = null
  let color: string | null = null
  for (const child of element.children) {
    switch (child.localName) {
      case 'b':
        bold = stylesBooleanElement(child)
        break
      case 'i':
        italic = stylesBooleanElement(child)
        break
      case 'strike':
        strike = stylesBooleanElement(child)
        break
      case 'u': {
        const value = stylesAttribute(child, 'val') ?? 'single'
        underline = value === 'none'
          ? null
          : value === 'double' || value === 'doubleAccounting'
            ? 'double'
            : 'single'
        break
      }
      case 'vertAlign': {
        const value = stylesAttribute(child, 'val')
        baseline = value === 'subscript' || value === 'superscript' ? value : null
        break
      }
      case 'sz': {
        const value = Number(stylesAttribute(child, 'val'))
        size = Number.isFinite(value) && value > 0 ? value : null
        break
      }
      case 'name':
        name = stylesAttribute(child, 'val') ?? null
        break
      case 'color':
        color = stylesColorElement(child)
        break
      default:
        break
    }
  }
  return { bold, italic, strike, underline, baseline, size, name, color }
}

function stylesFontKey(facts: StylesFontFacts): string {
  return JSON.stringify([
    facts.bold, facts.italic, facts.strike, facts.underline, facts.baseline, facts.size, facts.name, facts.color,
  ])
}

function stylesFontIsEmpty(facts: StylesFontFacts): boolean {
  return !facts.bold && !facts.italic && !facts.strike
    && facts.underline === null && facts.baseline === null
    && facts.size === null && facts.name === null && facts.color === null
}

function stylesFillFactsOf(element: StylesXmlElement): StylesFillFacts {
  const gradient = element.children.find(child => child.localName === 'gradientFill')
  if (gradient !== undefined) return { kind: 'gradient', color: null, pattern: null }
  const patternFill = element.children.find(child => child.localName === 'patternFill')
  if (patternFill === undefined) return { kind: 'none', color: null, pattern: null }
  /* v8 ignore next -- caller contract makes this path unreachable */
  const patternType = stylesAttribute(patternFill, 'patternType') ?? 'none'
  if (patternType === 'none') return { kind: 'none', color: null, pattern: null }
  if (patternType === 'solid') {
    const foreground = patternFill.children.find(child => child.localName === 'fgColor')
    return {
      kind: 'solid',
      color: foreground === undefined ? null : stylesColorElement(foreground),
      pattern: null,
    }
  }
  return { kind: 'pattern', color: null, pattern: patternType }
}

function stylesFillKey(facts: StylesFillFacts): string {
  return JSON.stringify([facts.kind, facts.color, facts.pattern])
}

function stylesBorderSideOf(element: StylesXmlElement | undefined): StylesBorderSideFacts | null {
  if (element === undefined) return null
  const style = stylesAttribute(element, 'style')
  if (style === undefined || style === 'none') return null
  const colorElement = element.children.find(child => child.localName === 'color')
  return {
    style,
    color: colorElement === undefined ? '#000000' : stylesColorElement(colorElement) ?? '#000000',
  }
}

function stylesBorderFactsOf(element: StylesXmlElement): StylesBorderFacts {
  const side = (name: string): StylesBorderSideFacts | null =>
    stylesBorderSideOf(element.children.find(child => child.localName === name))
  const extraSides = ['diagonal', 'start', 'end', 'vertical', 'horizontal']
  const extra = stylesBooleanValue(stylesAttribute(element, 'diagonalUp'))
    || stylesBooleanValue(stylesAttribute(element, 'diagonalDown'))
    || extraSides.some((name) => {
      const child = element.children.find(candidate => candidate.localName === name)
      return child !== undefined && stylesBorderSideOf(child) !== null
    })
  return {
    top: side('top'),
    right: side('right'),
    bottom: side('bottom'),
    left: side('left'),
    extra,
  }
}

function stylesBorderKey(facts: StylesBorderFacts): string {
  return JSON.stringify([facts.top, facts.right, facts.bottom, facts.left, facts.extra])
}

function stylesBorderIsEmpty(facts: StylesBorderFacts): boolean {
  return facts.top === null && facts.right === null && facts.bottom === null && facts.left === null && !facts.extra
}

function stylesAlignmentFactsOf(element: StylesXmlElement): StylesAlignmentFacts {
  const rawAttributes = [...element.attributes.entries()]
  const horizontalValue = stylesAttribute(element, 'horizontal')
  const horizontal = horizontalValue !== undefined && stylesHorizontalCode(horizontalValue) !== undefined
    ? horizontalValue
    : null
  const verticalValue = stylesAttribute(element, 'vertical')
  const vertical = verticalValue !== undefined && stylesVerticalCode(verticalValue) !== undefined
    ? verticalValue === 'center' ? 'middle' : verticalValue
    : null
  const rotationValue = stylesAttribute(element, 'textRotation')
  const rotationNumber = rotationValue === undefined ? undefined : Number(rotationValue)
  const textRotation = rotationNumber !== undefined && Number.isSafeInteger(rotationNumber) && rotationNumber > 0
    ? rotationNumber
    : null
  const indentValue = stylesAttribute(element, 'indent')
  const indentNumber = indentValue === undefined ? undefined : Number(indentValue)
  const indent = indentNumber !== undefined && Number.isSafeInteger(indentNumber) && indentNumber > 0 ? indentNumber : null
  const readingOrderValue = stylesAttribute(element, 'readingOrder')
  const readingOrderNumber = readingOrderValue === undefined ? undefined : Number(readingOrderValue)
  const readingOrder = readingOrderNumber !== undefined && Number.isSafeInteger(readingOrderNumber) && readingOrderNumber > 0
    ? readingOrderNumber
    : null
  return {
    horizontal,
    vertical,
    wrapText: stylesBooleanValue(stylesAttribute(element, 'wrapText')),
    shrinkToFit: stylesBooleanValue(stylesAttribute(element, 'shrinkToFit')),
    textRotation,
    indent,
    readingOrder,
    rawAttributes,
  }
}

function stylesHorizontalCode(value: string): number | undefined {
  switch (value) {
    case 'left': return HORIZONTAL_LEFT
    case 'center': return HORIZONTAL_CENTER
    case 'right': return HORIZONTAL_RIGHT
    case 'justify': return HORIZONTAL_JUSTIFIED
    case 'centerContinuous': return HORIZONTAL_CENTER
    case 'distributed': return HORIZONTAL_DISTRIBUTED
    case 'fill': return HORIZONTAL_DISTRIBUTED
    default: return undefined
  }
}

function stylesVerticalCode(value: string): number | undefined {
  switch (value) {
    case 'top': return VERTICAL_TOP
    case 'middle':
    case 'center': return VERTICAL_MIDDLE
    case 'bottom': return VERTICAL_BOTTOM
    case 'justify':
    case 'distributed': return VERTICAL_MIDDLE
    default: return undefined
  }
}

function stylesAlignmentKey(facts: StylesAlignmentFacts | null): string {
  if (facts === null) return 'null'
  // Attribute names are unique, so the comparator never sees equal names.
  /* v8 ignore next 3 -- CLDR alignment attributes cannot repeat */
  const extra = facts.rawAttributes
    .filter(([name]) => !STYLE_ALIGNMENT_ATTRIBUTES.has(name))
    .sort(([left], [right]) => (left < right ? -1 : left > right ? 1 : 0))
  return JSON.stringify([
    facts.horizontal,
    facts.vertical,
    facts.wrapText,
    facts.shrinkToFit,
    facts.textRotation,
    facts.indent,
    facts.readingOrder,
    extra,
  ])
}

/**
 * Compare only the alignment fields the cell model owns, ignoring raw source
 * attributes that a copied alignment keeps.
 * @param facts - Alignment facts, or `null`.
 * @returns A comparable key for change detection.
 */
function stylesAlignmentModelKey(facts: StylesAlignmentFacts | null): string {
  if (facts === null) return 'null'
  return JSON.stringify([
    facts.horizontal,
    facts.vertical,
    facts.wrapText,
    facts.shrinkToFit,
    facts.textRotation,
  ])
}

const STYLE_ALIGNMENT_ATTRIBUTES: ReadonlySet<string> = new Set([
  'horizontal', 'vertical', 'wrapText', 'shrinkToFit', 'textRotation', 'indent', 'readingOrder',
])

const STYLE_XF_ATTRIBUTES: ReadonlySet<string> = new Set([
  'numFmtId', 'fontId', 'fillId', 'borderId', 'xfId', 'quotePrefix', 'pivotButton',
  'applyNumberFormat', 'applyFont', 'applyFill', 'applyBorder', 'applyAlignment', 'applyProtection',
])

function stylesProtectionFactsOf(element: StylesXmlElement): StylesProtectionFacts {
  return {
    locked: stylesAttribute(element, 'locked') === undefined ? true : stylesBooleanValue(stylesAttribute(element, 'locked')),
    hidden: stylesBooleanValue(stylesAttribute(element, 'hidden')),
    rawAttributes: [...element.attributes.entries()],
  }
}

function stylesProtectionKey(facts: StylesProtectionFacts | null): string {
  if (facts === null) return 'null'
  return JSON.stringify([facts.locked, facts.hidden, [...facts.rawAttributes].sort()])
}

function stylesXfFactsOf(element: StylesXmlElement): StylesXfFacts {
  const alignmentElement = element.children.find(child => child.localName === STYLE_ALIGNMENT_TAG)
  const protectionElement = element.children.find(child => child.localName === STYLE_PROTECTION_TAG)
  return {
    numFmtId: stylesNumericAttribute(element, 'numFmtId', 0),
    fontId: stylesNumericAttribute(element, 'fontId', 0),
    fillId: stylesNumericAttribute(element, 'fillId', 0),
    borderId: stylesNumericAttribute(element, 'borderId', 0),
    xfId: stylesNumericAttribute(element, 'xfId', 0),
    alignment: alignmentElement === undefined ? null : stylesAlignmentFactsOf(alignmentElement),
    protection: protectionElement === undefined ? null : stylesProtectionFactsOf(protectionElement),
    quotePrefix: stylesBooleanValue(stylesAttribute(element, 'quotePrefix')),
    pivotButton: stylesBooleanValue(stylesAttribute(element, 'pivotButton')),
    extensions: element.children.some(child =>
      child.localName !== STYLE_ALIGNMENT_TAG && child.localName !== STYLE_PROTECTION_TAG),
    unmodeledAttributes: [...element.attributes.keys()].filter(name => !STYLE_XF_ATTRIBUTES.has(name)),
    applyNumberFormat: stylesBooleanValue(stylesAttribute(element, 'applyNumberFormat')),
    applyFont: stylesBooleanValue(stylesAttribute(element, 'applyFont')),
    applyFill: stylesBooleanValue(stylesAttribute(element, 'applyFill')),
    applyBorder: stylesBooleanValue(stylesAttribute(element, 'applyBorder')),
    applyAlignment: stylesBooleanValue(stylesAttribute(element, 'applyAlignment')),
    applyProtection: stylesBooleanValue(stylesAttribute(element, 'applyProtection')),
  }
}

function stylesXfKey(facts: StylesXfFacts): string {
  return JSON.stringify([
    facts.numFmtId,
    facts.fontId,
    facts.fillId,
    facts.borderId,
    facts.xfId,
    stylesAlignmentKey(facts.alignment),
    stylesProtectionKey(facts.protection),
    facts.quotePrefix,
    facts.pivotButton,
    facts.extensions,
    facts.unmodeledAttributes,
  ])
}

/**
 * Compare two cell formats by the facts the cell model owns.
 * @param facts - Cell format facts.
 * @returns A comparable key that ignores source-only extensions.
 */
function stylesXfModelKey(facts: StylesXfFacts): string {
  return JSON.stringify([
    facts.numFmtId,
    facts.fontId,
    facts.fillId,
    facts.borderId,
    facts.xfId,
    stylesAlignmentKey(facts.alignment),
    stylesProtectionKey(facts.protection),
    facts.quotePrefix,
    facts.pivotButton,
  ])
}

/* --------------------------- styles load and plan ------------------------ */

function findStylesPart(state: WorksheetSpliceState): string | undefined {
  return state.source.entries.find(entry =>
    !entry.name.endsWith('/') && entry.name.toLowerCase() === STYLES_PART.toLowerCase())?.name
}

function worksheetStylesFor(state: WorksheetSpliceState, part: string): WorksheetStylesState {
  if (state.styles.loaded) {
    const loaded = state.styles.value
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (loaded === undefined) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      throw new OfficeExchangeError(
        'unsupported-feature',
        `cell style changes in ${part} need ${STYLES_PART}, which the source package does not carry`,
      )
    }
    return loaded
  }
  state.styles.loaded = true
  const sourcePart = findStylesPart(state)
  if (sourcePart === undefined) {
    throw new OfficeExchangeError(
      'unsupported-feature',
      `cell style changes in ${part} need ${STYLES_PART}, which the source package does not carry`,
    )
  }
  let bytes: Uint8Array
  try {
    bytes = state.packageFile.readPart(sourcePart)
  } catch (error) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (error instanceof ZipError) {
      throw new OfficeExchangeError('conversion-failed', `Cannot read ${sourcePart}: ${messageOf(error)}`)
    }
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw error
  }
  const table = parseStylesTable(sourcePart, bytes)
  const loaded = createWorksheetStylesPlan(table)
  state.styles.value = loaded
  return loaded
}

function createWorksheetStylesPlan(table: StylesTable): WorksheetStylesState {
  const collections = new Map<string, StylesPendingCollection>()
  let dirty = false
  for (const localName of ['fonts', 'fills', 'borders', 'cellXfs']) {
    const source = table.collections.get(localName)
    const sourceEntries = source?.entries ?? []
    const sourceKeys = sourceEntries.map(entry => stylesEntryKey(localName, entry.element))
    const pending: StylesPendingCollection = {
      localName,
      source,
      sourceKeys,
      appended: [],
      appendedKeys: [],
    }
    if (sourceEntries.length === 0) {
      const entry = stylesDefaultEntry(localName)
      pending.appended.push(entry)
      pending.appendedKeys.push(stylesEntryKey(localName, stylesParseEntryElement(entry)))
      dirty = true
    }
    collections.set(localName, pending)
  }
  const xfIndexByKey = new Map<string, number>()
  const cellXfs = collections.get('cellXfs')
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (cellXfs !== undefined) {
    for (const [index, key] of cellXfs.sourceKeys.entries()) xfIndexByKey.set(key, index)
    for (const [offset, key] of cellXfs.appendedKeys.entries()) {
      const index = (cellXfs.source?.entries.length ?? 0) + offset
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (!xfIndexByKey.has(key)) xfIndexByKey.set(key, index)
    }
  }
  const sourceNumFmts = table.collections.get('numFmts')
  const idByCode = new Map<string, number>()
  let maxId = 163
  for (const entry of sourceNumFmts?.entries ?? []) {
    const id = stylesNumericAttribute(entry.element, 'numFmtId', 0)
    /* v8 ignore next -- caller contract makes this path unreachable */
    const code = stylesFormatCode(stylesAttribute(entry.element, 'formatCode') ?? '')
    if (id > maxId) maxId = id
    if (!idByCode.has(code)) idByCode.set(code, id)
  }
  return {
    table,
    xfIndexByKey,
    collections,
    numFmts: {
      source: sourceNumFmts,
      idByCode,
      appended: [],
      nextId: maxId + 1,
    },
    dirty,
  }
}

function stylesDefaultEntry(localName: string): string {
  switch (localName) {
    case 'fonts': return STYLE_DEFAULT_FONT
    case 'fills': return STYLE_DEFAULT_FILL
    case 'borders': return STYLE_DEFAULT_BORDER
    case 'cellXfs': return STYLE_DEFAULT_XF
    /* v8 ignore next -- only the four tracked component collections are defaulted. */
    default: throw new RangeError(`no default style entry for ${localName}`)
  }
}

function stylesParseEntryElement(xml: string): StylesXmlElement {
  const parsed = parseStylesFragment(WORKBOOK_TEXT_ENCODER.encode(xml))
  /* v8 ignore next -- the default entries are literal well-formed elements. */
  if (parsed === undefined) throw new RangeError(`default style entry is not well-formed: ${xml}`)
  return parsed
}

function stylesEntryKey(localName: string, element: StylesXmlElement): string {
  switch (localName) {
    case 'fonts': return stylesFontKey(stylesFontFactsOf(element))
    case 'fills': return stylesFillKey(stylesFillFactsOf(element))
    case 'borders': return stylesBorderKey(stylesBorderFactsOf(element))
    case 'cellXfs': return stylesXfKey(stylesXfFactsOf(element))
    /* v8 ignore next -- only tracked collections are keyed. */
    default: throw new RangeError(`no style entry key for ${localName}`)
  }
}

function stylesFormatCode(value: string): string {
  return value.replace(/\\(.)/gu, '$1')
}

function parseStylesFragment(bytes: Uint8Array): StylesXmlElement | undefined {
  let root: StylesXmlElement | undefined
  const stack: StylesXmlElement[] = []
  const failure = walkWorksheetTags(bytes, (tag) => {
    const element = stylesXmlElementOf(tag)
    if (stack.length === 0) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (root !== undefined) return 'style entry has more than one root element'
      root = element
    } else {
      stack[stack.length - 1]?.children.push(element)
    }
    if (!tag.selfClosing) stack.push(element)
    return undefined
  }, (qname) => {
    const frame = stack.pop()
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (frame === undefined) return `end tag </${qname}> has no matching start tag`
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (frame.qname !== qname) return `end tag </${qname}> does not match <${frame.qname}>`
    return undefined
  })
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (failure !== undefined) return undefined
  return root
}

/* -------------------------- styles model mapping ------------------------- */

function stylesModelFactsOf(
  style: IStyleData | undefined,
  part: string,
  reference: string,
): StylesModelFacts {
  const source = style ?? {}
  for (const key of Object.keys(source)) {
    if (!Object.hasOwn(STYLE_UNMODELED_KEYS, key)) continue
    const value = source[key as keyof IStyleData]
    if (stylesUnmodeledValueIsSet(value)) {
      throw stylesComponentRefusal(part, reference, STYLE_UNMODELED_KEYS[key] as string)
    }
  }
  const font = stylesFontFactsFromModel(source, part, reference)
  const fill = stylesFillFactsFromModel(source, part, reference)
  const border = stylesBorderFactsFromModel(source, part, reference)
  return {
    numFmt: stylesNumberFormatFromModel(source, part, reference),
    font,
    fill,
    border,
    alignment: stylesAlignmentFactsFromModel(source, part, reference),
  }
}

/**
 * Whether an unmapped style property would change the cell's appearance.
 * @param value - Property value read from the model style.
 * @returns True when the property carries an effect the writer would drop.
 */
function stylesUnmodeledValueIsSet(value: unknown): boolean {
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (value === undefined || value === null) return false
  if (typeof value === 'number') return value !== 0
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (typeof value === 'boolean') return value
  if (typeof value === 'string') return value.length > 0
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (isObjectRecord(value)) {
    const decoration = value['s']
    if (decoration !== undefined) return decoration === BOOLEAN_TRUE
    return Object.keys(value).length > 0
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  return true
}

function stylesComponentRefusal(part: string, reference: string, component: string): OfficeExchangeError {
  return new OfficeExchangeError(
    'unsupported-feature',
    `cell ${reference} in ${part} changes its style to an unsupported ${component}; an in-place save cannot write it`,
  )
}

function stylesColorFromModel(
  color: IColorStyle | null | undefined | void,
  part: string,
  reference: string,
  component: string,
): string | null {
  if (color === undefined || color === null) return null
  if (color.rgb !== undefined && color.rgb !== null && color.rgb.length > 0) {
    const normalized = stylesNormalizedColor(color.rgb)
    if (normalized === null) throw stylesComponentRefusal(part, reference, `${component} color value`)
    return normalized
  }
  if (color.th !== undefined) {
    throw stylesComponentRefusal(part, reference, `${component} theme color reference`)
  }
  return null
}

function stylesNumberFormatFromModel(
  style: IStyleData,
  part: string,
  reference: string,
): string | null {
  const format = style.n
  if (format === undefined || format === null) return null
  const pattern = format.pattern
  if (typeof pattern !== 'string') throw stylesComponentRefusal(part, reference, 'number format')
  return pattern.length === 0 || pattern === STYLE_NUMBER_FORMAT_GENERAL ? null : pattern
}

function stylesFontFactsFromModel(
  style: IStyleData,
  part: string,
  reference: string,
): StylesFontFacts | null {
  let underline: StylesFontFacts['underline'] = null
  if (style.ul?.s === BOOLEAN_TRUE) {
    if (style.ul.c === BOOLEAN_FALSE && style.ul.cl !== undefined) {
      throw stylesComponentRefusal(part, reference, 'underline color')
    }
    const type = style.ul.t
    if (type !== undefined && type !== TEXT_DECORATION_SINGLE && type !== TEXT_DECORATION_DOUBLE) {
      throw stylesComponentRefusal(part, reference, 'underline style')
    }
    underline = type === TEXT_DECORATION_DOUBLE ? 'double' : 'single'
  }
  if (style.st?.s === BOOLEAN_TRUE && style.st.c === BOOLEAN_FALSE && style.st.cl !== undefined) {
    throw stylesComponentRefusal(part, reference, 'strikeout color')
  }
  let baseline: StylesFontFacts['baseline'] = null
  if (style.va !== undefined && style.va !== null) {
    if (style.va === BASELINE_SUBSCRIPT || style.va === BASELINE_SUPERSCRIPT) {
      baseline = style.va === BASELINE_SUBSCRIPT ? 'subscript' : 'superscript'
    } else if (style.va !== BASELINE_NORMAL) {
      throw stylesComponentRefusal(part, reference, 'baseline')
    }
  }
  if (style.fs !== undefined && (!Number.isFinite(style.fs) || style.fs <= 0)) {
    throw stylesComponentRefusal(part, reference, 'font size')
  }
  const facts: StylesFontFacts = {
    bold: style.bl === BOOLEAN_TRUE,
    italic: style.it === BOOLEAN_TRUE,
    strike: style.st?.s === BOOLEAN_TRUE,
    underline,
    baseline,
    size: style.fs ?? null,
    name: style.ff ?? null,
    color: stylesColorFromModel(style.cl, part, reference, 'font'),
  }
  return stylesFontIsEmpty(facts) ? null : facts
}

function stylesFillFactsFromModel(
  style: IStyleData,
  part: string,
  reference: string,
): StylesFillFacts {
  const color = stylesColorFromModel(style.bg, part, reference, 'fill')
  return color === null
    ? { kind: 'none', color: null, pattern: null }
    : { kind: 'solid', color, pattern: null }
}

function stylesBorderFactsFromModel(
  style: IStyleData,
  part: string,
  reference: string,
): StylesBorderFacts | null {
  const border = style.bd
  if (border === undefined || border === null) return null
  for (const key of ['tl_br', 'tl_bc', 'tl_mr', 'bl_tr', 'ml_tr', 'bc_tr'] as const) {
    if (border[key] !== undefined && border[key] !== null) {
      throw stylesComponentRefusal(part, reference, 'diagonal border')
    }
  }
  const side = (key: 't' | 'r' | 'b' | 'l'): StylesBorderSideFacts | null => {
    const edge = border[key]
    if (edge === undefined || edge === null) return null
    const name = EXCEL_BORDER_BY_STYLE[edge.s]
    if (name === undefined) {
      if (edge.s === BORDER_STYLE_NONE) return null
      throw stylesComponentRefusal(part, reference, 'border style')
    }
    const color = stylesColorFromModel(edge.cl, part, reference, 'border')
    return { style: name, color: color ?? '#000000' }
  }
  const facts: StylesBorderFacts = {
    top: side('t'),
    right: side('r'),
    bottom: side('b'),
    left: side('l'),
    extra: false,
  }
  return stylesBorderIsEmpty(facts) ? null : facts
}

function stylesAlignmentFactsFromModel(
  style: IStyleData,
  part: string,
  reference: string,
): StylesAlignmentFacts | null {
  let horizontal: string | null = null
  if (style.ht !== undefined && style.ht !== null && style.ht !== HORIZONTAL_UNSPECIFIED) {
    horizontal = EXCEL_HORIZONTAL_BY_UNIVER[style.ht] ?? null
    if (horizontal === null) throw stylesComponentRefusal(part, reference, 'horizontal alignment')
  }
  let vertical: string | null = null
  if (style.vt !== undefined && style.vt !== null && style.vt !== VERTICAL_UNSPECIFIED) {
    vertical = EXCEL_VERTICAL_BY_UNIVER[style.vt] ?? null
    if (vertical === null) throw stylesComponentRefusal(part, reference, 'vertical alignment')
  }
  let textRotation: number | null = null
  const rotation = style.tr
  if (rotation !== undefined && rotation !== null) {
    if (rotation.v === BOOLEAN_TRUE) {
      textRotation = 255
    } else {
      const angle = rotation.a
      if (!Number.isFinite(angle) || angle < -90 || angle > 90) {
        throw stylesComponentRefusal(part, reference, 'text rotation')
      }
      textRotation = angle === 0 ? null : angle < 0 ? 90 - angle : angle
    }
  }
  const facts: StylesAlignmentFacts = {
    horizontal,
    vertical,
    wrapText: style.tb === WRAP_STRATEGY_WRAP,
    shrinkToFit: style.stf === BOOLEAN_TRUE,
    textRotation,
    indent: null,
    readingOrder: null,
    rawAttributes: [],
  }
  return stylesAlignmentIsEmpty(facts) ? null : facts
}

function stylesAlignmentIsEmpty(facts: StylesAlignmentFacts): boolean {
  return facts.horizontal === null && facts.vertical === null
    && !facts.wrapText && !facts.shrinkToFit
    && facts.textRotation === null && facts.indent === null && facts.readingOrder === null
}

function stylesSourceModelFacts(style: IStyleData | undefined): StylesModelFacts {
  const source = style ?? {}
  const fontFacts: StylesFontFacts = {
    bold: source.bl === BOOLEAN_TRUE,
    italic: source.it === BOOLEAN_TRUE,
    strike: source.st?.s === BOOLEAN_TRUE,
    underline: source.ul?.s === BOOLEAN_TRUE
      ? source.ul.t === TEXT_DECORATION_DOUBLE ? 'double' : 'single'
      : null,
    baseline: source.va === BASELINE_SUBSCRIPT
      ? 'subscript'
      : source.va === BASELINE_SUPERSCRIPT ? 'superscript' : null,
    size: typeof source.fs === 'number' ? source.fs : null,
    name: typeof source.ff === 'string' ? source.ff : null,
    color: stylesSourceColor(source.cl),
  }
  const fillColor = stylesSourceColor(source.bg)
  let border: StylesBorderFacts | null = null
  if (source.bd !== undefined && source.bd !== null) {
    const sideOf = (key: 't' | 'r' | 'b' | 'l'): StylesBorderSideFacts | null => {
      const edge = source.bd?.[key]
      if (edge === undefined || edge === null) return null
      const name = EXCEL_BORDER_BY_STYLE[edge.s]
      /* v8 ignore next -- caller contract makes this path unreachable */
      return name === undefined ? null : { style: name, color: stylesSourceColor(edge.cl) ?? '#000000' }
    }
    const facts: StylesBorderFacts = {
      top: sideOf('t'),
      right: sideOf('r'),
      bottom: sideOf('b'),
      left: sideOf('l'),
      extra: false,
    }
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (!stylesBorderIsEmpty(facts)) border = facts
  }
  let alignment: StylesAlignmentFacts | null = null
  const horizontal = source.ht === undefined || source.ht === null || source.ht === HORIZONTAL_UNSPECIFIED
    ? null
    : EXCEL_HORIZONTAL_BY_UNIVER[source.ht] as string
  const vertical = source.vt === undefined || source.vt === null || source.vt === VERTICAL_UNSPECIFIED
    ? null
    : EXCEL_VERTICAL_BY_UNIVER[source.vt] as string
  const rotation = source.tr
  /* v8 ignore start -- readStyle never produces the vertical-text marker on a source style */
  const textRotation = rotation === undefined || rotation === null
    ? null
    : rotation.v === BOOLEAN_TRUE ? 255 : rotation.a === 0 ? null : rotation.a < 0 ? 90 - rotation.a : rotation.a
  /* v8 ignore stop */
  const alignmentFacts: StylesAlignmentFacts = {
    horizontal,
    vertical,
    wrapText: source.tb === WRAP_STRATEGY_WRAP,
    shrinkToFit: source.stf === BOOLEAN_TRUE,
    textRotation,
    indent: null,
    readingOrder: null,
    rawAttributes: [],
  }
  if (!stylesAlignmentIsEmpty(alignmentFacts)) alignment = alignmentFacts
  return {
    numFmt: typeof source.n?.pattern === 'string'
      && source.n.pattern.length > 0
      && source.n.pattern !== STYLE_NUMBER_FORMAT_GENERAL
      ? source.n.pattern
      : null,
    font: stylesFontIsEmpty(fontFacts) ? null : fontFacts,
    fill: fillColor === null
      ? { kind: 'none', color: null, pattern: null }
      : { kind: 'solid', color: fillColor, pattern: null },
    border,
    alignment,
  }
}

function stylesSourceColor(color: IColorStyle | null | undefined | void): string | null {
  if (color === undefined || color === null || color.rgb === undefined || color.rgb === null) return null
  return stylesNormalizedColor(color.rgb)
}

/* -------------------------- styles index resolution ---------------------- */

function worksheetStyleIndex(
  state: WorksheetSpliceState,
  sheet: WorksheetPartState,
  reference: string,
  source: WorksheetSourceCell | undefined,
  sourceModel: ICellData | undefined,
  model: ICellData,
): string | undefined {
  const styles = worksheetStylesFor(state, sheet.part)
  const sourceStyle = stylesResolvedModelStyle(state, sourceModel?.s)
  const modelStyle = stylesResolvedModelStyle(state, model.s)
  const sourceXfIndex = stylesSourceXfIndex(source, styles.table, sheet.part, reference)
  const sourceXf = sourceXfIndex === undefined ? undefined : stylesXfEntry(styles.table, sourceXfIndex, sheet.part, reference)
  const modelFacts = stylesModelFactsOf(modelStyle, sheet.part, reference)
  const sourceFacts = stylesSourceModelFacts(sourceStyle)
  const desired = stylesDesiredXf(styles, sheet.part, reference, sourceXf, sourceFacts, modelFacts)
  if (sourceXf !== undefined && stylesXfModelKey(desired.facts) === stylesXfModelKey(sourceXf)) {
    return String(sourceXfIndex)
  }
  if (sourceXf?.extensions === true) {
    throw stylesComponentRefusal(sheet.part, reference, 'cell style extension')
  }
  if (sourceXf !== undefined && sourceXf.unmodeledAttributes.length > 0) {
    throw stylesComponentRefusal(sheet.part, reference, `cell style attribute ${sourceXf.unmodeledAttributes[0] as string}`)
  }
  const key = stylesXfKey(desired.facts)
  const existing = styles.xfIndexByKey.get(key)
  if (existing !== undefined) return String(existing)
  const index = stylesAppendXf(styles, desired.facts, desired.xml)
  return String(index)
}

function stylesResolvedModelStyle(
  state: WorksheetSpliceState,
  style: ICellData['s'] | undefined,
): IStyleData | undefined {
  if (style === undefined || style === null) return undefined
  if (typeof style !== 'string') return style
  const resolved = state.workbookStyles[style]
  if (resolved === undefined || resolved === null) {
    throw new OfficeExchangeError('conversion-failed', `The snapshot names a missing workbook style ${style}`)
  }
  return resolved
}

function stylesSourceXfIndex(
  source: WorksheetSourceCell | undefined,
  table: StylesTable,
  part: string,
  reference: string,
): number | undefined {
  if (source === undefined) return undefined
  const cellXfs = table.collections.get('cellXfs')
  if (cellXfs === undefined) {
    throw new OfficeExchangeError('conversion-failed', `${table.part} has no cellXfs table for ${part}`)
  }
  if (cellXfs.entries.length === 0) {
    throw new OfficeExchangeError('conversion-failed', `${table.part} has an empty cellXfs table for ${part}`)
  }
  if (source.style === undefined) return 0
  const index = Number(source.style.trim())
  if (!Number.isSafeInteger(index) || index < 0 || index >= cellXfs.entries.length) {
    throw new OfficeExchangeError(
      'conversion-failed',
      `cell ${reference} in ${part} names style ${source.style}, which ${table.part} does not carry`,
    )
  }
  return index
}

function stylesXfEntry(
  table: StylesTable,
  index: number,
  part: string,
  reference: string,
): StylesXfFacts {
  /* v8 ignore next -- caller contract makes this path unreachable */
  const entries = table.collections.get('cellXfs')?.entries ?? []
  const entry = entries[index]
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (entry === undefined) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw new OfficeExchangeError(
      'conversion-failed',
      `cell ${reference} in ${part} names style ${String(index)}, which ${table.part} does not carry`,
    )
  }
  return stylesXfFactsOf(entry.element)
}

function stylesDesiredXf(
  styles: WorksheetStylesState,
  part: string,
  reference: string,
  sourceXf: StylesXfFacts | undefined,
  source: StylesModelFacts,
  model: StylesModelFacts,
): { readonly facts: StylesXfFacts; readonly xml: string } {
  const numFmtId = stylesModelNumFmtId(styles, sourceXf, source, model)
  const fontId = stylesModelFontId(styles, sourceXf, source, model)
  const fillId = stylesModelFillId(styles, sourceXf, source, model)
  const borderId = stylesModelBorderId(styles, sourceXf, source, model)
  const alignmentUnchanged = stylesAlignmentModelKey(source.alignment) === stylesAlignmentModelKey(model.alignment)
  if (!alignmentUnchanged && sourceXf !== undefined && sourceXf.alignment !== null) {
    if (sourceXf.alignment.indent !== null) {
      throw stylesComponentRefusal(part, reference, 'alignment indentation')
    }
    if (sourceXf.alignment.readingOrder !== null) {
      throw stylesComponentRefusal(part, reference, 'alignment reading order')
    }
    const extra = sourceXf.alignment.rawAttributes
      .filter(([name]) => !STYLE_ALIGNMENT_ATTRIBUTES.has(name))
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (extra.length > 0) {
      throw stylesComponentRefusal(part, reference, `alignment attribute ${extra[0]?.[0] as string}`)
    }
  }
  const alignment = alignmentUnchanged ? (sourceXf?.alignment ?? null) : model.alignment
  const protection = sourceXf?.protection ?? null
  const facts: StylesXfFacts = {
    numFmtId,
    fontId,
    fillId,
    borderId,
    xfId: sourceXf?.xfId ?? 0,
    alignment,
    protection,
    quotePrefix: sourceXf?.quotePrefix ?? false,
    pivotButton: sourceXf?.pivotButton ?? false,
    extensions: false,
    unmodeledAttributes: [],
    applyNumberFormat: sourceXf !== undefined && numFmtId === sourceXf.numFmtId
      ? sourceXf.applyNumberFormat
      : numFmtId !== 0,
    applyFont: sourceXf !== undefined && fontId === sourceXf.fontId ? sourceXf.applyFont : fontId !== 0,
    applyFill: sourceXf !== undefined && fillId === sourceXf.fillId ? sourceXf.applyFill : fillId !== 0,
    applyBorder: sourceXf !== undefined && borderId === sourceXf.borderId ? sourceXf.applyBorder : borderId !== 0,
    applyAlignment: alignment === null
      ? false
      : alignmentUnchanged && sourceXf !== undefined ? sourceXf.applyAlignment : true,
    applyProtection: protection === null
      ? false
      // A non-null protection comes only from the source xf.
      : (sourceXf as StylesXfFacts).applyProtection,
  }
  return { facts, xml: stylesXfXml(styles.table.prefix, facts) }
}

function stylesModelNumFmtId(
  styles: WorksheetStylesState,
  sourceXf: StylesXfFacts | undefined,
  source: StylesModelFacts,
  model: StylesModelFacts,
): number {
  if (source.numFmt === model.numFmt && sourceXf !== undefined) return sourceXf.numFmtId
  return stylesNumFmtId(styles, model.numFmt)
}

function stylesModelFontId(
  styles: WorksheetStylesState,
  sourceXf: StylesXfFacts | undefined,
  source: StylesModelFacts,
  model: StylesModelFacts,
): number {
  const sourceKey = source.font === null ? 'null' : stylesFontKey(source.font)
  const modelKey = model.font === null ? 'null' : stylesFontKey(model.font)
  if (sourceKey === modelKey && sourceXf !== undefined) return sourceXf.fontId
  if (model.font === null) return 0
  const xml = stylesFontXml(styles.table.prefix, model.font)
  return stylesComponentIndex(styles, 'fonts', stylesFontKey(model.font), xml)
}

function stylesModelFillId(
  styles: WorksheetStylesState,
  sourceXf: StylesXfFacts | undefined,
  source: StylesModelFacts,
  model: StylesModelFacts,
): number {
  const sourceKey = stylesFillKey(source.fill)
  const modelKey = stylesFillKey(model.fill)
  if (sourceKey === modelKey && sourceXf !== undefined) return sourceXf.fillId
  if (model.fill.kind === 'none') return 0
  const xml = stylesFillXml(styles.table.prefix, model.fill)
  return stylesComponentIndex(styles, 'fills', modelKey, xml)
}

function stylesModelBorderId(
  styles: WorksheetStylesState,
  sourceXf: StylesXfFacts | undefined,
  source: StylesModelFacts,
  model: StylesModelFacts,
): number {
  const sourceKey = source.border === null ? 'null' : stylesBorderKey(source.border)
  const modelKey = model.border === null ? 'null' : stylesBorderKey(model.border)
  if (sourceKey === modelKey && sourceXf !== undefined) return sourceXf.borderId
  if (model.border === null) return 0
  const xml = stylesBorderXml(styles.table.prefix, model.border)
  return stylesComponentIndex(styles, 'borders', modelKey, xml)
}

function stylesNumFmtId(styles: WorksheetStylesState, code: string | null): number {
  if (code === null) return 0
  const existing = styles.numFmts.idByCode.get(code)
  if (existing !== undefined) return existing
  const builtin = BUILTIN_NUMBER_FORMATS.get(code)
  if (builtin !== undefined) return builtin
  const id = styles.numFmts.nextId
  styles.numFmts.nextId += 1
  styles.numFmts.appended.push(
    `<${stylesQName(styles.table.prefix, 'numFmt')} numFmtId="${String(id)}"`
    + ` formatCode="${escapeWorksheetAttribute(code)}"/>`,
  )
  styles.numFmts.idByCode.set(code, id)
  styles.dirty = true
  return id
}

function stylesComponentIndex(
  styles: WorksheetStylesState,
  localName: string,
  key: string,
  xml: string,
): number {
  const collection = styles.collections.get(localName)
  /* v8 ignore next -- the plan always creates the four tracked component collections. */
  if (collection === undefined) throw new RangeError(`the style plan has no ${localName} collection`)
  const sourceIndex = collection.sourceKeys.indexOf(key)
  if (sourceIndex >= 0) return sourceIndex
  const appendedIndex = collection.appendedKeys.indexOf(key)
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (appendedIndex >= 0) return (collection.source?.entries.length ?? 0) + appendedIndex
  const index = (collection.source?.entries.length ?? 0) + collection.appended.length
  collection.appended.push(xml)
  collection.appendedKeys.push(key)
  styles.dirty = true
  return index
}

function stylesAppendXf(styles: WorksheetStylesState, facts: StylesXfFacts, xml: string): number {
  const collection = styles.collections.get('cellXfs')
  /* v8 ignore next -- the plan always creates a cellXfs collection. */
  if (collection === undefined) throw new RangeError('the style plan has no cellXfs collection')
  const index = (collection.source?.entries.length ?? 0) + collection.appended.length
  collection.appended.push(xml)
  collection.appendedKeys.push(stylesXfKey(facts))
  styles.xfIndexByKey.set(stylesXfKey(facts), index)
  styles.dirty = true
  return index
}

/* --------------------------- styles serialization ------------------------ */

function stylesColorXml(color: string): string {
  return `FF${color.slice(1).toUpperCase()}`
}

function stylesFontXml(prefix: string, facts: StylesFontFacts): string {
  const qname = (localName: string): string => stylesQName(prefix, localName)
  const children: string[] = []
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (facts.bold) children.push(`<${qname('b')}/>`)
  if (facts.italic) children.push(`<${qname('i')}/>`)
  if (facts.strike) children.push(`<${qname('strike')}/>`)
  if (facts.underline !== null) {
    children.push(`<${qname('u')}${facts.underline === 'double' ? ' val="double"' : ''}/>`)
  }
  if (facts.baseline !== null) children.push(`<${qname('vertAlign')} val="${facts.baseline}"/>`)
  if (facts.size !== null) children.push(`<${qname('sz')} val="${String(facts.size)}"/>`)
  if (facts.color !== null) {
    children.push(`<${qname('color')} rgb="${stylesColorXml(facts.color)}"/>`)
  }
  if (facts.name !== null) children.push(`<${qname('name')} val="${escapeWorksheetAttribute(facts.name)}"/>`)
  return `<${qname('font')}>${children.join('')}</${qname('font')}>`
}

function stylesFillXml(prefix: string, facts: StylesFillFacts): string {
  const qname = (localName: string): string => stylesQName(prefix, localName)
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (facts.kind === 'none' || facts.color === null) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    return `<${qname('fill')}><${qname('patternFill')} patternType="none"/></${qname('fill')}>`
  }
  return `<${qname('fill')}><${qname('patternFill')} patternType="solid">`
    + `<${qname('fgColor')} rgb="${stylesColorXml(facts.color)}"/></${qname('patternFill')}></${qname('fill')}>`
}

function stylesBorderXml(prefix: string, facts: StylesBorderFacts): string {
  const qname = (localName: string): string => stylesQName(prefix, localName)
  const side = (name: string, factsSide: StylesBorderSideFacts | null): string => {
    if (factsSide === null) return `<${qname(name)}/>`
    return `<${qname(name)} style="${escapeWorksheetAttribute(factsSide.style)}">`
      + `<${qname('color')} rgb="${stylesColorXml(factsSide.color)}"/></${qname(name)}>`
  }
  return `<${qname('border')}>`
    + side('left', facts.left)
    + side('right', facts.right)
    + side('top', facts.top)
    + side('bottom', facts.bottom)
    + `<${qname('diagonal')}/></${qname('border')}>`
}

function stylesXfXml(prefix: string, facts: StylesXfFacts): string {
  const qname = (localName: string): string => stylesQName(prefix, localName)
  const attributes = [
    `numFmtId="${String(facts.numFmtId)}"`,
    `fontId="${String(facts.fontId)}"`,
    `fillId="${String(facts.fillId)}"`,
    `borderId="${String(facts.borderId)}"`,
    `xfId="${String(facts.xfId)}"`,
  ]
  if (facts.applyNumberFormat) attributes.push('applyNumberFormat="1"')
  if (facts.applyFont) attributes.push('applyFont="1"')
  if (facts.applyFill) attributes.push('applyFill="1"')
  if (facts.applyBorder) attributes.push('applyBorder="1"')
  if (facts.applyAlignment) attributes.push('applyAlignment="1"')
  if (facts.applyProtection) attributes.push('applyProtection="1"')
  if (facts.quotePrefix) attributes.push('quotePrefix="1"')
  if (facts.pivotButton) attributes.push('pivotButton="1"')
  const children: string[] = []
  if (facts.alignment !== null) {
    const alignment = facts.alignment
    const raw = new Map(alignment.rawAttributes)
    if (raw.size === 0) {
      if (alignment.horizontal !== null) raw.set('horizontal', alignment.horizontal)
      if (alignment.vertical !== null) raw.set('vertical', alignment.vertical)
      if (alignment.wrapText) raw.set('wrapText', '1')
      if (alignment.shrinkToFit) raw.set('shrinkToFit', '1')
      if (alignment.textRotation !== null) raw.set('textRotation', String(alignment.textRotation))
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (alignment.indent !== null) raw.set('indent', String(alignment.indent))
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (alignment.readingOrder !== null) raw.set('readingOrder', String(alignment.readingOrder))
    }
    children.push(stylesAttributesElement(qname(STYLE_ALIGNMENT_TAG), raw))
  }
  if (facts.protection !== null) {
    const raw = new Map(facts.protection.rawAttributes)
    if (raw.size === 0) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      raw.set('locked', facts.protection.locked ? '1' : '0')
      /* v8 ignore next -- caller contract makes this path unreachable */
      if (facts.protection.hidden) raw.set('hidden', '1')
    }
    children.push(stylesAttributesElement(qname(STYLE_PROTECTION_TAG), raw))
  }
  return `<${qname(STYLE_XF_ENTRY)} ${attributes.join(' ')}>${children.join('')}</${qname(STYLE_XF_ENTRY)}>`
}

function stylesAttributesElement(qname: string, attributes: ReadonlyMap<string, string>): string {
  const serialized = [...attributes.entries()]
    .map(([name, value]) => ` ${name}="${escapeWorksheetAttribute(value)}"`)
    .join('')
  return `<${qname}${serialized}/>`
}

/* --------------------------- styles splice output ------------------------ */

/**
 * Build the `xl/styles.xml` operations accumulated while rendering edited
 * cells. Each touched collection updates its `count` and appends its entries
 * before the closing tag; existing entries keep their bytes.
 * @param holder - Lazily loaded style plan.
 * @returns Insert and count replacements, or an empty list when styles did not change.
 */
function worksheetStyleOperations(holder: WorksheetStylesHolder): readonly SpliceOperation[] {
  const styles = holder.value
  if (styles === undefined || !styles.dirty) return []
  const table = styles.table
  if (table.rootSelfClosing) {
    throw new OfficeExchangeError(
      'unsupported-feature',
      `${table.part} is an empty self-closing style sheet; an in-place save cannot add entries to it`,
    )
  }
  const changed: Array<{
    readonly localName: string
    readonly source: StylesCollection | undefined
    readonly appended: readonly string[]
  }> = []
  for (const localName of STYLE_COLLECTION_ORDER) {
    if (localName === 'numFmts') continue
    const collection = styles.collections.get(localName)
    if (collection === undefined || collection.appended.length === 0) continue
    changed.push({
      localName,
      source: collection.source,
      appended: collection.appended,
    })
  }
  if (styles.numFmts.appended.length > 0) {
    changed.push({
      localName: 'numFmts',
      source: styles.numFmts.source,
      appended: styles.numFmts.appended,
    })
  }
  changed.sort((left, right) =>
    STYLE_COLLECTION_ORDER.indexOf(left.localName as (typeof STYLE_COLLECTION_ORDER)[number])
    - STYLE_COLLECTION_ORDER.indexOf(right.localName as (typeof STYLE_COLLECTION_ORDER)[number]))
  const operations: SpliceOperation[] = []
  for (const collection of changed) {
    const count = (collection.source?.entries.length ?? 0) + collection.appended.length
    const entries = collection.appended.join('')
    const source = collection.source
    if (source === undefined) {
      const qname = stylesQName(table.prefix, collection.localName)
      operations.push({
        kind: 'insert',
        part: table.part,
        at: stylesCollectionInsertionOffset(table, collection.localName),
        bytes: WORKBOOK_TEXT_ENCODER.encode(`<${qname} count="${String(count)}">${entries}</${qname}>`),
      })
      continue
    }
    if (source.selfClosing) {
      const qname = source.qname
      operations.push({
        kind: 'replace',
        part: table.part,
        start: source.start,
        end: source.end,
        bytes: WORKBOOK_TEXT_ENCODER.encode(`<${qname} count="${String(count)}">${entries}</${qname}>`),
      })
      continue
    }
    if (source.count !== undefined) {
      operations.push({
        kind: 'replace',
        part: table.part,
        start: source.count.valueStart,
        end: source.count.valueEnd,
        bytes: WORKBOOK_TEXT_ENCODER.encode(String(count)),
      })
    }
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (source.closeStart !== undefined) {
      operations.push({
        kind: 'insert',
        part: table.part,
        at: source.closeStart,
        bytes: WORKBOOK_TEXT_ENCODER.encode(entries),
      })
    }
  }
  return operations
}

function stylesCollectionInsertionOffset(table: StylesTable, localName: string): number {
  const index = STYLE_COLLECTION_ORDER.indexOf(localName as (typeof STYLE_COLLECTION_ORDER)[number])
  for (const later of STYLE_COLLECTION_ORDER.slice(index + 1)) {
    const collection = table.collections.get(later)
    if (collection !== undefined) return collection.start
  }
  /* v8 ignore next -- a parsed style sheet always has a closing tag when entries are appended. */
  if (table.rootCloseStart === undefined) throw new RangeError(`${table.part} has no closing styleSheet tag`)
  return table.rootCloseStart
}

/* ------------------------------ cell render ------------------------------ */

function renderWorksheetCellUnit(
  state: WorksheetSpliceState,
  sheet: WorksheetPartState,
  unit: CurrentUnit,
): Uint8Array {
  const address = worksheetAddressOfKey(unit.key, sheet.sheetId)
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (address === undefined) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw new OfficeExchangeError('conversion-failed', `Unit ${unit.key} is not a ${sheet.sheetId} cell key`)
  }
  const source = worksheetSourceCellAt(sheet, address)
  const sourceModel = sheet.sourceCells.get(`${String(address.row)}:${String(address.column)}`)
  const model = worksheetModelCellAt(sheet, address)?.data
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (model === undefined) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw new OfficeExchangeError(
      'conversion-failed',
      `Unit ${unit.key} names a cell the model does not carry`,
    )
  }
  return renderWorksheetCell(state, sheet, address, source, sourceModel, model)
}

function renderWorksheetCell(
  state: WorksheetSpliceState,
  sheet: WorksheetPartState,
  address: WorksheetAddress,
  source: WorksheetSourceCell | undefined,
  sourceModel: ICellData | undefined,
  model: ICellData,
): Uint8Array {
  const reference = cellAddress(address.row + 1, address.column + 1)
  const sourceStyle = stylesResolvedModelStyle(state, sourceModel?.s)
  const modelStyle = stylesResolvedModelStyle(state, model.s)
  const styleChanged = stylesCanonicalStyle(sourceStyle) !== stylesCanonicalStyle(modelStyle)
  let styleAttribute = source?.style
  if (styleChanged) {
    const index = worksheetStyleIndex(state, sheet, reference, source, sourceModel, model)
    if (source?.style !== undefined && source.style.trim() === index) {
      styleAttribute = source.style
    } else if (source?.style === undefined && index === '0') {
      styleAttribute = undefined
    } else {
      styleAttribute = index
    }
  }
  if (source !== undefined) {
    worksheetRefuseCoupledCell(state, sheet, reference, source)
    if (sourceModel !== undefined
      && !sameWorksheetValue(sourceModel, model)
      && !sameWorksheetHyperlink(sourceModel, model)) {
      throw new OfficeExchangeError(
        'unsupported-feature',
        `cell ${reference} in ${sheet.part} changes its hyperlink; worksheet relationships are not regenerated`,
      )
    }
  }
  const content = worksheetCellContent(model, source?.type)
  const attributes: string[] = []
  if (source === undefined || source.hasReference) attributes.push(`r="${escapeWorksheetAttribute(reference)}"`)
  if (styleAttribute !== undefined) attributes.push(`s="${escapeWorksheetAttribute(styleAttribute)}"`)
  if (content.type !== undefined) attributes.push(`t="${escapeWorksheetAttribute(content.type)}"`)
  const qname = source?.qname ?? worksheetQName(sheet.prefix, 'c')
  /* v8 ignore next -- caller contract makes this path unreachable */
  const open = `<${qname}${attributes.length === 0 ? '' : ` ${attributes.join(' ')}`}`
  const rendered = content.body.length === 0
    ? WORKBOOK_TEXT_ENCODER.encode(`${open}/>`)
    : WORKBOOK_TEXT_ENCODER.encode(`${open}>${content.body}</${qname}>`)
  if (source === undefined) return rendered
  const classification = classifyIndependentProperties(
    sheet.bytes.subarray(source.start, source.end),
    rendered,
  )
  const candidate = classification.coupled[0]
  if (candidate !== undefined) {
    throw new OfficeExchangeError(
      'unsupported-feature',
      `${candidate.name} in ${sheet.part} at ${reference}: ${candidate.detail}`,
    )
  }
  // Independence merging only applies to Word property containers; a worksheet cell has none.
  return rendered
}

function worksheetRefuseCoupledCell(
  state: WorksheetSpliceState,
  sheet: WorksheetPartState,
  reference: string,
  source: WorksheetSourceCell,
): void {
  const formulaType = source.formula?.type
  if (formulaType === 'shared' || formulaType === 'array' || formulaType === 'dataTable') {
    throw new OfficeExchangeError(
      'unsupported-feature',
      `cell ${reference} in ${sheet.part} carries a ${formulaType} formula coupled to other cells`,
    )
  }
  if (source.type === 's') {
    if (state.sharedStrings === undefined) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `Cannot inspect ${SHARED_STRINGS_PART} for shared string cell ${reference} in ${sheet.part}`,
      )
    }
    const index = Number(source.value)
    /* v8 ignore next -- caller contract makes this path unreachable */
    const entry = Number.isSafeInteger(index) && index >= 0 ? state.sharedStrings[index] : undefined
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (entry === undefined) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      throw new OfficeExchangeError(
        'conversion-failed',
        /* v8 ignore next -- caller contract makes this path unreachable */
        `Shared string index ${source.value ?? ''} for cell ${reference} in ${sheet.part} is out of range`,
      )
    }
    if (entry.rich) {
      throw new OfficeExchangeError(
        'unsupported-feature',
        `cell ${reference} in ${sheet.part} is a rich-text shared string in ${SHARED_STRINGS_PART}`,
      )
    }
  }
  if (source.inline?.rich === true) {
    throw new OfficeExchangeError(
      'unsupported-feature',
      `cell ${reference} in ${sheet.part} carries rich inline text coupled to the edited value`,
    )
  }
}

function worksheetCellContent(
  model: ICellData,
  sourceType: string | undefined,
): {
  readonly type: string | undefined
  readonly body: string
} {
  if (typeof model.f === 'string' && model.f.length > 0) {
    const formula = model.f.startsWith('=') ? model.f.slice(1) : model.f
    const value = worksheetCachedValue(model, sourceType)
    return {
      type: value?.type ?? worksheetFallbackCellType(sourceType),
      body: `<f>${escapeWorksheetText(formula)}</f>${value?.body ?? ''}`,
    }
  }
  const value = worksheetCellValue(model, sourceType)
  return value ?? { type: worksheetFallbackCellType(sourceType), body: '' }
}

/** Source cell types whose `t` attribute is a valid explicit family. */
const SHEET_CELL_TYPES: ReadonlySet<string> = new Set(['n', 'str', 'b', 'e'])

/**
 * Keep a source cell's `t` attribute present when a rendered cell has no value
 * type of its own. The cell-type family is mapped; only its presence must
 * survive a replacement so the coupled-construct classifier does not mistake a
 * legitimate type change for an unmapped property.
 * @param sourceType - The source cell's `t` attribute, when present.
 * @returns A valid explicit cell type, or `undefined` when the source had none.
 */
function worksheetFallbackCellType(sourceType: string | undefined): string | undefined {
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (sourceType === undefined) return undefined
  /* v8 ignore next -- caller contract makes this path unreachable */
  return SHEET_CELL_TYPES.has(sourceType) ? sourceType : 'n'
}

function worksheetCachedValue(
  model: ICellData,
  sourceType: string | undefined,
): { readonly type: string | undefined; readonly body: string } | undefined {
  if (model.v === undefined || model.v === null) return undefined
  const text = String(model.v)
  if (model.t === CELL_BOOLEAN || typeof model.v === 'boolean') {
    /* v8 ignore next -- caller contract makes this path unreachable */
    return { type: 'b', body: `<v>${model.v === true || model.v === 1 ? '1' : '0'}</v>` }
  }
  if (sourceType === 'e' && XML_ERROR_VALUES.has(text)) {
    return { type: 'e', body: `<v>${escapeWorksheetText(text)}</v>` }
  }
  if (typeof model.v === 'string' || model.t === CELL_STRING) return { type: 'str', body: `<v>${escapeWorksheetText(text)}</v>` }
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (model.t === CELL_NUMBER || typeof model.v === 'number') {
    /* v8 ignore next -- caller contract makes this path unreachable */
    return { type: sourceType === undefined ? undefined : 'n', body: `<v>${text}</v>` }
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  return undefined
}

function worksheetCellValue(
  model: ICellData,
  sourceType: string | undefined,
): { readonly type: string | undefined; readonly body: string } | undefined {
  const value = model.v
  if (value === undefined || value === null) return undefined
  if (model.t === CELL_BOOLEAN || typeof value === 'boolean') {
    return { type: 'b', body: `<v>${value === true || value === 1 ? '1' : '0'}</v>` }
  }
  if (typeof value === 'number') {
    return { type: sourceType === undefined ? undefined : 'n', body: `<v>${String(value)}</v>` }
  }
  const text = value
  if (sourceType === 'e' && XML_ERROR_VALUES.has(text)) {
    return { type: 'e', body: `<v>${escapeWorksheetText(text)}</v>` }
  }
  return {
    type: 'inlineStr',
    body: `<is><t xml:space="preserve">${escapeWorksheetText(text)}</t></is>`,
  }
}

function worksheetSourceCellAt(sheet: WorksheetPartState, address: WorksheetAddress): WorksheetSourceCell | undefined {
  return sheet.cells.find(cell => cell.address.row === address.row && cell.address.column === address.column)
}

function worksheetModelCellAt(sheet: WorksheetPartState, address: WorksheetAddress): WorksheetModelCell | undefined {
  return sheet.modelCells.find(cell => cell.address.row === address.row && cell.address.column === address.column)
}

function sameWorksheetValue(left: ICellData, right: ICellData): boolean {
  return (left.v ?? null) === (right.v ?? null)
    && (left.f ?? null) === (right.f ?? null)
    && sameWorksheetDocument(left.p, right.p)
}

/**
 * Canonicalize one resolved model style for change detection.
 * @param style - Resolved style data, or `undefined` when the cell has none.
 * @returns Canonical JSON, with an empty style normalized to no style.
 */
function stylesCanonicalStyle(style: IStyleData | undefined): string {
  if (style === undefined || Object.keys(style).length === 0) return 'null'
  return worksheetCanonicalJson(style)
}

function sameWorksheetHyperlink(left: ICellData, right: ICellData): boolean {
  return sameWorksheetDocument(left.p, right.p)
}

function sameWorksheetDocument(
  left: ICellData['p'] | undefined,
  right: ICellData['p'] | undefined,
): boolean {
  return worksheetCanonicalJson(left ?? null) === worksheetCanonicalJson(right ?? null)
}

function worksheetCanonicalJson(value: unknown): string {
  if (value === null || value === undefined) return 'null'
  if (typeof value === 'string' || typeof value === 'boolean') return JSON.stringify(value)
  if (typeof value === 'number') return Number.isFinite(value) ? JSON.stringify(value) : 'null'
  if (Array.isArray(value)) return `[${value.map(worksheetCanonicalJson).join(',')}]`
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (isObjectRecord(value)) {
    const keys = Object.keys(value).sort()
    return `{${keys.map(key => `${JSON.stringify(key)}:${worksheetCanonicalJson(value[key])}`).join(',')}}`
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  return 'null'
}

/* ------------------------------ row render ------------------------------- */

function renderWorksheetRowUnit(
  state: WorksheetSpliceState,
  sheet: WorksheetPartState,
  unit: CurrentUnit,
): Uint8Array {
  const alignment = sheet.alignment
  const modelRow = sheet.modelRows.find(row => worksheetRowKey(sheet.sheetId, row.index) === unit.key)
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (alignment === undefined || modelRow === undefined) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw new OfficeExchangeError('conversion-failed', `Unit ${unit.key} names no model row in ${sheet.part}`)
  }
  const pair = alignment.pairs.find(candidate => candidate.model.index === modelRow.index)
  if (pair === undefined) return renderInsertedWorksheetRow(state, sheet, modelRow)
  return renderAlignedWorksheetRow(state, sheet, pair.source, modelRow)
}

function renderInsertedWorksheetRow(
  state: WorksheetSpliceState,
  sheet: WorksheetPartState,
  modelRow: WorksheetModelRow,
): Uint8Array {
  const rowName = worksheetQName(sheet.prefix, 'row')
  const cells = modelRow.cells.map((cell) => {
    const bytes = renderWorksheetCell(
      state,
      sheet,
      cell.address,
      undefined,
      undefined,
      cell.data,
    )
    return WORKBOOK_TEXT_DECODER.decode(bytes)
  })
  const xml = `<${rowName} r="${String(modelRow.index + 1)}">${cells.join('')}</${rowName}>`
  return WORKBOOK_TEXT_ENCODER.encode(xml)
}

function renderAlignedWorksheetRow(
  state: WorksheetSpliceState,
  sheet: WorksheetPartState,
  sourceRow: WorksheetSourceRowEntry,
  modelRow: WorksheetModelRow,
): Uint8Array {
  const rowBytes = sheet.bytes.subarray(sourceRow.row.start, sourceRow.row.end)
  const edits: WorksheetByteEdit[] = []
  const rowReference = worksheetAttribute(sourceRow.row.tag, 'r')
  if (rowReference !== undefined && sourceRow.index !== modelRow.index) {
    edits.push({
      start: rowReference.valueStart - sourceRow.row.start,
      end: rowReference.valueEnd - sourceRow.row.start,
      bytes: WORKBOOK_TEXT_ENCODER.encode(String(modelRow.index + 1)),
    })
  }

  const sourceCells = sourceRow.row.cells.map(cell => ({
    index: cell.address.column,
    signature: worksheetCellSignatureForAlignment(sheet, cell),
    cell,
  }))
  const modelCells = modelRow.cells.map(cell => ({
    index: cell.address.column,
    signature: cell.digest,
    cell,
  }))
  const alignment = alignWorksheetSequences(sourceCells, modelCells)

  for (const pair of alignment.pairs) {
    const sourceCell = pair.source.cell
    const modelCell = pair.model.cell
    const sourceModel = sheet.sourceCells.get(
      `${String(sourceCell.address.row)}:${String(sourceCell.address.column)}`,
    )
    if (pair.source.signature === pair.model.signature) {
      if (sourceCell.address.row !== modelCell.address.row
        || sourceCell.address.column !== modelCell.address.column) {
        const rewritten = rewriteWorksheetCellReference(sheet, sourceCell, modelCell.reference)
        /* v8 ignore next -- caller contract makes this path unreachable */
        if (rewritten !== undefined) edits.push(rewritten)
      }
      continue
    }
    const bytes = renderWorksheetCell(
      state,
      sheet,
      modelCell.address,
      sourceCell,
      sourceModel,
      modelCell.data,
    )
    edits.push({
      start: sourceCell.start - sourceRow.row.start,
      end: sourceCell.end - sourceRow.row.start,
      bytes,
    })
  }

  for (const deleted of alignment.deleted) {
    const sourceCell = deleted.cell
    const sourceModel = sheet.sourceCells.get(
      `${String(sourceCell.address.row)}:${String(sourceCell.address.column)}`,
    )
    if (sourceModel === undefined) continue
    edits.push({
      start: sourceCell.start - sourceRow.row.start,
      end: sourceCell.end - sourceRow.row.start,
      bytes: new Uint8Array(0),
    })
  }

  const inserted = alignment.inserted
  if (inserted.length > 0) {
    const pairByModel = new Map<number, { readonly start: number; readonly end: number }>()
    for (const pair of alignment.pairs) {
      pairByModel.set(pair.model.cell.address.column, {
        start: pair.source.cell.start - sourceRow.row.start,
        end: pair.source.cell.end - sourceRow.row.start,
      })
    }
    /* v8 ignore next -- caller contract makes this path unreachable */
    for (const entry of [...inserted].sort((left, right) => left.index - right.index)) {
      const at = worksheetCellInsertOffset(entry.index, alignment.pairs, pairByModel, sourceRow.row)
      const bytes = renderWorksheetCell(
        state,
        sheet,
        entry.cell.address,
        undefined,
        undefined,
        entry.cell.data,
      )
      edits.push({ start: at, end: at, bytes })
    }
  }

  return applyWorksheetEdits(rowBytes, edits)
}

function worksheetCellSignatureForAlignment(
  sheet: WorksheetPartState,
  cell: WorksheetSourceCell,
): string {
  const source = sheet.sourceCells.get(`${String(cell.address.row)}:${String(cell.address.column)}`)
  if (source !== undefined) return digestSheetCell(source)
  return createHash('sha256').update(JSON.stringify({
    style: cell.style ?? null,
    type: cell.type ?? null,
    formula: cell.formula?.text ?? null,
    /* v8 ignore next -- caller contract makes this path unreachable */
    value: cell.value ?? null,
    inline: cell.inline?.text ?? null,
  }), 'utf8').digest('hex')
}

function worksheetCellInsertOffset(
  column: number,
  pairs: readonly {
    readonly source: { readonly cell: WorksheetSourceCell }
    readonly model: { readonly cell: WorksheetModelCell }
  }[],
  pairByModel: ReadonlyMap<number, { readonly start: number; readonly end: number }>,
  row: WorksheetSourceRow,
): number {
  let previous: { readonly start: number; readonly end: number } | undefined
  let next: { readonly start: number; readonly end: number } | undefined
  for (const pair of pairs) {
    const span = pairByModel.get(pair.model.cell.address.column)
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (span === undefined) continue
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (pair.model.cell.address.column < column && (previous === undefined || span.end > previous.end)) {
      previous = span
    }
    if (pair.model.cell.address.column > column && (next === undefined || span.start < next.start)) {
      next = span
    }
  }
  if (previous !== undefined) return previous.end
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (next !== undefined) return next.start
  /* v8 ignore next -- caller contract makes this path unreachable */
  return (row.closeStart ?? row.end) - row.start
}

/** One byte replacement inside a row's bytes, relative to the row start. */
interface WorksheetByteEdit {
  /** Inclusive offset relative to the row's first byte. */
  readonly start: number
  /** Exclusive offset relative to the row's first byte. */
  readonly end: number
  /** Replacement bytes. */
  readonly bytes: Uint8Array
}

function applyWorksheetEdits(rowBytes: Uint8Array, edits: readonly WorksheetByteEdit[]): Uint8Array {
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (edits.length === 0) return rowBytes.slice()
  const sorted = [...edits].sort((left, right) => left.start - right.start || left.end - right.end)
  const operations: ApplySpliceOperation[] = sorted.map(edit => (
    edit.start === edit.end
      ? { kind: 'insert', at: edit.start, bytes: edit.bytes }
      : { kind: 'replace', start: edit.start, end: edit.end, bytes: edit.bytes }
  ))
  return applySplices(rowBytes, operations)
}

function rewriteWorksheetCellReference(
  sheet: WorksheetPartState,
  cell: WorksheetSourceCell,
  reference: string,
): WorksheetByteEdit | undefined {
  const attribute = cell.referenceAttribute
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (attribute === undefined) return undefined
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (attribute.value === reference) return undefined
  const row = worksheetRowContaining(sheet, cell)
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (row === undefined) return undefined
  return {
    start: attribute.valueStart - row.start,
    end: attribute.valueEnd - row.start,
    bytes: WORKBOOK_TEXT_ENCODER.encode(reference),
  }
}

function worksheetRowContaining(
  sheet: WorksheetPartState,
  cell: WorksheetSourceCell,
): WorksheetSourceRow | undefined {
  return sheet.rows.find(row => row.start <= cell.start && cell.end <= row.end)
}

/* --------------------------- apply and disposition ----------------------- */

function dimensionOperations(state: WorksheetSpliceState): readonly SpliceOperation[] {
  const operations: SpliceOperation[] = []
  for (const sheet of state.sheets) {
    const dimension = sheet.dimension
    if (dimension === undefined) continue
    const wanted = worksheetUsedRange(sheet)
    if (worksheetRangeEquals(worksheetSourceUsedRange(sheet), wanted)) continue
    const attribute = dimension.referenceAttribute
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (attribute === undefined) continue
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (worksheetRangeEquals(attribute.value, wanted)) continue
    operations.push({
      kind: 'replace',
      part: sheet.part,
      start: attribute.valueStart,
      end: attribute.valueEnd,
      bytes: WORKBOOK_TEXT_ENCODER.encode(wanted),
    })
  }
  return operations
}

function worksheetSourceUsedRange(sheet: WorksheetPartState): string {
  let minRow = Number.POSITIVE_INFINITY
  let minColumn = Number.POSITIVE_INFINITY
  let maxRow = 0
  let maxColumn = 0
  let count = 0
  for (const key of sheet.sourceCells.keys()) {
    const separator = key.indexOf(':')
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (separator < 0) continue
    const row = Number(key.slice(0, separator))
    const column = Number(key.slice(separator + 1))
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column)) continue
    minRow = Math.min(minRow, row)
    minColumn = Math.min(minColumn, column)
    maxRow = Math.max(maxRow, row)
    maxColumn = Math.max(maxColumn, column)
    count += 1
  }
  /* v8 ignore next -- caller contract makes this path unreachable */
  if (count === 0) return 'A1'
  const start = cellAddress(minRow + 1, minColumn + 1)
  const end = cellAddress(maxRow + 1, maxColumn + 1)
  return start === end ? start : `${start}:${end}`
}

function worksheetUsedRange(sheet: WorksheetPartState): string {
  if (sheet.modelCells.length === 0) return 'A1'
  let minRow = Number.POSITIVE_INFINITY
  let minColumn = Number.POSITIVE_INFINITY
  let maxRow = 0
  let maxColumn = 0
  for (const cell of sheet.modelCells) {
    minRow = Math.min(minRow, cell.address.row)
    minColumn = Math.min(minColumn, cell.address.column)
    maxRow = Math.max(maxRow, cell.address.row)
    maxColumn = Math.max(maxColumn, cell.address.column)
  }
  const start = cellAddress(minRow + 1, minColumn + 1)
  const end = cellAddress(maxRow + 1, maxColumn + 1)
  return start === end ? start : `${start}:${end}`
}

function worksheetRangeEquals(left: string, right: string): boolean {
  const normalize = (value: string): string => {
    const upper = value.trim().toUpperCase()
    const [first, second] = upper.split(':')
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (second === undefined || second === first) return first ?? ''
    return upper
  }
  return normalize(left) === normalize(right)
}

/**
 * Drop plan operations that would reproduce the source bytes. A shifted row
 * whose `r` attributes are implicit renders identically, and a replacement
 * with identical bytes is not a span the edit justifies.
 * @param state - Save state holding each worksheet part's source bytes.
 * @param operations - Plan operations before application.
 * @returns Non-carry operations that change at least one byte.
 */
function worksheetMeaningfulOperations(
  state: WorksheetSpliceState,
  operations: readonly SpliceOperation[],
): readonly SpliceOperation[] {
  return operations.filter((operation) => {
    if (operation.kind === 'carry') return false
    if (operation.kind === 'insert' || operation.kind === 'append') return operation.bytes.byteLength > 0
    if (operation.kind === 'delete') return true
    const sheet = state.partsByName.get(operation.part)
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (sheet === undefined || operation.end > sheet.bytes.byteLength) return true
    return !bytesEqual(operation.bytes, sheet.bytes.subarray(operation.start, operation.end))
  })
}

function sortWorksheetOperations(operations: readonly SpliceOperation[]): readonly SpliceOperation[] {
  return [...operations].sort((left, right) => {
    if (left.part !== right.part) return left.part < right.part ? -1 : 1
    return worksheetOperationStart(left) - worksheetOperationStart(right)
      || worksheetOperationRank(left) - worksheetOperationRank(right)
  })
}

function worksheetOperationStart(operation: SpliceOperation): number {
  switch (operation.kind) {
    case 'insert': return operation.at
    /* v8 ignore next -- caller contract makes this path unreachable */
    case 'append': return operation.at
    case 'replace': return operation.start
    case 'delete': return operation.start
    /* v8 ignore next -- caller contract makes this path unreachable */
    case 'carry': return 0
  }
}

function worksheetOperationRank(operation: SpliceOperation): number {
  switch (operation.kind) {
    case 'insert': return 0
    /* v8 ignore next -- caller contract makes this path unreachable */
    case 'append': return 1
    case 'replace': return 2
    case 'delete': return 3
    /* v8 ignore next -- caller contract makes this path unreachable */
    case 'carry': return 4
  }
}

function applyWorksheetOperations(
  state: WorksheetSpliceState,
  operations: readonly SpliceOperation[],
  signal: AbortSignal | undefined,
): { readonly bytes: Uint8Array; readonly operations: readonly SpliceOperation[] } {
  const byPart = byteOperationsByPart(operations)
  const effective: SpliceOperation[] = []
  const changedParts = new Set<string>()
  for (const [part, partOperations] of byPart) {
    const sheet = state.partsByName.get(part)
    const sourceBytes = sheet?.bytes ?? readWorksheetSourcePart(state.packageFile, part)
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (sourceBytes === undefined) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      throw new OfficeExchangeError('conversion-failed', `Splice operation names unknown part ${part}`)
    }
    const next = applySplices(sourceBytes, worksheetApplyOperations(partOperations))
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (bytesEqual(next, sourceBytes)) continue
    try {
      SyntaxValidator.validate(WORKBOOK_TEXT_DECODER.decode(next))
    } catch (error) {
      /* v8 ignore next -- caller contract makes this path unreachable */
      throw new OfficeExchangeError(
        'conversion-failed',
        `Spliced ${part} is not well-formed XML: ${messageOf(error)}`,
      )
    }
    state.packageFile.replacePartKnownDifferent(part, next)
    changedParts.add(part)
    effective.push(...partOperations)
    signal?.throwIfAborted()
  }
  assertSplicePackageGraph(state.packageFile)
  const bytes = state.packageFile.save()
  if (bytes.byteLength > state.bounds.maxFileBytes) {
    throw new OfficeExchangeError(
      'too-large',
      `Saved workbook is ${String(bytes.byteLength)} bytes; the limit is ${String(state.bounds.maxFileBytes)}`,
    )
  }
  return { bytes, operations: effective }
}

/**
 * Require the prospectively written package to be a closed OPC graph.
 * @param packageFile - Package holding the spliced parts.
 * @throws {OfficeExchangeError} `conversion-failed` when the output graph is incomplete.
 */
function assertSplicePackageGraph(packageFile: OfficePackage): void {
  try {
    assertOpcGraph(packageFile)
  } catch (error) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (error instanceof ZipError) {
      throw new OfficeExchangeError('conversion-failed', `The saved workbook package is not a valid OPC graph: ${error.message}`)
    }
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw error
  }
}

function worksheetApplyOperations(operations: readonly SpliceOperation[]): readonly ApplySpliceOperation[] {
  const result: ApplySpliceOperation[] = []
  for (const operation of operations) {
    switch (operation.kind) {
      case 'replace':
        result.push({ kind: 'replace', start: operation.start, end: operation.end, bytes: operation.bytes })
        break
      case 'insert':
        result.push({ kind: 'insert', at: operation.at, bytes: operation.bytes })
        break
      /* v8 ignore next -- caller contract makes this path unreachable */
      case 'append':
        /* v8 ignore next -- caller contract makes this path unreachable */
        result.push({ kind: 'append', at: operation.at, bytes: operation.bytes })
        /* v8 ignore next -- caller contract makes this path unreachable */
        break
      case 'delete':
        result.push({ kind: 'delete', start: operation.start, end: operation.end })
        break
      /* v8 ignore next -- caller contract makes this path unreachable */
      case 'carry':
        /* v8 ignore next -- caller contract makes this path unreachable */
        break
    }
  }
  return result
}

function readWorksheetSourcePart(packageFile: OfficePackage, part: string): Uint8Array | undefined {
  try {
    return packageFile.readPart(part)
  } catch (error) {
    /* v8 ignore next -- caller contract makes this path unreachable */
    if (error instanceof ZipError) return undefined
    /* v8 ignore next -- caller contract makes this path unreachable */
    throw error
  }
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
