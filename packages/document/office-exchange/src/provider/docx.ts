/**
 * DOCX mapping between OOXML documents and Univer `IDocumentData`.
 *
 * SUPPORTED: paragraph order and paragraph breaks, bold, italic, underline,
 * font family, font size, text color, highlight or shading, strike, vertical
 * alignment, caps, small caps, and run borders; tab characters inside runs;
 * paragraph alignment, spacing, line spacing, indentation, tab stops, outline
 * level, shading, and paragraph borders; document font and paragraph defaults
 * from `word/styles.xml`; named heading paragraph styles (`Heading1` through
 * `Heading5`); flow tables; section page setup; header/footer text; lists and
 * numbering; and hyperlinks. Table support covers rows, cells, cell text, grid
 * column widths, preferred cell widths, merged cells (`w:gridSpan` and
 * `w:vMerge`), cell borders, shading, vertical alignment, and margins, row
 * height and break/header properties, table width/alignment/indent/layout, the
 * paragraph/section-break pair each cell requires, and tables nested in a cell.
 * Section support maps `w:sectPr` page size, orientation,
 * margins, header/footer distances, columns, and break type into
 * `ISectionBreak`, and maps header and footer parts into
 * `IDocumentData.headers`/`footers` with their default, first-page, and
 * even-page references. Numbering maps `word/numbering.xml` definitions into
 * `IDocumentData.lists` and each paragraph's `w:numPr` into
 * `IParagraph.bullet`; hyperlinks map `w:hyperlink` into `\x1F`/`\x1E`
 * custom-range tokens with their relationship target or bookmark anchor.
 *
 * The installed Univer runtime tokens define the stream conventions used here:
 * `\r` ends a paragraph and `\n` ends a section. Tables use `\x1A` (table
 * start), `\x1B` (row start), `\x1C` (cell start), `\x1D` (cell end), `\x0E`
 * (row end), and `\x0F` (table end). `IParagraph.startIndex` points at its
 * paragraph break, `ISectionBreak.startIndex` points at its section break, and
 * `ITextRun.st`/`ed` are inclusive/exclusive UTF-16 offsets into `dataStream`.
 * A table range in `body.tables` ends immediately after the table end token and
 * its structural rows live in `tableSource[tableId]`.
 *
 * PRESERVED OPAQUELY: each run that carries a `w:drawing`, `w:pict`,
 * `w:object`, field marker, `w:footnoteReference`, or `w:endnoteReference` is
 * captured byte-for-byte in `body.docxRawBlocks`, as are paragraph-level
 * `w:pict`, `w:sdt`, `w:fldSimple`, and bookmark markers. Source media parts,
 * note parts, relationships, and content-type declarations are carried the same
 * way. The writer re-emits them unchanged, so pictures, shapes, embedded
 * objects, field markup, bookmarks, and note bodies survive an edit-and-save
 * without becoming editable or rendered content. A raw block anchors at the
 * stream offset before the run and adds no stream character, so paragraph and
 * run offsets are unchanged.
 *
 * A body-level `w:sdt` maps to no editable text: the reader inserts an empty
 * paragraph that occupies the unit slot, and the sidecar anchors that unit to
 * the whole `w:sdt` element with kind `contentControl`, so a sibling edit never
 * rewrites it and an edit of the unit itself is refused.
 *
 * IN-PLACE SAVE: `readDocumentWithSource` attaches
 * `data.officeExchange.source`, a per-unit anchor index over `word/document.xml`
 * and each header and footer part. `writeDocumentFromSource` recomputes every
 * unit's digest, replaces only changed anchor spans, merges independent source
 * properties the serializer does not emit, and appends required styles,
 * numbering definitions, and relationships without regenerating an auxiliary
 * part. A coupled construct, unit reordering, table geometry change, or missing
 * span refuses with a typed error and produces no bytes.
 *
 * UNSUPPORTED AS EDITABLE CONTENT: charts, drawing parts, shapes, comments,
 * tracked changes, macros, page borders and page numbering, theme styles and
 * theme-derived colors, named style definitions other than the heading styles
 * above, floating objects with their own drawing parts, and property values the
 * Univer model has no field for (for example a run border's line size or a
 * decimal tab stop's alignment). The editable model exposes none of these; a
 * regeneration round-trip drops them, while an in-place save carries them
 * verbatim outside the edited unit. Macro-bearing packages are read without
 * executing macros, and an in-place save keeps the VBA project and the
 * macro-enabled main content type unchanged.
 */
import { XMLParser } from 'fast-xml-parser'
import { SyntaxValidator } from 'fast-xml-validator'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import type { Unzipped, Zippable } from 'fflate'
import {
  BaselineOffset,
  BorderStyleTypes,
  BooleanNumber,
  BulletAlignment,
  ColumnSeparatorType,
  CustomRangeType,
  DashStyleType,
  HorizontalAlign,
  ListGlyphType,
  NumberUnitType,
  PageOrientType,
  SectionType,
  SpacingRule,
  TabStopAlignment,
  TabStopLeader,
  TableLayoutType,
  TextDecoration,
  VerticalAlignmentType,
} from '@univerjs/core'
import type {
  IBorderData,
  IBorderStyleData,
  IBullet,
  ICustomRange,
  ICustomTable,
  IDocumentBody,
  IDocumentData,
  IDocumentStyle,
  IDocxRawBlock,
  IFooters,
  IHeaders,
  IHyperlink,
  IHyperLinkCustomRange,
  IListData,
  ILists,
  INestingLevel,
  INumberUnit,
  IParagraph,
  IParagraphBorder,
  IParagraphStyle,
  ISectionBreak,
  ISectionColumnProperties,
  ITable,
  ITableCell,
  ITableCellBorder,
  ITableCellMargin,
  ITableRow,
  ITableRowSize,
  ITables,
  ITabStop,
  ITextDecoration,
  ITextRun,
  ITextStyle,
  IWidthInTableSize,
  LocaleType,
  NamedStyleType,
  TableAlignmentType,
  TableRowHeightRule,
  TableSizeType,
  TableTextWrapType,
} from '@univerjs/core'
import { type OfficeExchangeConfig, resolveConfig } from '../config.ts'
import { OfficeExchangeError } from '../errors.ts'
import { assertOpcGraph, open } from '../package/ooxml-package.ts'
import type { OfficePackage } from '../package/ooxml-package.ts'
import type { SourcePackage } from '../package/types.ts'
import { applySplices } from '../source/apply-splices.ts'
import type { ApplySpliceOperation } from '../source/apply-splices.ts'
import { buildSaveDisposition } from '../source/save-disposition.ts'
import { parseSourceState, SourceStateError } from '../source/source-state.ts'
import type { OfficeSourceFormat, OfficeSourceState, UnitAnchor } from '../source/source-state.ts'
import { buildSplicePlan, classifyIndependentProperties } from '../source/splice.ts'
import type {
  ConstructIntersection,
  CoupledConstructCandidate,
  CurrentUnit,
  IndependentProperty,
  IndependentPropertyClassification,
  SpliceAppend,
  SpliceOperation,
  SplicePlan,
} from '../source/splice.ts'
import { digestParagraph, digestTableCell } from '../source/unit-digest.ts'
import type { ParagraphUnit, TableCellUnit, TableRowUnit, TableUnit } from '../source/unit-digest.ts'
import { declarationEndInText, isXmlTextWhitespace, tagEndInText } from '../source/xml-text.ts'
import { scanSpans } from '../source/xml-spans.ts'
import type { ElementSpan, SpanKind } from '../source/xml-spans.ts'
import type { OfficeWriteResult } from '../snapshots.ts'
import type { UnsupportedFeatureKind } from '../types.ts'

const DOCUMENT_XML = 'word/document.xml'
const STYLES_XML = 'word/styles.xml'
const SETTINGS_XML = 'word/settings.xml'
const NUMBERING_XML = 'word/numbering.xml'
const DOCUMENT_RELS = 'word/_rels/document.xml.rels'
const HEADER_PART = /^word\/header[^/]*\.xml$/u
const FOOTER_PART = /^word\/footer[^/]*\.xml$/u
const STORY_RELS_PART = /^word\/_rels\/(?:header|footer)[^/]*\.xml\.rels$/u
const MEDIA_PART = /^word\/media\//u
const NOTE_PART = /^word\/(?:footnotes|endnotes)\.xml$/u
const NOTE_RELS_PART = /^word\/_rels\/(?:footnotes|endnotes)\.xml\.rels$/u
const CONTENT_TYPES_XML = '[Content_Types].xml'
const DEFAULT_COLUMN_SPACING = 720
const DEFAULT_PAGE_MARGIN = 1440 / 15
const PARAGRAPH = '\r'
const SECTION_BREAK = '\n'
const TABLE_START = '\x1A'
const TABLE_ROW_START = '\x1B'
const TABLE_CELL_START = '\x1C'
const TABLE_CELL_END = '\x1D'
const TABLE_ROW_END = '\x0E'
const TABLE_END = '\x0F'
const CUSTOM_RANGE_START = '\x1F'
const CUSTOM_RANGE_END = '\x1E'
const HYPERLINK_RANGE_TYPE: number = CustomRangeType.HYPERLINK
const TAB = '\t'

const ENTITY_COUNT_LIMIT = 10_000
const EXPANSION_BYTE_LIMIT = 200 * 1024 * 1024
const MAX_SOURCE_ENTRIES = 10_000
const SOURCE_ANCHOR_KINDS: readonly SpanKind[] = [
  'paragraph',
  'table',
  'tableRow',
  'tableCell',
  'contentControl',
  'preserved',
]
const TWIPS_PER_PIXEL = 15
const PIXELS_PER_POINT = 4 / 3
const POINTS_PER_PIXEL = 0.75
const BORDER_UNITS_PER_PIXEL = 6
const AUTO_LINE_SPACING_UNITS = 240
const HALF_POINTS_PER_POINT = 2
const DEFAULT_BORDER_SIZE = 4
const RUN_BORDER_SIZE = 8
const DEFAULT_BORDER_COLOR = '000000'
const DEFAULT_COLUMN_WIDTH = 160
const DEFAULT_ROW_HEIGHT = 30
const DETERMINISTIC_DATE = new Date('2026-01-01T00:00:00.000Z')

/** Media extension defaults used when a source package declares none. */
const IMAGE_CONTENT_TYPES: Readonly<Record<string, string>> = {
  bmp: 'image/bmp',
  emf: 'image/x-emf',
  gif: 'image/gif',
  jpe: 'image/jpeg',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  png: 'image/png',
  svg: 'image/svg+xml',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  wmf: 'image/x-wmf',
}

/** Content types of the note parts this provider preserves. */
const NOTE_CONTENT_TYPES: Readonly<Record<string, string>> = {
  'word/footnotes.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  'word/endnotes.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
}

/** Namespaces a preserved drawing run may reference without declaring them locally. */
const PRESERVED_NAMESPACES = ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
  + ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
  + ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
  + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
  + ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
  + ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
  + ' xmlns:v="urn:schemas-microsoft-com:vml"'
  + ' xmlns:o="urn:schemas-microsoft-com:office:office"'
  + ' xmlns:w10="urn:schemas-microsoft-com:office:word"'

const BOOLEAN_TRUE: BooleanNumber = 1
const TEXT_DECORATION_SINGLE: TextDecoration = 12
const TEXT_DECORATION_DOUBLE: TextDecoration = 10
const LOCALE_EN_US = 'enUS' as LocaleType
const TABLE_ALIGN_START: TableAlignmentType = 0
const TABLE_ALIGN_CENTER: TableAlignmentType = 1
const TABLE_ALIGN_END: TableAlignmentType = 2
const TABLE_WRAP_NONE: TableTextWrapType = 0
const TABLE_SIZE_UNSPECIFIED: TableSizeType = 0
const TABLE_SIZE_SPECIFIED: TableSizeType = 1
const TABLE_LAYOUT_AUTO: TableLayoutType = 0
const TABLE_LAYOUT_FIXED: TableLayoutType = 1
const TABLE_ROW_HEIGHT_AUTO: TableRowHeightRule = 0
const TABLE_ROW_HEIGHT_AT_LEAST: TableRowHeightRule = 1
const TABLE_ROW_HEIGHT_EXACT: TableRowHeightRule = 2
const RELATIVE_H_PAGE = 0
const RELATIVE_V_PAGE = 0

const HEADING_STYLE_IDS: readonly string[] = ['Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5']

const NAMED_STYLE_BY_ID: Readonly<Record<string, NamedStyleType>> = {
  Heading1: 4,
  Heading2: 5,
  Heading3: 6,
  Heading4: 7,
  Heading5: 8,
}

/** Defaults written when a snapshot declares no document text style. */
const DEFAULT_TEXT_STYLE: ITextStyle = { ff: 'Calibri', fs: 11 }

/** OOXML `w:highlight` names and their RGB values. */
const HIGHLIGHT_COLORS: Readonly<Record<string, string>> = {
  black: '000000',
  blue: '0000FF',
  cyan: '00FFFF',
  darkBlue: '000080',
  darkCyan: '008080',
  darkGray: '808080',
  darkGreen: '008000',
  darkMagenta: '800080',
  darkRed: '800000',
  darkYellow: '808000',
  green: '00FF00',
  lightGray: 'C0C0C0',
  magenta: 'FF00FF',
  red: 'FF0000',
  white: 'FFFFFF',
  yellow: 'FFFF00',
}

/** OOXML `w:u` values and their Univer text decoration. */
const UNDERLINE_VALUES: readonly (readonly [string, TextDecoration])[] = [
  ['single', TextDecoration.SINGLE],
  ['words', TextDecoration.WORDS],
  ['double', TextDecoration.DOUBLE],
  ['thick', TextDecoration.THICK],
  ['dotted', TextDecoration.DOTTED],
  ['dottedHeavy', TextDecoration.DOTTED_HEAVY],
  ['dash', TextDecoration.DASH],
  ['dashedHeavy', TextDecoration.DASHED_HEAVY],
  ['dashLong', TextDecoration.DASH_LONG],
  ['dashLongHeavy', TextDecoration.DASH_LONG_HEAVY],
  ['dotDash', TextDecoration.DOT_DASH],
  ['dotDotDash', TextDecoration.DOT_DOT_DASH],
  ['dashDotHeavy', TextDecoration.DASH_DOT_HEAVY],
  ['dashDotDotHeavy', TextDecoration.DASH_DOT_DOT_HEAVY],
  ['wave', TextDecoration.WAVE],
  ['wavyDouble', TextDecoration.WAVY_DOUBLE],
  ['wavyHeavy', TextDecoration.WAVY_HEAVY],
  ['singleAccounting', TextDecoration.SINGLE_ACCOUNTING],
  ['doubleAccounting', TextDecoration.DOUBLE_ACCOUNTING],
]

/** OOXML border values and their Univer paragraph-border dash style. */
const BORDER_DASH_STYLES: readonly (readonly [string, DashStyleType])[] = [
  ['single', DashStyleType.SOLID],
  ['thick', DashStyleType.SOLID],
  ['double', DashStyleType.SOLID],
  ['medium', DashStyleType.SOLID],
  ['dotted', DashStyleType.DOT],
  ['hair', DashStyleType.DOT],
  ['dashed', DashStyleType.DASH],
  ['dashDot', DashStyleType.DASH],
  ['dashDotDot', DashStyleType.DASH],
  ['dashLong', DashStyleType.DASH],
  ['dashLongHeavy', DashStyleType.DASH],
  ['mediumDashed', DashStyleType.DASH],
  ['mediumDashDot', DashStyleType.DASH],
  ['mediumDashDotDot', DashStyleType.DASH],
  ['slantDashDot', DashStyleType.DASH],
]

/** OOXML border values and the Univer run-border weight that stands in for them. */
const BORDER_WEIGHTS: readonly (readonly [string, BorderStyleTypes])[] = [
  ['single', BorderStyleTypes.THIN],
  ['hair', BorderStyleTypes.HAIR],
  ['dotted', BorderStyleTypes.DOTTED],
  ['dashed', BorderStyleTypes.DASHED],
  ['dashDot', BorderStyleTypes.DASH_DOT],
  ['dashDotDot', BorderStyleTypes.DASH_DOT_DOT],
  ['double', BorderStyleTypes.DOUBLE],
  ['medium', BorderStyleTypes.MEDIUM],
  ['mediumDashed', BorderStyleTypes.MEDIUM_DASHED],
  ['mediumDashDot', BorderStyleTypes.MEDIUM_DASH_DOT],
  ['mediumDashDotDot', BorderStyleTypes.MEDIUM_DASH_DOT_DOT],
  ['slantDashDot', BorderStyleTypes.SLANT_DASH_DOT],
  ['thick', BorderStyleTypes.THICK],
  ['threeDEmboss', BorderStyleTypes.MEDIUM],
  ['threeDEngrave', BorderStyleTypes.MEDIUM],
  ['inset', BorderStyleTypes.MEDIUM],
  ['outset', BorderStyleTypes.MEDIUM],
]

/** OOXML `w:jc` values and their Univer horizontal alignment. */
const HORIZONTAL_ALIGNMENTS: readonly (readonly [string, HorizontalAlign])[] = [
  ['left', HorizontalAlign.LEFT],
  ['start', HorizontalAlign.LEFT],
  ['center', HorizontalAlign.CENTER],
  ['right', HorizontalAlign.RIGHT],
  ['end', HorizontalAlign.RIGHT],
  ['both', HorizontalAlign.BOTH],
  ['justified', HorizontalAlign.JUSTIFIED],
  ['distribute', HorizontalAlign.DISTRIBUTED],
]

/** OOXML `w:tab` alignment values and their Univer tab stop alignment. */
const TAB_STOP_ALIGNMENTS: readonly (readonly [string, TabStopAlignment])[] = [
  ['left', TabStopAlignment.START],
  ['start', TabStopAlignment.START],
  ['center', TabStopAlignment.CENTER],
  ['right', TabStopAlignment.END],
  ['end', TabStopAlignment.END],
]

/** OOXML `w:tab` leader values and their Univer tab stop leader. */
const TAB_STOP_LEADERS: readonly (readonly [string, TabStopLeader])[] = [
  ['none', TabStopLeader.NONE],
  ['dot', TabStopLeader.DOT],
  ['hyphen', TabStopLeader.HYPHEN],
  ['underscore', TabStopLeader.UNDERSCORE],
  ['heavy', TabStopLeader.HEAVY],
  ['middleDot', TabStopLeader.MIDDLE_DOT],
]

/** OOXML `w:type` section values and their Univer section break kind. */
const SECTION_TYPES: readonly (readonly [string, SectionType])[] = [
  ['continuous', SectionType.CONTINUOUS],
  ['nextColumn', SectionType.NEXT_COLUMN],
  ['nextPage', SectionType.NEXT_PAGE],
  ['evenPage', SectionType.EVEN_PAGE],
  ['oddPage', SectionType.ODD_PAGE],
]

/** OOXML `w:numFmt` values and their Univer list glyph type. */
const NUMBER_FORMATS: readonly (readonly [string, ListGlyphType])[] = [
  ['decimal', ListGlyphType.DECIMAL],
  ['decimalZero', ListGlyphType.DECIMAL_ZERO],
  ['upperRoman', ListGlyphType.UPPER_ROMAN],
  ['lowerRoman', ListGlyphType.LOWER_ROMAN],
  ['upperLetter', ListGlyphType.UPPER_LETTER],
  ['lowerLetter', ListGlyphType.LOWER_LETTER],
  ['ordinal', ListGlyphType.ORDINAL],
  ['cardinalText', ListGlyphType.CARDINAL_TEXT],
  ['ordinalText', ListGlyphType.ORDINAL_TEXT],
  ['hex', ListGlyphType.HEX],
  ['chicago', ListGlyphType.CHICAGO],
  ['ideographDigital', ListGlyphType.IDEOGRAPH_DIGITAL],
  ['japaneseCounting', ListGlyphType.JAPANESE_COUNTING],
  ['aiueo', ListGlyphType.AIUEO],
  ['iroha', ListGlyphType.IROHA],
  ['decimalFullWidth', ListGlyphType.DECIMAL_FULL_WIDTH],
  ['decimalHalfWidth', ListGlyphType.DECIMAL_HALF_WIDTH],
  ['japaneseLegal', ListGlyphType.JAPANESE_LEGAL],
  ['japaneseDigitalTenThousand', ListGlyphType.JAPANESE_DIGITAL_TEN_THOUSAND],
  ['decimalEnclosedCircle', ListGlyphType.DECIMAL_ENCLOSED_CIRCLE],
  ['decimalFullWidth2', ListGlyphType.DECIMAL_FULL_WIDTH2],
  ['aiueoFullWidth', ListGlyphType.AIUEO_FULL_WIDTH],
  ['irohaFullWidth', ListGlyphType.IROHA_FULL_WIDTH],
  ['bullet', ListGlyphType.BULLET],
  ['ganada', ListGlyphType.GANADA],
  ['chosung', ListGlyphType.CHOSUNG],
  ['decimalEnclosedFullstop', ListGlyphType.DECIMAL_ENCLOSED_FULLSTOP],
  ['decimalEnclosedParen', ListGlyphType.DECIMAL_ENCLOSED_PAREN],
  ['decimalEnclosedCircleChinese', ListGlyphType.DECIMAL_ENCLOSED_CIRCLE_CHINESE],
  ['ideographEnclosedCircle', ListGlyphType.IDEOGRAPH_ENCLOSED_CIRCLE],
  ['ideographTraditional', ListGlyphType.IDEOGRAPH_TRADITIONAL],
  ['ideographZodiac', ListGlyphType.IDEOGRAPH_ZODIAC],
  ['ideographZodiacTraditional', ListGlyphType.IDEOGRAPH_ZODIAC_TRADITIONAL],
  ['taiwaneseCounting', ListGlyphType.TAIWANESE_COUNTING],
  ['ideographLegalTraditional', ListGlyphType.IDEOGRAPH_LEGAL_TRADITIONAL],
  ['taiwaneseCountingThousand', ListGlyphType.TAIWANESE_COUNTING_THOUSAND],
  ['taiwaneseDigital', ListGlyphType.TAIWANESE_DIGITAL],
  ['chineseCounting', ListGlyphType.CHINESE_COUNTING],
  ['chineseLegalSimplified', ListGlyphType.CHINESE_LEGAL_SIMPLIFIED],
  ['chineseCountingThousand', ListGlyphType.CHINESE_COUNTING_THOUSAND],
  ['koreanDigital', ListGlyphType.KOREAN_DIGITAL],
  ['koreanCounting', ListGlyphType.KOREAN_COUNTING],
  ['koreanLegal', ListGlyphType.KOREAN_LEGAL],
  ['koreanDigital2', ListGlyphType.KOREAN_DIGITAL2],
  ['vietnameseCounting', ListGlyphType.VIETNAMESE_COUNTING],
  ['russianLower', ListGlyphType.RUSSIAN_LOWER],
  ['russianUpper', ListGlyphType.RUSSIAN_UPPER],
  ['none', ListGlyphType.NONE],
  ['numberInDash', ListGlyphType.NUMBER_IN_DASH],
  ['hebrew1', ListGlyphType.HEBREW1],
  ['hebrew2', ListGlyphType.HEBREW2],
  ['arabicAlpha', ListGlyphType.ARABIC_ALPHA],
  ['arabicAbjad', ListGlyphType.ARABIC_ABJAD],
  ['hindiVowels', ListGlyphType.HINDI_VOWELS],
  ['hindiConsonants', ListGlyphType.HINDI_CONSONANTS],
  ['hindiNumbers', ListGlyphType.HINDI_NUMBERS],
  ['hindiCounting', ListGlyphType.HINDI_COUNTING],
  ['thaiLetters', ListGlyphType.THAI_LETTERS],
  ['thaiNumbers', ListGlyphType.THAI_NUMBERS],
  ['thaiCounting', ListGlyphType.THAI_COUNTING],
]

/** Header/footer reference variants declared by `w:headerReference` and `w:footerReference`. */
type StoryVariant = 'default' | 'first' | 'even'

/** Section fields that address one header or footer variant. */
type StoryReferenceKey =
  | 'defaultHeaderId'
  | 'firstPageHeaderId'
  | 'evenPageHeaderId'
  | 'defaultFooterId'
  | 'firstPageFooterId'
  | 'evenPageFooterId'

/** Section fields that address one header variant. */
const HEADER_REFERENCE_KEYS: Readonly<Record<StoryVariant, StoryReferenceKey>> = {
  default: 'defaultHeaderId',
  first: 'firstPageHeaderId',
  even: 'evenPageHeaderId',
}

/** Section fields that address one footer variant. */
const FOOTER_REFERENCE_KEYS: Readonly<Record<StoryVariant, StoryReferenceKey>> = {
  default: 'defaultFooterId',
  first: 'firstPageFooterId',
  even: 'evenPageFooterId',
}

const XML_PARSER_OPTIONS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: true,
  preserveOrder: true,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
}

const TABLE_BORDERS = '<w:tblBorders>'
  + '<w:top w:val="single" w:sz="4" w:color="auto"/>'
  + '<w:left w:val="single" w:sz="4" w:color="auto"/>'
  + '<w:bottom w:val="single" w:sz="4" w:color="auto"/>'
  + '<w:right w:val="single" w:sz="4" w:color="auto"/>'
  + '<w:insideH w:val="single" w:sz="4" w:color="auto"/>'
  + '<w:insideV w:val="single" w:sz="4" w:color="auto"/>'
  + '</w:tblBorders>'

/** A dynamically shaped node produced by `fast-xml-parser` in preserved order. */
interface ParserNode {
  readonly [key: string]: unknown
}

/** One XML element with its already-validated children and attributes. */
interface XmlElement {
  /** Local element name with its namespace prefix removed. */
  readonly tag: string
  /** Child nodes in source order, including text nodes. */
  readonly children: readonly ParserNode[]
  /** Attribute values keyed by the parser's attribute prefix. */
  readonly attributes: ParserNode
}

/** Accumulates a document stream and the Univer metadata addressing it. */
interface ReadContext {
  readonly parts: string[]
  length: number
  readonly paragraphs: IParagraph[]
  readonly textRuns: ITextRun[]
  readonly sectionBreaks: ISectionBreak[]
  readonly tables: ICustomTable[]
  readonly tableSource: ITables
  readonly customRanges: IHyperLinkCustomRange[]
  readonly rawRuns: RawRunSource | undefined
  readonly rawElements: RawElementSource | undefined
  readonly rawBlocks: IDocxRawBlock[]
  readonly bounds: OfficeExchangeConfig
  readonly anchored: boolean
  readonly syntheticParagraphs: Set<number>
  cells: number
  tableSerial: number
}

/** One part whose source anchors a read may record. */
interface StoryAnchorSource {
  /** Decoded package entry name. */
  readonly part: string
  /** Uncompressed part bytes. */
  readonly bytes: Uint8Array
  /** Editable body read from the part. */
  readonly body: IDocumentBody
  /** Table metadata addressing the body. */
  readonly tableSource: ITables
  /** Stream offsets of paragraphs the reader synthesized without source XML. */
  readonly syntheticParagraphs: ReadonlySet<number>
}

/** Collects one header or footer part's editable units while its body is read. */
type StoryAnchorCollector = (source: StoryAnchorSource) => void

/**
 * Numbering definitions read from `word/numbering.xml`, keyed by the list type
 * each paragraph's bullet names.
 */
interface NumberingCatalog {
  /** List definitions keyed by list type. */
  readonly lists: ILists
  /** `w:numId` mapped to the list type its definition produced. */
  readonly listTypeByNumId: ReadonlyMap<string, string>
  /** Per-instance level starts keyed by `${numId}:${level}`. */
  readonly startByNumLevel: ReadonlyMap<string, number>
}

/** One external hyperlink relationship a part's writer must declare. */
interface LinkRel {
  /** Relationship id used by `r:id`. */
  readonly relId: string
  /** External target URL. */
  readonly target: string
}

/** Numbering definitions and paragraph `w:numId` assignments for one write. */
interface NumberingPlan {
  /** Encoded `word/numbering.xml`, when the snapshot declares lists or bullets. */
  readonly xml: string | undefined
  /** List type mapped to the `w:numId` its paragraphs reference. */
  readonly numIdByListType: ReadonlyMap<string, number>
}

/** Hyperlink relationships for one part, keyed by their target URL. */
interface LinkPlan {
  /** Relationships to declare in the part's relationship XML. */
  readonly rels: readonly LinkRel[]
  /** Relationship id for a URL, when the part declares it. */
  readonly relIdOf: (url: string) => string | undefined
}

/** A hyperlink range clipped to one paragraph. */
interface ParagraphLink {
  /** Inclusive stream offset of the start token. */
  readonly start: number
  /** Exclusive stream offset after the end token. */
  readonly end: number
  /** Target URL or bookmark anchor. */
  readonly properties: IHyperlink
}

/** A table cell's content range inside `dataStream`. */
interface ParsedCell {
  /** Inclusive offset of the first content character. */
  readonly start: number
  /** Exclusive offset of the cell end token. */
  readonly end: number
}

/** A table parsed from `dataStream` together with its exclusive end. */
interface ParsedTable {
  /** Cell content ranges grouped by row. */
  readonly rows: ReadonlyArray<ReadonlyArray<ParsedCell>>
  /** Exclusive offset immediately after the table end token. */
  readonly end: number
}

/** Section fields mapped from one `w:sectPr` element. */
type SectionSetup = Omit<ISectionBreak, 'sectionId' | 'startIndex'>

/** One package relationship declared by `word/_rels/document.xml.rels`. */
interface DocumentRelationship {
  /** Relationship type URI; the trailing segment names the target role. */
  readonly type: string
  /** Target part name, relative to `word/`. */
  readonly target: string
}

/** Package relationship ids mapped to their declarations. */
type DocumentRelationships = ReadonlyMap<string, DocumentRelationship>

/** Header and footer part ids present in the package, used to reject dangling references. */
interface StoryPartIds {
  /** Header part identities. */
  readonly headers: ReadonlySet<string>
  /** Footer part identities. */
  readonly footers: ReadonlySet<string>
}

/** A document or header/footer body with the table metadata addressing it. */
interface StoryData {
  /** Paragraph, table, and run content. */
  readonly body?: IDocumentBody
  /** Table metadata keyed by table id. */
  readonly tableSource?: ITables
}

/** One header or footer part scheduled for the writer. */
interface StoryPart {
  /** Segment identity the section references. */
  readonly id: string
  /** Package entry name. */
  readonly partName: string
  /** Document relationship id addressing the part. */
  readonly relId: string
  /** Encoded part XML. */
  readonly xml: string
  /** Package entry name of the part's relationship XML, when it declares links. */
  readonly relsPartName?: string
  /** Encoded relationship XML for hyperlinks inside the part. */
  readonly relsXml?: string
}

/** Header or footer parts plus the relationship ids their sections use. */
interface StoryPlan {
  /** Parts in deterministic order. */
  readonly parts: readonly StoryPart[]
  /** Segment identity mapped to its relationship id. */
  readonly relIds: ReadonlyMap<string, string>
}

/** Section properties addressed from the document stream and the body's final section. */
interface SectionPlan {
  /** Non-final sections keyed by the stream offset of their section break token. */
  readonly breakAt: ReadonlyMap<number, ISectionBreak>
  /** Body-level `w:sectPr` for the final section. */
  readonly finalXml: string
}

/** One source package entry preserved byte-for-byte as base64. */
interface PreservedPart {
  /** Package entry name. */
  readonly name: string
  /** Base64 encoding of the entry bytes. */
  readonly base64: string
}

/** Source relationship elements the writer re-emits beside its generated ones. */
interface PreservedRelationships {
  /** Elements from `word/_rels/document.xml.rels`. */
  readonly document?: readonly string[]
  /** Elements from header and footer relationship parts, keyed by segment identity. */
  readonly stories?: Readonly<Record<string, readonly string[]>>
}

/** Opaque DOCX package state carried by a document snapshot outside the editable model. */
interface DocxPreservedPackage {
  /** Media, note, and note-relationship entries re-emitted unchanged. */
  readonly parts: readonly PreservedPart[]
  /** Relationship elements that address the preserved entries. */
  readonly relationships: PreservedRelationships
  /** Content-type declaration elements retained from `[Content_Types].xml`. */
  readonly contentTypes: readonly string[]
}

/** A document snapshot that carries the Exchange layer's opaque DOCX package state. */
interface OfficeExchangeDocumentData extends IDocumentData {
  /** Source package state the editable model cannot address. */
  officeExchange?: {
    /** DOCX state re-emitted on write. */
    readonly docx?: DocxPreservedPackage
    /** In-place save anchors read from the source package. */
    readonly source?: OfficeSourceState
  }
}

/** One source element captured verbatim for lossless re-emission. */
interface RawXmlRecord {
  /** Exact source XML from the opening tag through the closing tag. */
  readonly xml: string
  /** Relationship and note identity values used to match the parsed element. */
  readonly key: string
  taken: boolean
}

/** Source elements that a parsed element may consume once, matched by identity or order. */
interface RawXmlSource {
  /**
   * Take the source XML for one parsed element.
   * @param element - Parsed element the reader carries verbatim.
   * @returns The exact source element XML, or `undefined` when none remains.
   */
  take(element: XmlElement): string | undefined
}

/** Source runs that a parsed run may consume once, matched by identity or order. */
type RawRunSource = RawXmlSource

/** Paragraph-level elements the reader preserves instead of interpreting. */
const RAW_ELEMENT_TAGS: readonly string[] = [
  'pict',
  'sdt',
  'fldSimple',
  'bookmarkStart',
  'bookmarkEnd',
]
const RAW_ELEMENT_PATTERN = new RegExp(
  `<w:(${RAW_ELEMENT_TAGS.join('|')})(?=[\\s/>])[^>]*>`,
  'gu',
)

/** Raw paragraph-level elements that a parsed child may consume once, matched by identity or order. */
type RawElementSource = RawXmlSource

/** Raw-block positions for one body, marking which positions the writer emitted. */
interface RawBlockPlan {
  /** Source XML anchored at each stream offset, in source order. */
  readonly byIndex: ReadonlyMap<number, readonly string[]>
  /** Offsets already emitted, so one anchor never duplicates content. */
  readonly emitted: Set<number>
}

/** Sequential relationship ids that skip ids retained from the source package. */
interface RelationshipAllocator {
  /**
   * Allocate the next unused relationship id.
   * @returns An id of the form `rId<n>` absent from the source package.
   */
  next(): string
}

/**
 * Read `.docx` bytes into a Univer document snapshot. Paragraph text, paragraph
 * breaks, run and paragraph formatting, `word/styles.xml` defaults, heading
 * styles, flow tables, section page setup, header/footer text, list numbering,
 * and hyperlinks are mapped; the opaque constructs the module note lists are
 * captured as raw blocks, and the classes it marks unsupported as editable are
 * not.
 * @param bytes - Complete DOCX package bytes.
 * @param config - Bounds applied before parsing; defaults when omitted.
 * @returns Document snapshot with Univer stream tokens, run styles, table and
 * header/footer metadata, section page setup, and opaque preserved runs and
 * package entries.
 * @throws {OfficeExchangeError} `too-large` when a byte, package-entry,
 * expansion, or table-cell bound is exceeded; `conversion-failed` when the
 * input is not a readable Word document.
 */
export function readDocument(bytes: Uint8Array, config?: OfficeExchangeConfig): Promise<IDocumentData> {
  return Promise.resolve().then(() => readDocumentData(bytes, config, false).data)
}

/**
 * Read `.docx` or `.docm` bytes into a snapshot that also carries the
 * `data.officeExchange.source` sidecar when every editable unit resolves.
 *
 * Paragraph anchors are keyed by `(part, paragraphId)` and table-cell anchors by
 * `(part, tableId, row/cell path)`. A part the span scanner refuses, a unit the
 * model and source disagree on, or an anchor payload beyond the configured
 * budget leaves the sidecar absent, so the caller falls back to regeneration.
 * @param bytes - Complete DOCX package bytes.
 * @param format - Source format recorded in the sidecar.
 * @param config - Bounds applied before parsing and when validating anchors.
 * @returns Document snapshot with `para_*` paragraph ids and source anchors.
 * @throws {OfficeExchangeError} `too-large` when a bound is exceeded,
 * `conversion-failed` when the input is not a readable Word document.
 */
export function readDocumentWithSource(
  bytes: Uint8Array,
  format: 'docx' | 'docm',
  config?: OfficeExchangeConfig,
): Promise<IDocumentData> {
  return Promise.resolve().then(() => {
    const bounds = resolveConfig(config ?? {})
    const read = readDocumentData(bytes, bounds, true)
    const source = buildDocumentSourceState(bytes, read, format, bounds)
    if (source !== undefined) {
      read.data.officeExchange = { source }
    }
    return read.data
  })
}

/** One parsed document plus the parts an anchored read may anchor. */
interface DocumentRead {
  /** Editable snapshot. */
  readonly data: OfficeExchangeDocumentData
  /** Document and header/footer parts whose bodies were read. */
  readonly storySources: readonly StoryAnchorSource[]
}

function readDocumentData(
  bytes: Uint8Array,
  config: OfficeExchangeConfig | undefined,
  collectAnchors: boolean,
): DocumentRead {
  const bounds = resolveConfig(config ?? {})
  if (bytes.byteLength > bounds.maxFileBytes) {
    throw new OfficeExchangeError(
      'too-large',
      `Document is ${bytes.byteLength} bytes; the limit is ${bounds.maxFileBytes}`,
    )
  }
  const parts = readPackage(bytes)
  if (parts.document === undefined) {
    throw new OfficeExchangeError('conversion-failed', 'Package has no word/document.xml part')
  }
  const body = parseDocumentBody(parts.document)
  const relationships = readRelationships(parts.relationships)
  const numbering = parts.numbering === undefined ? emptyNumbering() : readNumbering(parts.numbering)
  const identifiers: StoryPartIds = {
    headers: new Set(parts.headers.keys()),
    footers: new Set(parts.footers.keys()),
  }
  const context = createReadContext(
    bounds,
    readRawRunCatalog(strFromU8(parts.document)),
    readRawElementCatalog(strFromU8(parts.document)),
    collectAnchors,
  )
  for (const child of childElements(body)) {
    if (child.tag === 'p') appendParagraph(context, child, relationships, identifiers, numbering)
    if (child.tag === 'tbl') appendTable(context, child, relationships, numbering)
    if (child.tag === 'sdt') appendContentControl(context)
  }
  const finalSection = findChild(body, 'sectPr')
  finalizeBody(context, finalSection === undefined ? undefined : readSectionSetup(finalSection, relationships, identifiers))
  const storySources: StoryAnchorSource[] = []
  const collect: StoryAnchorCollector | undefined = collectAnchors ? source => storySources.push(source) : undefined
  const storyRelationships = readStoryRelationships(parts.storyRelationships)
  const headers = readStoryParts(parts.headers, 'hdr', bounds, numbering, storyRelationships, (id, sectionBody, tableSource) => (
    { headerId: id, body: sectionBody, tableSource }
  ), collect)
  const footers = readStoryParts(parts.footers, 'ftr', bounds, numbering, storyRelationships, (id, sectionBody, tableSource) => (
    { footerId: id, body: sectionBody, tableSource }
  ), collect)
  const documentStyle = readDocumentStyle(parts.styles)
  if (readEvenAndOddHeaders(parts.settings)) documentStyle.evenAndOddHeaders = BOOLEAN_TRUE
  const documentBody: IDocumentBody = {
    dataStream: context.parts.join(''),
    textRuns: context.textRuns,
    paragraphs: context.paragraphs,
    sectionBreaks: context.sectionBreaks,
    tables: context.tables,
    customRanges: context.customRanges,
  }
  if (context.rawBlocks.length > 0) documentBody.docxRawBlocks = context.rawBlocks
  const snapshot: OfficeExchangeDocumentData = {
    id: 'office-exchange-document',
    locale: LOCALE_EN_US,
    title: '',
    body: documentBody,
    documentStyle,
    tableSource: context.tableSource,
    resources: [],
  }
  const preserved = collectPreservedPackage(parts)
  if (preserved !== undefined) snapshot.officeExchange = { docx: preserved }
  if (Object.keys(numbering.lists).length > 0) snapshot.lists = numbering.lists
  if (headers.size > 0) snapshot.headers = Object.fromEntries(headers)
  if (footers.size > 0) snapshot.footers = Object.fromEntries(footers)
  if (collectAnchors) {
    storySources.unshift({
      part: DOCUMENT_XML,
      bytes: parts.document,
      body: documentBody,
      tableSource: context.tableSource,
      syntheticParagraphs: context.syntheticParagraphs,
    })
  }
  return { data: snapshot, storySources }
}

/**
 * Write a Univer document snapshot to `.docx` bytes. The inverse of
 * `readDocument` for the supported features; preserved image and note runs and
 * their package entries are re-emitted, and the unsupported features listed in
 * the module note are absent from the output.
 * @param data - Document snapshot from `readDocument` or the editor.
 * @param config - Bounds validated before writing; defaults when omitted.
 * @returns A deterministic DOCX package that `readDocument` can reopen.
 * @throws {OfficeExchangeError} `too-large` when the table-cell bound is
 * exceeded; `conversion-failed` when the stream is structurally invalid,
 * preserved package state is malformed, or serialization fails.
 */
export function writeDocument(data: IDocumentData, config?: OfficeExchangeConfig): Promise<Uint8Array> {
  return Promise.resolve().then(() => writeDocumentBytes(data, config))
}

function writeDocumentBytes(data: IDocumentData, config?: OfficeExchangeConfig): Uint8Array {
  const bounds = resolveConfig(config ?? {})
  const preserved = preservedPackageOf(data)
  const stream = data.body?.dataStream ?? ''
  const paragraphs = new Map<number, IParagraph>()
  for (const paragraph of data.body?.paragraphs ?? []) paragraphs.set(paragraph.startIndex, paragraph)
  const numbering = planNumbering(data)
  const documentRelIds = createRelationshipAllocator(
    relationshipIdsOf(preserved?.relationships.document ?? []),
  )
  const stylesRelId = documentRelIds.next()
  const settings = writeSettingsXml(data.documentStyle)
  const settingsRelId = settings === undefined ? undefined : documentRelIds.next()
  const numberingRelId = numbering.xml === undefined ? undefined : documentRelIds.next()
  const headers = planStoryParts(data.headers, 'header', documentRelIds, bounds, numbering, preserved)
  const footers = planStoryParts(data.footers, 'footer', documentRelIds, bounds, numbering, preserved)
  const links = planLinks(data.body?.customRanges, documentRelIds)
  const relIds = new Map([...headers.relIds, ...footers.relIds])
  const sections = planSections(data.body?.sectionBreaks ?? [], relIds)
  const bodyXml = writeBody(stream, paragraphs, data, sections, relIds, links, numbering, bounds)
  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(contentTypesXml(
      headers.parts,
      footers.parts,
      settings !== undefined,
      numbering.xml !== undefined,
      preserved,
    )),
    '_rels/.rels': strToU8(rootRelsXml()),
    'word/_rels/document.xml.rels': strToU8(
      documentRelsXml(
        stylesRelId,
        settingsRelId,
        numberingRelId,
        headers.parts,
        footers.parts,
        links.rels,
        preserved?.relationships.document ?? [],
      ),
    ),
    [DOCUMENT_XML]: strToU8(documentXml(bodyXml, sections.finalXml)),
    [STYLES_XML]: strToU8(stylesXml(data.documentStyle)),
  }
  if (settings !== undefined) entries[SETTINGS_XML] = strToU8(settings)
  if (numbering.xml !== undefined) entries[NUMBERING_XML] = strToU8(numbering.xml)
  for (const part of [...headers.parts, ...footers.parts]) {
    entries[part.partName] = strToU8(part.xml)
    if (part.relsPartName !== undefined && part.relsXml !== undefined) {
      entries[part.relsPartName] = strToU8(part.relsXml)
    }
  }
  for (const part of preserved?.parts ?? []) {
    if (entries[part.name] !== undefined) {
      throw new OfficeExchangeError('conversion-failed', `Preserved part ${part.name} conflicts with a generated part`)
    }
    entries[part.name] = decodeBase64(part.base64)
  }
  const zippable: Zippable = {}
  for (const name of Object.keys(entries).sort()) {
    const content = entries[name]
    /* v8 ignore next -- Object.keys came from this record, so every lookup resolves. */
    if (content === undefined) continue
    zippable[name] = [content, { mtime: DETERMINISTIC_DATE }]
  }
  try {
    return zipSync(zippable, { level: 6 })
  } catch (error) {
    throw new OfficeExchangeError('conversion-failed', `Cannot write document: ${messageOf(error)}`)
  }
}

/**
 * Write a document snapshot back into the package it was read from by splicing
 * only the units the edit changed.
 *
 * Units whose digest still matches the sidecar are copied from the source
 * bytes, changed units replace their anchor span, new units insert at a
 * neighbouring anchor boundary, and removed units delete their span. Independent
 * unmapped properties inside a replaced unit are merged into the rendered unit;
 * a coupled construct, unit reordering, table geometry change, or unlocatable
 * span raises a typed refusal and produces no bytes.
 *
 * Auxiliary parts are append-only. A required style, numbering definition, or
 * relationship is appended before the part's closing tag, and orphaned elements
 * are left for the regeneration writer; `word/settings.xml` is never written.
 * @param snapshot - Editable document snapshot carrying `officeExchange.source`.
 * @param source - Complete source package bytes read at the caller's expected version.
 * @param config - Bounds validated before splicing; omitted fields keep their defaults.
 * @param signal - Optional cancellation checked before, during, and after splicing.
 * @returns Complete output bytes and the per-entry save disposition.
 * @throws {OfficeExchangeError} `too-large` when the source or output exceeds a
 * byte bound; `unsupported-feature` when a coupled construct overlaps an edited
 * span; `conversion-failed` when the snapshot has no usable source state, the
 * source bytes do not match it, a span cannot be located, or a spliced part is
 * not well-formed.
 */
export function writeDocumentFromSource(
  snapshot: IDocumentData,
  source: Uint8Array,
  config?: OfficeExchangeConfig,
  signal?: AbortSignal,
): Promise<OfficeWriteResult> {
  return Promise.resolve().then(() => writeDocumentSourceBytes(snapshot, source, resolveConfig(config ?? {}), signal))
}

function writeDocumentSourceBytes(
  snapshot: IDocumentData,
  source: Uint8Array,
  bounds: OfficeExchangeConfig,
  signal: AbortSignal | undefined,
): OfficeWriteResult {
  signal?.throwIfAborted()
  if (source.byteLength > bounds.maxFileBytes) {
    throw new OfficeExchangeError(
      'too-large',
      `Source document is ${String(source.byteLength)} bytes; the limit is ${String(bounds.maxFileBytes)}`,
    )
  }
  const packageFile = openSplicePackage(source)
  const state = createSpliceState(snapshot, packageFile, bounds)
  const rendered = new Map<string, Uint8Array>()
  const changedParts = new Set<string>()
  for (const part of state.parts) {
    const changed = part.units.filter(unit => !spliceUnitIsCarried(part, unit))
    if (changed.length > 0) changedParts.add(part.part)
    if (part.units.length !== part.anchors.length) changedParts.add(part.part)
    if (changed.length === 0) continue
    ensureSplicePartStyles(state, part)
    ensureSplicePartNumbering(state, part)
    for (const unit of changed) {
      rendered.set(spliceUnitIdentity(part.part, unit.kind, unit.key), renderSpliceUnit(state, part, unit))
    }
  }
  signal?.throwIfAborted()

  const appends = planSpliceAppends(state, changedParts)
  const classifications = new Map<string, IndependentPropertyClassification>()
  const coupledRefusal = { refused: false }
  const plan = buildSplicePlan({
    source: state.source,
    sourceState: state.sourceState,
    units: state.parts.flatMap(part => part.units),
    spans: state.parts.flatMap(part => part.spans),
    renderUnit: (unit) => {
      const bytes = rendered.get(spliceUnitIdentity(unit.part, unit.kind, unit.key))
      /* v8 ignore next 5 -- The render loop stores every unit buildSplicePlan can ask for. */
      if (bytes === undefined) {
        throw new OfficeExchangeError(
          'conversion-failed',
          `Unit ${unit.key} in ${unit.part} was not rendered before splicing`,
        )
      }
      return bytes
    },
    classifyCoupled: (edit) => {
      if (edit.kind !== 'replace') return undefined
      const part = state.partsByName.get(edit.part)
      const bytes = rendered.get(spliceUnitIdentity(edit.part, edit.unitKind, edit.key))
      /* v8 ignore next -- Replace edits come from parts and units already rendered above. */
      if (part === undefined || bytes === undefined) return undefined
      const classification = classifyIndependentProperties(part.bytes.subarray(edit.start, edit.end), bytes)
      classifications.set(
        spliceSpanKey(edit.part, edit.start, edit.end),
        classification,
      )
      const candidate = classification.coupled[0]
      if (candidate === undefined) return undefined
      coupledRefusal.refused = true
      return {
        kind: spliceCoupledKind(candidate),
        detail: `${candidate.name} in ${edit.part} at ${String(edit.start)}-${String(edit.end)}: ${candidate.detail}`,
        part: edit.part,
        start: edit.start,
      } satisfies ConstructIntersection
    },
    appends,
  })
  if (plan.refusal !== undefined) {
    throw new OfficeExchangeError(
      coupledRefusal.refused ? 'unsupported-feature' : 'conversion-failed',
      plan.refusal.detail,
    )
  }

  const merged = mergeSpliceOperations(state, plan, classifications)
  const operations = adjustSpliceInsertBoundaries(state, merged)
  const written = applySpliceOperations(state, operations, signal)
  return {
    bytes: written,
    disposition: buildSaveDisposition(plan.entries),
  }
}

/** One source part whose anchored units an in-place save can splice. */
interface SplicePart {
  /** Decoded part name. */
  readonly part: string
  /** Uncompressed source bytes of the part. */
  readonly bytes: Uint8Array
  /** Editable body mapped from the part. */
  readonly body: IDocumentBody
  /** Table metadata addressing the body's stream. */
  readonly tableSource: ITables
  /** Document data stream of the body. */
  readonly stream: string
  /** Paragraph records keyed by their stream break offset. */
  readonly paragraphs: ReadonlyMap<number, IParagraph>
  /** Paragraph records keyed by model identity. */
  readonly paragraphsByKey: ReadonlyMap<string, IParagraph>
  /** Every run in the body, in stream order. */
  readonly runs: readonly ITextRun[]
  /** Every table in the body with its cell content ranges. */
  readonly records: readonly TableRecord[]
  /** Table records keyed by model table id. */
  readonly tableRecords: ReadonlyMap<string, TableRecord>
  /** Source anchors recorded for the part. */
  readonly anchors: readonly UnitAnchor[]
  /** Anchor kind keyed by anchor identity. */
  readonly anchorKinds: ReadonlyMap<string, UnitAnchor['kind']>
  /** Read-time digest keyed by anchor identity. */
  readonly anchorDigests: ReadonlyMap<string, string>
  /** Current model units in source order. */
  readonly units: readonly CurrentUnit[]
  /** Spans the scanner located in the part. */
  readonly spans: readonly ElementSpan[]
  /** Hyperlink relationship state for the part. */
  readonly links: SpliceLinkState
}

/** Mutable state for the styles part the save may append to. */
interface SpliceStyleState {
  /** Source `word/styles.xml` bytes, when the package carries the part. */
  readonly bytes: Uint8Array | undefined
  /** Style ids the part already declares. */
  readonly present: Set<string>
  /** Complete `w:style` elements to append, in allocation order. */
  readonly appended: string[]
}

/** Mutable state for the numbering part the save may append to. */
interface SpliceNumberingState {
  /** Source `word/numbering.xml` bytes, when the package carries the part. */
  readonly bytes: Uint8Array | undefined
  /** List type mapped to the existing `w:numId` its paragraphs reference. */
  readonly numIdByListType: Map<string, number>
  /** Complete `w:abstractNum` elements to insert before the first instance. */
  readonly appendedAbstracts: string[]
  /** Complete `w:num` elements to append at the part's closing tag. */
  readonly appendedNums: string[]
  /** Per-list-type level starts the model declares. */
  readonly starts: ReadonlyMap<string, ReadonlyMap<number, number>>
  /** Next unused `w:numId`. */
  nextNumId: number
  /** Next unused `w:abstractNumId`. */
  nextAbstractId: number
}

/** Mutable hyperlink relationship state for one part. */
interface SpliceLinkState {
  /** Decoded part whose paragraphs may declare links. */
  readonly part: string
  /** Relationship part that declares the part's hyperlinks. */
  readonly relsPart: string
  /** Source relationship bytes, when the package carries the part. */
  readonly relsBytes: Uint8Array | undefined
  /** Target URL mapped to its declared relationship id. */
  readonly relIdByUrl: Map<string, string>
  /** Complete relationship elements to append. */
  readonly appended: string[]
  /** Link plan the paragraph serializer resolves URLs through. */
  readonly plan: LinkPlan
  /** Allocator that skips every id the source already declares. */
  readonly allocator: RelationshipAllocator
}

/** Everything one in-place document save needs across its parts. */
interface SpliceState {
  /** Opened source package. */
  readonly packageFile: OfficePackage
  /** Indexed source package as opened. */
  readonly source: SourcePackage
  /** Editable snapshot being written. */
  readonly data: IDocumentData
  /** Validated per-unit sidecar. */
  readonly sourceState: Extract<OfficeSourceState, { readonly anchorMode: 'per-unit' }>
  /** Validated bounds. */
  readonly bounds: OfficeExchangeConfig
  /** Editable parts in deterministic order. */
  readonly parts: readonly SplicePart[]
  /** Editable parts keyed by decoded part name. */
  readonly partsByName: ReadonlyMap<string, SplicePart>
  /** Styles state shared by every part. */
  readonly styles: SpliceStyleState
  /** Numbering state shared by every part. */
  readonly numbering: SpliceNumberingState
  /** Raw XML anchored at model stream offsets. */
  readonly rawBlocks: RawBlockPlan
  /** Table-cell count rendered so far, bounded by `maxCells`. */
  readonly cells: { count: number }
}

function openSplicePackage(source: Uint8Array): OfficePackage {
  try {
    return open(source, { maxEntries: MAX_SOURCE_ENTRIES })
  } catch (error) {
    throw new OfficeExchangeError('conversion-failed', `Cannot open the source package: ${messageOf(error)}`)
  }
}

function createSpliceState(
  data: IDocumentData,
  packageFile: OfficePackage,
  bounds: OfficeExchangeConfig,
): SpliceState {
  const value = documentSourceStateValue(data)
  if (value === undefined) {
    throw new OfficeExchangeError(
      'conversion-failed',
      'Document snapshot carries no source anchors; use the regeneration writer',
    )
  }
  let sourceState = parseDocumentSourceState(value, packageFile.source, bounds)
  if (sourceState.format !== 'docx' && sourceState.format !== 'docm') {
    throw new OfficeExchangeError(
      'unsupported-format',
      `Source state names a ${sourceState.format} package, not a document`,
    )
  }
  if (sourceState.anchorMode !== 'per-unit') {
    throw new OfficeExchangeError('conversion-failed', 'A document source state must carry per-unit anchors')
  }
  if (
    sourceState.packageDigest !== packageFile.source.digest
    || sourceState.byteLength !== packageFile.source.bytes.byteLength
  ) {
    sourceState = rebuildDocumentSourceState(packageFile.source.bytes, sourceState, data, bounds)
  }
  return createSpliceParts(data, packageFile, sourceState, bounds)
}

/**
 * Validate a document sidecar against the source package holding it.
 *
 * A sidecar written for different bytes can carry spans beyond the current
 * parts; those spans are replaced by a re-anchor after this structural check,
 * so only the part-name, identity, digest, and budget rules apply. A sidecar
 * that names the current package keeps its full span bounds.
 * @param value - Untrusted `data.officeExchange.source` value.
 * @param source - Source package read for this save.
 * @param bounds - Validated anchor count and byte budgets.
 * @returns The validated sidecar.
 * @throws {OfficeExchangeError} `conversion-failed` when the sidecar is malformed.
 */
function parseDocumentSourceState(
  value: unknown,
  source: SourcePackage,
  bounds: OfficeExchangeConfig,
): OfficeSourceState {
  const partByteLengths = new Map<string, number>()
  for (const entry of source.entries) {
    /* v8 ignore next -- source.entries never contains directory records. */
    if (!entry.name.endsWith('/')) partByteLengths.set(entry.name, entry.uncompressedSize)
  }
  const stale = !sidecarDescribesSource(value, source)
  const contextLengths = stale
    ? new Map([...partByteLengths.keys()].map(name => [name, Number.MAX_SAFE_INTEGER]))
    : partByteLengths
  try {
    return parseSourceState(value, {
      partByteLengths: contextLengths,
      maxSourceAnchors: bounds.maxSourceAnchors,
      maxSourceAnchorBytes: bounds.maxSourceAnchorBytes,
    })
  } catch (error) {
    /* v8 ignore start -- parseSourceState throws SourceStateError; a caller-supplied budget fails before this parse. */
    if (error instanceof SourceStateError) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `Document snapshot source state is malformed: ${error.message}`,
      )
    }
    throw error
    /* v8 ignore stop */
  }
}

function sidecarDescribesSource(value: unknown, source: SourcePackage): boolean {
  if (!isRecord(value)) return true
  const digest = value['packageDigest']
  const byteLength = value['byteLength']
  if (typeof digest === 'string' && digest !== source.digest) return false
  if (typeof byteLength === 'number' && byteLength !== source.bytes.byteLength) return false
  return true
}

/**
 * Re-anchor a source that changed since the snapshot was read.
 *
 * A successful save rewrites the file, so the snapshot's read-time offsets and
 * digests no longer describe it. Reading the current source again yields
 * anchors for its units; those units are aligned to the snapshot's model keys
 * so a second save diffs the model against the same identities. A source whose
 * units cannot be aligned unambiguously refuses the save; the caller's version
 * guard already proved the file is the one it expects.
 * @param source - Current source package bytes.
 * @param previous - Sidecar recorded for the bytes the snapshot was read from.
 * @param data - Snapshot whose model keys the anchors must use.
 * @param bounds - Validated anchor count and payload limits.
 * @returns Per-unit anchors for the current source.
 * @throws {OfficeExchangeError} `conversion-failed` when the source cannot be read, anchored, or aligned.
 */
function rebuildDocumentSourceState(
  source: Uint8Array,
  previous: Extract<OfficeSourceState, { readonly anchorMode: 'per-unit' }>,
  data: IDocumentData,
  bounds: OfficeExchangeConfig,
): Extract<OfficeSourceState, { readonly anchorMode: 'per-unit' }> {
  const read = readDocumentData(source, bounds, true)
  const rebuilt = buildDocumentSourceState(source, read, previous.format, bounds)
  if (rebuilt === undefined || rebuilt.anchorMode !== 'per-unit') {
    throw new OfficeExchangeError(
      'conversion-failed',
      'The current source cannot be re-anchored for an in-place save',
    )
  }
  const modelUnits = documentModelUnits(data, previous, rebuilt.sourceParts)
  const anchors = alignSourceAnchors(rebuilt.anchors, rebuilt.sourceParts, modelUnits)
  if (anchors === undefined) {
    throw new OfficeExchangeError(
      'conversion-failed',
      'The current source units cannot be aligned with the snapshot units',
    )
  }
  return { ...rebuilt, anchors }
}

/**
 * Project the snapshot's unit sequences for alignment with a rebuilt source.
 * @param data - Snapshot whose units are being aligned.
 * @param previous - Sidecar naming the unit kinds recorded at read.
 * @param parts - Story parts the rebuilt source anchors.
 * @returns Current units keyed by part, in stream order.
 * @throws {OfficeExchangeError} `conversion-failed` when the snapshot no longer carries a part.
 */
function documentModelUnits(
  data: IDocumentData,
  previous: Extract<OfficeSourceState, { readonly anchorMode: 'per-unit' }>,
  parts: readonly string[],
): ReadonlyMap<string, readonly CurrentUnit[]> {
  const unitsByPart = new Map<string, readonly CurrentUnit[]>()
  for (const part of parts) {
    const story = spliceStoryOf(data, part)
    if (story === undefined) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `The snapshot no longer carries the header or footer anchored to ${part}`,
      )
    }
    const anchorKinds = new Map<string, UnitAnchor['kind']>()
    for (const anchor of previous.anchors) {
      if (anchor.part === part) anchorKinds.set(spliceUnitIdentity(part, anchor.kind, anchor.key), anchor.kind)
    }
    const records = spliceTableRecords(part, story.body, story.tableSource)
    unitsByPart.set(
      part,
      spliceUnitsOf(part, story.body.dataStream, story.body, records, story.tableSource, anchorKinds),
    )
  }
  return unitsByPart
}

/**
 * Align rebuilt source anchors with the snapshot's model keys.
 * @param anchors - Fresh anchors built from the current source.
 * @param parts - Story parts the rebuilt source anchors.
 * @param modelUnits - Snapshot units keyed by part.
 * @returns Anchors carrying model keys, or `undefined` when alignment is ambiguous.
 */
function alignSourceAnchors(
  anchors: readonly UnitAnchor[],
  parts: readonly string[],
  modelUnits: ReadonlyMap<string, readonly CurrentUnit[]>,
): readonly UnitAnchor[] | undefined {
  const anchorsByPart = new Map<string, UnitAnchor[]>()
  for (const anchor of anchors) {
    const list = anchorsByPart.get(anchor.part)
    if (list === undefined) anchorsByPart.set(anchor.part, [anchor])
    else list.push(anchor)
  }
  for (const part of anchorsByPart.keys()) {
    /* v8 ignore next -- Rebuilt anchors and rebuilt sourceParts come from one read. */
    if (!parts.includes(part)) return undefined
  }
  const aligned: UnitAnchor[] = []
  for (const part of [...parts].sort(compareSpliceNames)) {
    /* v8 ignore next 3 -- Rebuilt parts map both sequences; equal starts leave the comparator's second operand unused. */
    const partAnchors = [...(anchorsByPart.get(part) ?? [])].sort((left, right) =>
      left.start - right.start || left.end - right.end)
    /* v8 ignore next -- documentModelUnits records every rebuilt part. */
    const partUnits = modelUnits.get(part) ?? []
    const result = alignPartAnchors(part, partAnchors, partUnits)
    if (result === undefined) return undefined
    aligned.push(...result)
  }
  return aligned
}

function alignPartAnchors(
  part: string,
  anchors: readonly UnitAnchor[],
  units: readonly CurrentUnit[],
): readonly UnitAnchor[] | undefined {
  const matches = digestMatches(anchors, units)
  const matchedAnchors = new Set(matches.map(match => match.anchor))
  const matchedUnits = new Set(matches.map(match => match.unit))
  const unmatchedAnchorDigests = new Set<string>()
  for (const [index, anchor] of anchors.entries()) {
    if (!matchedAnchors.has(index)) unmatchedAnchorDigests.add(anchor.digest)
  }
  for (const [index, unit] of units.entries()) {
    if (!matchedUnits.has(index) && unmatchedAnchorDigests.has(unit.digest)) return undefined
  }

  const reserved = new Set(units.map(unit => spliceUnitIdentity(part, unit.kind, unit.key)))
  const aligned: UnitAnchor[] = []
  let anchorCursor = 0
  let unitCursor = 0
  for (const match of [...matches, undefined]) {
    const anchorEnd = match?.anchor ?? anchors.length
    const unitEnd = match?.unit ?? units.length
    const gap = alignAnchorGap(part, anchors, units, anchorCursor, anchorEnd, unitCursor, unitEnd, reserved)
    if (gap === undefined) return undefined
    aligned.push(...gap)
    if (match === undefined) continue
    const anchor = anchors[match.anchor]
    const unit = units[match.unit]
    /* v8 ignore next -- Digest matches always name existing entries. */
    if (anchor === undefined || unit === undefined) return undefined
    aligned.push({ ...anchor, key: unit.key })
    reserved.add(spliceUnitIdentity(part, unit.kind, unit.key))
    anchorCursor = match.anchor + 1
    unitCursor = match.unit + 1
  }
  return aligned
}

interface AnchorDigestMatch {
  readonly anchor: number
  readonly unit: number
}

function digestMatches(
  anchors: readonly UnitAnchor[],
  units: readonly CurrentUnit[],
): readonly AnchorDigestMatch[] {
  const positions = new Map<string, number[]>()
  for (const [index, unit] of units.entries()) {
    const list = positions.get(unit.digest)
    if (list === undefined) positions.set(unit.digest, [index])
    else list.push(index)
  }
  const cursors = new Map<string, number>()
  const matches: AnchorDigestMatch[] = []
  let lastUnit = -1
  for (const [index, anchor] of anchors.entries()) {
    const list = positions.get(anchor.digest)
    if (list === undefined) continue
    let cursor = cursors.get(anchor.digest) ?? 0
    let candidate = list[cursor]
    while (candidate !== undefined && candidate <= lastUnit) {
      cursor += 1
      candidate = list[cursor]
    }
    const unit = list[cursor]
    if (unit === undefined) {
      cursors.set(anchor.digest, cursor)
      continue
    }
    matches.push({ anchor: index, unit })
    lastUnit = unit
    cursors.set(anchor.digest, cursor + 1)
  }
  return matches
}

function alignAnchorGap(
  part: string,
  anchors: readonly UnitAnchor[],
  units: readonly CurrentUnit[],
  anchorStart: number,
  anchorEnd: number,
  unitStart: number,
  unitEnd: number,
  reserved: Set<string>,
): readonly UnitAnchor[] | undefined {
  const anchorCount = anchorEnd - anchorStart
  const unitCount = unitEnd - unitStart
  if (anchorCount === 0) return []
  if (unitCount === 0) {
    const carried: UnitAnchor[] = []
    for (let offset = 0; offset < anchorCount; offset += 1) {
      const anchor = anchors[anchorStart + offset]
      /* v8 ignore next -- The loop index stays inside the anchor array. */
      if (anchor === undefined) return undefined
      carried.push(withUniqueKey(part, anchor, reserved))
    }
    return carried
  }
  if (anchorCount !== unitCount) return undefined
  const aligned: UnitAnchor[] = []
  for (let offset = 0; offset < anchorCount; offset += 1) {
    const anchor = anchors[anchorStart + offset]
    const unit = units[unitStart + offset]
    /* v8 ignore next 2 -- Equal gap lengths name existing entries. */
    if (anchor === undefined || unit === undefined) return undefined
    aligned.push({ ...anchor, key: unit.key })
    reserved.add(spliceUnitIdentity(part, unit.kind, unit.key))
  }
  return aligned
}

function withUniqueKey(part: string, anchor: UnitAnchor, reserved: Set<string>): UnitAnchor {
  let key = anchor.key
  let serial = 1
  while (reserved.has(spliceUnitIdentity(part, anchor.kind, key))) {
    key = `${anchor.key}~${String(serial)}`
    serial += 1
  }
  reserved.add(spliceUnitIdentity(part, anchor.kind, key))
  return { ...anchor, key }
}

function createSpliceParts(
  data: IDocumentData,
  packageFile: OfficePackage,
  sourceState: Extract<OfficeSourceState, { readonly anchorMode: 'per-unit' }>,
  bounds: OfficeExchangeConfig,
): SpliceState {
  const anchorsByPart = new Map<string, UnitAnchor[]>()
  for (const anchor of sourceState.anchors) {
    if (anchor.kind === 'sheetCell') {
      throw new OfficeExchangeError(
        'conversion-failed',
        `Document anchor ${anchor.key} names a sheet cell in ${anchor.part}`,
      )
    }
    const list = anchorsByPart.get(anchor.part)
    if (list === undefined) anchorsByPart.set(anchor.part, [anchor])
    else list.push(anchor)
  }
  const declaredParts = new Set(sourceState.sourceParts)
  for (const anchor of sourceState.anchors) {
    if (!declaredParts.has(anchor.part)) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `Anchor ${anchor.key} names ${anchor.part}, which sourceParts does not list`,
      )
    }
  }
  const parts: SplicePart[] = []
  for (const partName of [...sourceState.sourceParts].sort(compareSpliceNames)) {
    const story = spliceStoryOf(data, partName)
    if (story === undefined) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `The snapshot no longer carries the header or footer anchored to ${partName}`,
      )
    }
    const bytes = packageFile.readPart(partName)
    const scan = scanSpans(bytes, partName, SOURCE_ANCHOR_KINDS)
    if (scan.failure !== undefined) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `ambiguous-span: ${scan.failure.detail} in ${scan.failure.part} at byte ${String(scan.failure.offset)}`,
      )
    }
    const records = spliceTableRecords(partName, story.body, story.tableSource)
    const anchors = anchorsByPart.get(partName) ?? []
    const anchorKinds = new Map<string, UnitAnchor['kind']>()
    const anchorDigests = new Map<string, string>()
    for (const anchor of anchors) {
      const identity = spliceUnitIdentity(anchor.part, anchor.kind, anchor.key)
      anchorKinds.set(identity, anchor.kind)
      anchorDigests.set(identity, anchor.digest)
    }
    const stream = story.body.dataStream
    const paragraphs = new Map<number, IParagraph>()
    for (const paragraph of story.body.paragraphs ?? []) paragraphs.set(paragraph.startIndex, paragraph)
    parts.push({
      part: partName,
      bytes,
      body: story.body,
      tableSource: story.tableSource,
      stream,
      paragraphs,
      paragraphsByKey: new Map((story.body.paragraphs ?? []).map(paragraph => [paragraph.paragraphId, paragraph])),
      runs: story.body.textRuns ?? [],
      records,
      tableRecords: new Map(records.map(record => [record.table.tableId, record])),
      anchors,
      anchorKinds,
      anchorDigests,
      units: spliceUnitsOf(partName, stream, story.body, records, story.tableSource, anchorKinds),
      spans: scan.spans,
      links: createSpliceLinks(packageFile, partName),
    })
  }
  for (const headerId of Object.keys(data.headers ?? {})) {
    if (!declaredParts.has(`word/${headerId}.xml`)) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `The snapshot declares header ${headerId}, which an in-place save cannot add to the package`,
      )
    }
  }
  for (const footerId of Object.keys(data.footers ?? {})) {
    if (!declaredParts.has(`word/${footerId}.xml`)) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `The snapshot declares footer ${footerId}, which an in-place save cannot add to the package`,
      )
    }
  }
  return {
    packageFile,
    source: packageFile.source,
    data,
    sourceState,
    bounds,
    parts,
    partsByName: new Map(parts.map(part => [part.part, part])),
    styles: createSpliceStyles(packageFile),
    numbering: createSpliceNumbering(data, packageFile),
    rawBlocks: planRawBlocks(data.body),
    cells: { count: 0 },
  }
}

function documentSourceStateValue(data: IDocumentData): unknown {
  const extension: unknown = (data as OfficeExchangeDocumentData).officeExchange
  if (extension === undefined) return undefined
  if (!isRecord(extension)) {
    throw new OfficeExchangeError('conversion-failed', 'Document snapshot officeExchange state is malformed')
  }
  return extension['source']
}

/** One editable story body with the table metadata addressing it. */
interface SpliceStory {
  /** Editable body read from the part. */
  readonly body: IDocumentBody
  /** Table metadata addressing the body. */
  readonly tableSource: ITables
}

function spliceStoryOf(data: IDocumentData, part: string): SpliceStory | undefined {
  if (part === DOCUMENT_XML) {
    const body = data.body
    return body === undefined ? undefined : { body, tableSource: data.tableSource ?? {} }
  }
  const path = part.toLowerCase()
  const id = partIdOf(part)
  const header = HEADER_PART.test(path) ? data.headers?.[id] : undefined
  if (header !== undefined) return { body: header.body, tableSource: header.tableSource ?? {} }
  const footer = FOOTER_PART.test(path) ? data.footers?.[id] : undefined
  if (footer !== undefined) return { body: footer.body, tableSource: footer.tableSource ?? {} }
  return undefined
}

function spliceTableRecords(
  part: string,
  body: IDocumentBody,
  tableSource: ITables,
): readonly TableRecord[] {
  const tables = [...(body.tables ?? [])].sort((left, right) => left.startIndex - right.startIndex)
  const records: TableRecord[] = []
  for (const table of tables) {
    const model = tableSource[table.tableId]
    if (model === undefined) {
      throw new OfficeExchangeError('conversion-failed', `Table ${table.tableId} in ${part} has no model metadata`)
    }
    const ranges = readTableCellRanges(body.dataStream, table, model)
    if (ranges === undefined) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `Table ${table.tableId} in ${part} does not line up with its stream tokens`,
      )
    }
    records.push({ table, ranges })
  }
  return records
}

function spliceUnitsOf(
  part: string,
  stream: string,
  body: IDocumentBody,
  records: readonly TableRecord[],
  tableSource: ITables,
  anchorKinds: ReadonlyMap<string, UnitAnchor['kind']>,
): readonly CurrentUnit[] {
  const paragraphs = body.paragraphs ?? []
  const orderedParagraphs = sortedParagraphs(paragraphs)
  const runs = body.textRuns ?? []
  const paragraphByStart = new Map<number, IParagraph>()
  for (const paragraph of orderedParagraphs) paragraphByStart.set(paragraph.startIndex, paragraph)
  const recordByStart = new Map<number, TableRecord>()
  for (const record of records) recordByStart.set(record.table.startIndex, record)
  const units: CurrentUnit[] = []
  let index = 0
  while (index < stream.length) {
    const token = stream[index]
    if (token === TABLE_START) {
      const record = recordByStart.get(index)
      if (record === undefined) {
        throw new OfficeExchangeError(
          'conversion-failed',
          `Table token at stream offset ${String(index)} in ${part} has no model table`,
        )
      }
      for (const [row, cells] of record.ranges.entries()) {
        for (let cell = 0; cell < cells.length; cell += 1) {
          const unit = cellUnitOf(stream, orderedParagraphs, runs, records, tableSource, record.table.tableId, row, cell)
          /* v8 ignore next 5 -- cellUnitOf reads the same record and model row this loop walks. */
          if (unit === undefined) {
            throw new OfficeExchangeError(
              'conversion-failed',
              `Cell ${String(row)}/${String(cell)} of table ${record.table.tableId} in ${part} has no model content`,
            )
          }
          units.push({
            kind: 'tableCell',
            part,
            key: `${record.table.tableId}/${String(row)}/${String(cell)}`,
            digest: digestTableCell(unit),
          })
        }
      }
      index = record.table.endIndex
      continue
    }
    if (token === PARAGRAPH) {
      const paragraph = paragraphByStart.get(index)
      if (paragraph === undefined) {
        throw new OfficeExchangeError(
          'conversion-failed',
          `Paragraph break at stream offset ${String(index)} in ${part} has no model paragraph`,
        )
      }
      const key = paragraph.paragraphId
      const anchorKind = anchorKinds.get(spliceUnitIdentity(part, 'paragraph', key))
        ?? anchorKinds.get(spliceUnitIdentity(part, 'contentControl', key))
      if (anchorKind !== undefined || !isReaderSyntheticParagraph(stream, index)) {
        units.push({
          kind: anchorKind ?? 'paragraph',
          part,
          key,
          digest: digestParagraph(paragraphUnitOf(stream, paragraph, runs)),
        })
      }
      index += 1
      if (stream[index] === SECTION_BREAK) index += 1
      continue
    }
    index += 1
  }
  return units
}

/**
 * Report whether a paragraph is the empty spacer the reader inserted around a
 * table without a source `w:p` of its own. Such a paragraph has no anchor and
 * is not part of the editable unit sequence.
 * @param stream - Document data stream.
 * @param startIndex - Offset of the paragraph's break token.
 * @returns `true` for the reader's pre-table and trailing body spacers.
 */
function isReaderSyntheticParagraph(stream: string, startIndex: number): boolean {
  if (paragraphContentStart(stream, startIndex) !== startIndex) return false
  const previous = stream[startIndex - 1]
  const next = stream[startIndex + 1]
  if (next === TABLE_START) return true
  if (next === SECTION_BREAK) return startIndex === 0 || previous === TABLE_END
  return false
}

function spliceUnitIsCarried(part: SplicePart, unit: CurrentUnit): boolean {
  return part.anchorDigests.get(spliceUnitIdentity(part.part, unit.kind, unit.key)) === unit.digest
}

function spliceUnitIdentity(part: string, kind: string, key: string): string {
  return `${part}\u0000${kind}\u0000${key}`
}

function spliceSpanKey(part: string, start: number, end: number): string {
  return `${part}\u0000${String(start)}\u0000${String(end)}`
}

function createSpliceStyles(packageFile: OfficePackage): SpliceStyleState {
  const bytes = sourcePartBytes(packageFile, STYLES_XML)
  const present = new Set<string>()
  if (bytes !== undefined) {
    for (const match of strFromU8(bytes).matchAll(/<(?:[A-Za-z_][\w.-]*:)?style(?=[\s/>])[^>]*>/gu)) {
      const styleId = tagAttribute(match[0], 'w:styleId') ?? tagAttribute(match[0], 'styleId')
      if (styleId !== undefined) present.add(styleId)
    }
  }
  return { bytes, present, appended: [] }
}

function createSpliceNumbering(data: IDocumentData, packageFile: OfficePackage): SpliceNumberingState {
  const bytes = sourcePartBytes(packageFile, NUMBERING_XML)
  const numIdByListType = new Map<string, number>()
  let nextNumId = 1
  let nextAbstractId = 0
  if (bytes !== undefined) {
    const text = strFromU8(bytes)
    const catalog = readNumbering(bytes)
    for (const [numId, listType] of catalog.listTypeByNumId) {
      const numeric = Number(numId)
      if (!Number.isSafeInteger(numeric) || numeric < 0) continue
      numIdByListType.set(listType, numeric)
      nextNumId = Math.max(nextNumId, numeric + 1)
    }
    for (const match of text.matchAll(/<(?:[A-Za-z_][\w.-]*:)?abstractNum(?=[\s/>])[^>]*>/gu)) {
      nextAbstractId = Math.max(nextAbstractId, spliceElementId(match[0], 'abstractNumId') + 1)
    }
    for (const match of text.matchAll(/<(?:[A-Za-z_][\w.-]*:)?num(?=[\s/>])[^>]*>/gu)) {
      nextNumId = Math.max(nextNumId, spliceElementId(match[0], 'numId') + 1)
    }
  }
  return {
    bytes,
    numIdByListType,
    appendedAbstracts: [],
    appendedNums: [],
    starts: listStarts(data),
    nextNumId,
    nextAbstractId,
  }
}

function spliceElementId(element: string, name: string): number {
  const value = Number(tagAttribute(element, `w:${name}`) ?? tagAttribute(element, name))
  return Number.isSafeInteger(value) && value >= 0 ? value : -1
}

function createSpliceLinks(packageFile: OfficePackage, part: string): SpliceLinkState {
  const relsPart = part === DOCUMENT_XML ? DOCUMENT_RELS : `word/_rels/${part.slice('word/'.length)}.rels`
  const relsBytes = sourcePartBytes(packageFile, relsPart)
  const relationships = readRelationships(relsBytes)
  const relIdByUrl = new Map<string, string>()
  for (const [relId, relationship] of relationships) {
    if (relationship.type.endsWith('/hyperlink')) relIdByUrl.set(relationship.target, relId)
  }
  const allocator = createRelationshipAllocator(new Set(relationships.keys()))
  const appended: string[] = []
  const plan: LinkPlan = {
    rels: [],
    relIdOf: (url) => {
      const existing = relIdByUrl.get(url)
      if (existing !== undefined) return existing
      if (relsBytes === undefined) {
        throw new OfficeExchangeError(
          'conversion-failed',
          `The source package has no ${relsPart}, so a hyperlink in ${part} cannot be declared`,
        )
      }
      const relId = allocator.next()
      relIdByUrl.set(url, relId)
      appended.push(
        `<Relationship Id="${relId}"`
        + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'
        + ` Target="${escapeXml(url)}" TargetMode="External"/>`,
      )
      return relId
    },
  }
  return { part, relsPart, relsBytes, relIdByUrl, appended, allocator, plan }
}

function sourcePartBytes(packageFile: OfficePackage, name: string): Uint8Array | undefined {
  if (!packageFile.source.entries.some(entry => entry.name === name)) return undefined
  return packageFile.readPart(name)
}

function compareSpliceNames(left: string, right: string): number {
  return left < right ? -1 : 1
}

/** Styles part content type the reader maps. */
const SPLICE_STYLE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'
/** Numbering part content type the reader maps. */
const SPLICE_NUMBERING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'
/** Document main content type for a macro-free package. */
const SPLICE_DOCUMENT_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
/** Document main content type for a macro-enabled package. */
const SPLICE_DOCUMENT_MACRO_CONTENT_TYPE = 'application/vnd.ms-word.document.macroEnabled.main+xml'
/** Header part content type the reader maps. */
const SPLICE_HEADER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
/** Footer part content type the reader maps. */
const SPLICE_FOOTER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
/** Footnote part content type the reader maps. */
const SPLICE_FOOTNOTES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'
/** Endnote part content type the reader maps. */
const SPLICE_ENDNOTES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml'
/** Settings part content type the reader maps. */
const SPLICE_SETTINGS_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
/** OPC relationship part content type. */
const SPLICE_RELS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml'

/**
 * Render one current unit as the replacement bytes its anchor span receives.
 * @param state - Splice state carrying raw blocks, numbering, and bounds.
 * @param part - Part the unit belongs to.
 * @param unit - Current model unit.
 * @returns Serialized `<w:p>` or `<w:tc>` bytes.
 */
function renderSpliceUnit(state: SpliceState, part: SplicePart, unit: CurrentUnit): Uint8Array {
  const story: StoryData = { body: part.body, tableSource: part.tableSource }
  if (unit.kind === 'paragraph' || unit.kind === 'contentControl') {
    const paragraph = part.paragraphsByKey.get(unit.key)
    /* v8 ignore next -- Units are built from this part's paragraph map. */
    if (paragraph === undefined) {
      throw new OfficeExchangeError('conversion-failed', `Paragraph ${unit.key} in ${part.part} has no model record`)
    }
    const start = paragraphContentStart(part.stream, paragraph.startIndex)
    return strToU8(writeParagraph(
      part.stream,
      start,
      paragraph.startIndex,
      paragraph,
      story,
      '',
      part.links.plan,
      numberingPlanOf(state),
      state.rawBlocks,
    ))
  }
  const [tableId, rowText, columnText] = unit.key.split('/')
  const row = Number(rowText)
  const column = Number(columnText)
  if (tableId === undefined || !Number.isSafeInteger(row) || !Number.isSafeInteger(column)) {
    throw new OfficeExchangeError('conversion-failed', `Cell unit ${unit.key} in ${part.part} is not addressable`)
  }
  const record = part.tableRecords.get(tableId)
  const model = part.tableSource[tableId]
  const range = record?.ranges[row]?.[column]
  const sourceCell = model?.tableRows[row]?.tableCells[column]
  /* v8 ignore next -- The unit key comes from this table's own record and model. */
  if (record === undefined || model === undefined || range === undefined || sourceCell === undefined) {
    throw new OfficeExchangeError('conversion-failed', `Cell unit ${unit.key} in ${part.part} has no model record`)
  }
  const widths = spliceColumnWidths(model, record)
  const span = cellSpan(sourceCell)
  let gridColumn = 0
  for (let index = 0; index < column; index += 1) {
    gridColumn += cellSpan(model.tableRows[row]?.tableCells[index])
  }
  const width = sourceCell.size?.width.v ?? spanWidth(widths, gridColumn, span)
  const content = writeCell(
    part.stream,
    { start: range.start, end: range.end },
    part.paragraphs,
    story,
    state.bounds,
    state.cells,
    part.links.plan,
    numberingPlanOf(state),
    state.rawBlocks,
  )
  return strToU8(`<w:tc>${writeCellProperties(sourceCell, width, span)}${content}</w:tc>`)
}

function spliceColumnWidths(model: ITable, record: TableRecord): readonly number[] {
  const sourceColumns = model.tableColumns
  const rowWidths = record.ranges.map(row => row.length)
  const spanWidths = model.tableRows.map(row => (
    row.tableCells.reduce((count, cell) => count + cellSpan(cell), 0)
  ))
  const columnCount = Math.max(1, ...rowWidths, ...spanWidths, sourceColumns.length)
  const widths: number[] = []
  for (let index = 0; index < columnCount; index += 1) {
    /* v8 ignore next -- columnCount includes every span position a row can occupy. */
    widths.push(sourceColumns[index]?.size.width.v ?? DEFAULT_COLUMN_WIDTH)
  }
  return widths
}

function numberingPlanOf(state: SpliceState): NumberingPlan {
  return { xml: undefined, numIdByListType: state.numbering.numIdByListType }
}

function ensureSplicePartStyles(state: SpliceState, part: SplicePart): void {
  for (const paragraph of part.body.paragraphs ?? []) {
    const styleId = paragraph.styleId ?? styleIdForNamedStyle(paragraph.paragraphStyle?.namedStyleType)
    if (styleId !== undefined) ensureSpliceStyle(state, styleId)
  }
}

function ensureSpliceStyle(state: SpliceState, styleId: string): void {
  if (state.styles.present.has(styleId)) return
  if (state.styles.bytes === undefined) {
    throw new OfficeExchangeError(
      'conversion-failed',
      `The edit requires style ${styleId}, but the source package has no word/styles.xml to declare it`,
    )
  }
  state.styles.appended.push(spliceStyleXml(styleId))
  state.styles.present.add(styleId)
}

function spliceStyleXml(styleId: string): string {
  const headingIndex = HEADING_STYLE_IDS.indexOf(styleId)
  if (headingIndex >= 0) {
    return `<w:style w:type="paragraph" w:styleId="${escapeXml(styleId)}">`
      + `<w:name w:val="heading ${String(headingIndex + 1)}"/>`
      + '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
      + `<w:pPr><w:outlineLvl w:val="${String(headingIndex)}"/></w:pPr>`
      + '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
  }
  return `<w:style w:type="paragraph" w:styleId="${escapeXml(styleId)}">`
    + `<w:name w:val="${escapeXml(styleId)}"/></w:style>`
}

function ensureSplicePartNumbering(state: SpliceState, part: SplicePart): void {
  for (const paragraph of part.body.paragraphs ?? []) {
    const listType = paragraph.bullet?.listType
    if (listType !== undefined) ensureSpliceNumbering(state, listType)
  }
}

function ensureSpliceNumbering(state: SpliceState, listType: string): void {
  const numbering = state.numbering
  if (numbering.numIdByListType.has(listType)) return
  if (numbering.bytes === undefined) {
    throw new OfficeExchangeError(
      'conversion-failed',
      `The edit requires list ${listType}, but the source package has no word/numbering.xml to declare it`,
    )
  }
  const numId = numbering.nextNumId
  const abstractId = numbering.nextAbstractId
  numbering.nextNumId += 1
  numbering.nextAbstractId += 1
  const levels = state.data.lists?.[listType]?.nestingLevel ?? [defaultNestingLevel(0)]
  let abstractXml = `<w:abstractNum w:abstractNumId="${String(abstractId)}">`
    + '<w:multiLevelType w:val="hybridMultilevel"/>'
  for (const [index, level] of levels.entries()) abstractXml += writeNestingLevel(level, index)
  abstractXml += '</w:abstractNum>'
  let numXml = `<w:num w:numId="${String(numId)}">`
    + `<w:abstractNumId w:val="${String(abstractId)}"/>`
  const starts = numbering.starts.get(listType)
  if (starts !== undefined) {
    for (const [level, start] of [...starts.entries()].sort(([left], [right]) => left - right)) {
      numXml += `<w:lvlOverride w:ilvl="${String(level)}">`
        + `<w:startOverride w:val="${String(start)}"/></w:lvlOverride>`
    }
  }
  numXml += '</w:num>'
  numbering.appendedAbstracts.push(abstractXml)
  numbering.appendedNums.push(numXml)
  numbering.numIdByListType.set(listType, numId)
}

/**
 * Plan the auxiliary elements an edit adds to existing parts.
 * @param state - Splice state accumulated while rendering.
 * @param changedParts - Parts whose units differ from their anchors.
 * @returns Appends sorted by part, each before its part's closing tag.
 */
function planSpliceAppends(state: SpliceState, changedParts: ReadonlySet<string>): readonly SpliceAppend[] {
  const appends: SpliceAppend[] = []
  if (state.numbering.appendedAbstracts.length > 0 || state.numbering.appendedNums.length > 0) {
    const bytes = state.numbering.bytes
    /* v8 ignore next -- ensureSpliceNumbering records the source bytes with every append it plans. */
    if (bytes === undefined) throw new OfficeExchangeError('conversion-failed', 'Numbering append has no source part')
    const closing = findClosingTagOffset(bytes, 'numbering')
    const firstNum = findFirstStartTagOffset(bytes, 'num')
    const abstracts = state.numbering.appendedAbstracts.join('')
    const nums = state.numbering.appendedNums.join('')
    if (firstNum !== undefined && abstracts.length > 0) {
      appends.push({ part: NUMBERING_XML, at: firstNum, bytes: strToU8(abstracts) })
      appends.push({ part: NUMBERING_XML, at: closing, bytes: strToU8(nums) })
    } else {
      appends.push({ part: NUMBERING_XML, at: closing, bytes: strToU8(`${abstracts}${nums}`) })
    }
  }
  if (state.styles.appended.length > 0) {
    const bytes = state.styles.bytes
    /* v8 ignore next -- ensureSpliceStyle records the source bytes with every append it plans. */
    if (bytes === undefined) throw new OfficeExchangeError('conversion-failed', 'Styles append has no source part')
    appends.push({
      part: STYLES_XML,
      at: findClosingTagOffset(bytes, 'styles'),
      bytes: strToU8(state.styles.appended.join('')),
    })
  }
  const candidates = new Set(changedParts)
  if (state.numbering.appendedAbstracts.length > 0 || state.numbering.appendedNums.length > 0) {
    candidates.add(NUMBERING_XML)
  }
  if (state.styles.appended.length > 0) candidates.add(STYLES_XML)
  for (const part of [...state.parts].sort((left, right) => compareSpliceNames(left.part, right.part))) {
    if (part.links.appended.length === 0) continue
    const bytes = part.links.relsBytes
    /* v8 ignore next 4 -- createSpliceLinks records the relationship part bytes with every append it plans. */
    if (bytes === undefined) {
      throw new OfficeExchangeError('conversion-failed', `Relationship append for ${part.part} has no source part`)
    }
    candidates.add(part.links.relsPart)
    appends.push({
      part: part.links.relsPart,
      at: findClosingTagOffset(bytes, 'Relationships'),
      bytes: strToU8(part.links.appended.join('')),
    })
  }
  const contentTypes = planSpliceContentTypes(state, candidates)
  if (contentTypes !== undefined) appends.push(contentTypes)
  return appends
}

/**
 * Declare any touched part the source content-types manifest does not cover.
 * A well-formed source needs no append; a source missing a declaration for one
 * of its own parts still saves with an OPC-valid output.
 * @param state - Splice state with the source package and format.
 * @param touched - Parts the plan changes.
 * @returns One append for `[Content_Types].xml`, or `undefined` when every
 * touched part is already declared.
 */
function planSpliceContentTypes(
  state: SpliceState,
  touched: ReadonlySet<string>,
): SpliceAppend | undefined {
  const bytes = sourcePartBytes(state.packageFile, CONTENT_TYPES_XML)
  if (bytes === undefined) {
    throw new OfficeExchangeError('conversion-failed', 'The source package has no [Content_Types].xml part')
  }
  const text = strFromU8(bytes)
  const defaults = new Set<string>()
  const overrides = new Set<string>()
  for (const match of text.matchAll(/<(?:[A-Za-z_][\w.-]*:)?(Default|Override)(?=[\s/>])[^>]*>/gu)) {
    if (match[1] === 'Default') {
      const extension = tagAttribute(match[0], 'Extension')
      if (extension !== undefined) defaults.add(extension.toLowerCase())
      continue
    }
    const partName = tagAttribute(match[0], 'PartName')
    if (partName !== undefined) overrides.add(partName.replace(/^\/+/u, ''))
  }
  const elements: string[] = []
  for (const part of [...touched].sort(compareSpliceNames)) {
    if (part === CONTENT_TYPES_XML || overrides.has(part)) continue
    const extension = part.slice(part.lastIndexOf('.') + 1).toLowerCase()
    if (defaults.has(extension)) continue
    if (part.endsWith('.rels')) {
      elements.push(`<Default Extension="rels" ContentType="${SPLICE_RELS_CONTENT_TYPE}"/>`)
      defaults.add('rels')
      continue
    }
    const contentType = spliceContentTypeOf(state, part)
    /* v8 ignore next 4 -- Only known story, style, numbering, and relationship parts reach this plan. */
    if (contentType === undefined) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `The source package declares no content type for ${part}`,
      )
    }
    elements.push(`<Override PartName="/${part}" ContentType="${contentType}"/>`)
    overrides.add(part)
  }
  if (elements.length === 0) return undefined
  return {
    part: CONTENT_TYPES_XML,
    at: findClosingTagOffset(bytes, 'Types'),
    bytes: strToU8(elements.join('')),
  }
}

function spliceContentTypeOf(state: SpliceState, part: string): string | undefined {
  if (part === DOCUMENT_XML) {
    return state.sourceState.format === 'docm' ? SPLICE_DOCUMENT_MACRO_CONTENT_TYPE : SPLICE_DOCUMENT_CONTENT_TYPE
  }
  if (part === STYLES_XML) return SPLICE_STYLE_CONTENT_TYPE
  if (part === NUMBERING_XML) return SPLICE_NUMBERING_CONTENT_TYPE
  /* v8 ignore start -- Saved entries are only document, header, footer, style, numbering, and relationship parts. */
  if (part === SETTINGS_XML) return SPLICE_SETTINGS_CONTENT_TYPE
  if (part === 'word/footnotes.xml') return SPLICE_FOOTNOTES_CONTENT_TYPE
  if (part === 'word/endnotes.xml') return SPLICE_ENDNOTES_CONTENT_TYPE
  if (HEADER_PART.test(part.toLowerCase())) return SPLICE_HEADER_CONTENT_TYPE
  if (FOOTER_PART.test(part.toLowerCase())) return SPLICE_FOOTER_CONTENT_TYPE
  return undefined
  /* v8 ignore stop */
}

/**
 * Locate the byte offset of a part's first start tag with the given local name.
 * @param bytes - Encoded part XML.
 * @param localName - Element local name.
 * @returns Byte offset of its opening `<`, or `undefined` when the part has none.
 */
function findFirstStartTagOffset(bytes: Uint8Array, localName: string): number | undefined {
  const text = strFromU8(bytes)
  const pattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}(?=[\\s/>])`, 'u')
  const match = pattern.exec(text)
  if (match === null) return undefined
  return strToU8(text.slice(0, match.index)).byteLength
}

/**
 * Locate the byte offset of a part's root closing tag.
 * @param bytes - Encoded part XML.
 * @param localName - Root element local name.
 * @returns Byte offset of the final `</…localName>` start.
 */
function findClosingTagOffset(bytes: Uint8Array, localName: string): number {
  const text = strFromU8(bytes)
  const pattern = new RegExp(`</(?:[A-Za-z_][\\w.-]*:)?${localName}\\s*>`, 'gu')
  let character = -1
  let match = pattern.exec(text)
  while (match !== null) {
    character = match.index
    match = pattern.exec(text)
  }
  if (character < 0) {
    throw new OfficeExchangeError('conversion-failed', `Part has no closing ${localName} tag`)
  }
  return strToU8(text.slice(0, character)).byteLength
}

/** Property names the Word serializer reproduces from the model. */
const SPLICE_MAPPED_PROPERTIES: Readonly<
  Record<'pPr' | 'rPr' | 'tcPr' | 'trPr', ReadonlySet<string>>
> = {
  pPr: new Set(['pStyle', 'numPr', 'jc', 'spacing', 'ind', 'tabs', 'pBdr', 'shd', 'outlineLvl']),
  rPr: new Set([
    'rFonts', 'b', 'i', 'caps', 'smallCaps', 'dstrike', 'strike', 'color',
    'sz', 'szCs', 'shd', 'highlight', 'u', 'vertAlign', 'bdr',
  ]),
  tcPr: new Set(['tcW', 'gridSpan', 'vMerge', 'tcBorders', 'shd', 'tcMar', 'vAlign']),
  trPr: new Set(['trHeight', 'cantSplit', 'tblHeader']),
}

/** ECMA-376 child order for the Word property containers this writer merges. */
const SPLICE_PROPERTY_ORDER: Readonly<
  Record<'pPr' | 'rPr' | 'tcPr' | 'trPr', readonly string[]>
> = {
  pPr: [
    'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
    'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
    'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
    'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
    'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle',
    'rPr', 'sectPr', 'pPrChange',
  ],
  rPr: [
    'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
    'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color',
    'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd',
    'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath',
    'rPrChange',
  ],
  tcPr: [
    'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar',
    'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel',
    'cellMerge', 'tcPrChange',
  ],
  trPr: [
    'cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight',
    'tblHeader', 'tblCellSpacing', 'jc', 'hidden', 'ins', 'del', 'trPrChange',
  ],
}

/**
 * Keep only source properties the renderer cannot express.
 * @param properties - Merge candidates reported by the classifier.
 * @returns Properties whose element name the serializer never emits.
 * @throws {OfficeExchangeError} When a spreadsheet-root property appears in a Word unit.
 */
function spliceMergeProperties(
  properties: readonly IndependentProperty[],
): readonly IndependentProperty[] {
  const mergeable: IndependentProperty[] = []
  for (const property of properties) {
    /* v8 ignore next 5 -- The doct classifier never tags a Word property with the spreadsheet or root container. */
    if (property.container === 'c' || property.container === 'root') {
      throw new OfficeExchangeError(
        'conversion-failed',
        `A Word unit carries spreadsheet property ${property.name}`,
      )
    }
    if (!SPLICE_MAPPED_PROPERTIES[property.container].has(spliceLocalName(property.name))) {
      mergeable.push(property)
    }
  }
  return mergeable
}

function spliceCoupledKind(candidate: CoupledConstructCandidate): UnsupportedFeatureKind {
  if (candidate.detail.includes('tracked change') || candidate.detail.includes('revision-change')) {
    return 'tracked-changes'
  }
  if (candidate.detail.includes('comment range')) return 'comment'
  return 'other'
}

/**
 * Merge independent source properties into each replaced unit's rendered bytes.
 * @param state - Splice state carrying the source part bytes.
 * @param plan - Plan whose replace operations receive the merged bytes.
 * @param classifications - Merge candidates keyed by span.
 * @returns Operations with merged replacement bytes.
 */
function mergeSpliceOperations(
  state: SpliceState,
  plan: SplicePlan,
  classifications: ReadonlyMap<string, IndependentPropertyClassification>,
): readonly SpliceOperation[] {
  return plan.operations.map((operation) => {
    if (operation.kind !== 'replace') return operation
    const classification = classifications.get(spliceSpanKey(operation.part, operation.start, operation.end))
    /* v8 ignore next -- The classifier records a result for every replace edit it is shown. */
    if (classification === undefined) return operation
    const properties = spliceMergeProperties(classification.merge)
    if (properties.length === 0 && classification.rootAttributes.length === 0) return operation
    const part = state.partsByName.get(operation.part)
    /* v8 ignore next 2 -- Replace operations originate from a part in state.partsByName. */
    if (part === undefined) {
      throw new OfficeExchangeError('conversion-failed', `Replace operation names unknown part ${operation.part}`)
    }
    const sourceSpan = part.bytes.subarray(operation.start, operation.end)
    return {
      ...operation,
      bytes: mergeSpliceProperties(
        sourceSpan,
        operation.bytes,
        properties,
        classification.rootAttributes,
        operation.part,
      ),
    }
  })
}

/**
 * Re-emit unmapped source properties inside a regenerated unit.
 *
 * Each property is placed in the rendered container at the same structural
 * path it occupied in the source, ordered by the ECMA-376 property sequence. A
 * container the renderer omitted is created at its schema position; a path the
 * rendered unit no longer contains is a typed refusal because the property's
 * owner changed. Root metadata attributes are appended to the rendered root
 * start tag, whose carried source text is already XML-escaped.
 * @param sourceSpan - Source bytes of the replaced unit.
 * @param rendered - Renderer bytes for the same unit.
 * @param properties - Independent properties the renderer does not emit.
 * @param rootAttributes - Root metadata attributes the renderer does not emit.
 * @param part - Part name used in refusal details.
 * @returns The rendered unit with the source properties merged.
 */
function mergeSpliceProperties(
  sourceSpan: Uint8Array,
  rendered: Uint8Array,
  properties: readonly IndependentProperty[],
  rootAttributes: readonly IndependentProperty[],
  part: string,
): Uint8Array {
  const sourceRoot = parseSpliceXml(strFromU8(sourceSpan))
  const renderedText = strFromU8(rendered)
  const renderedRoot = parseSpliceXml(renderedText)
  /* v8 ignore next 3 -- Classification merges only spans whose source and rendered unit both parse. */
  if (sourceRoot === undefined || renderedRoot === undefined) {
    throw new OfficeExchangeError('conversion-failed', `Cannot inspect the replaced unit in ${part}`)
  }
  interface MergeGroup {
    readonly kind: 'pPr' | 'rPr' | 'tcPr' | 'trPr'
    readonly path: readonly SpliceXmlNode[]
    readonly properties: IndependentProperty[]
  }
  const groups = new Map<string, MergeGroup>()
  for (const property of properties) {
    /* v8 ignore next 2 -- The doct classifier never tags a Word property with the spreadsheet container. */
    if (property.container === 'c' || property.container === 'root') {
      throw new OfficeExchangeError('conversion-failed', `A Word unit in ${part} carries a spreadsheet property`)
    }
    const container = findSpliceContainer(
      sourceRoot,
      strFromU8(sourceSpan.subarray(0, property.start)).length,
      property.container,
    )
    /* v8 ignore next 4 -- Classifier offsets point inside the source container they name. */
    if (container === undefined) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `Property ${property.name} in ${part} has no ${property.container} container`,
      )
    }
    const parent = container.parent
    if (parent === undefined || !spliceContainerParentAllowed(property.container, parent.localName)) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `Property ${property.name} in ${part} is not inside a ${property.container} element`,
      )
    }
    const path = spliceNodePath(container)
    const key = `${property.container}/${path.map(node => String(node.siblingIndex)).join('.')}`
    const group = groups.get(key)
    if (group === undefined) groups.set(key, { kind: property.container, path, properties: [property] })
    else group.properties.push(property)
  }
  interface TextEdit {
    readonly start: number
    readonly end: number
    readonly text: string
    readonly order: number
    readonly sequence: number
  }
  const edits: TextEdit[] = []
  let sequence = 0
  for (const group of groups.values()) {
    const sorted = [...group.properties].sort((left, right) => (
      splicePropertyRank(group.kind, left.name) - splicePropertyRank(group.kind, right.name)
      || left.start - right.start
    ))
    const renderedContainer = resolveSplicePath(renderedRoot, group.path)
    if (renderedContainer === undefined) {
      const parentNode = group.path[group.path.length - 2]
      const renderedParent = resolveSplicePath(renderedRoot, group.path.slice(0, -1))
      if (parentNode === undefined || renderedParent === undefined || renderedParent.localName !== parentNode.localName) {
        throw new OfficeExchangeError(
          'conversion-failed',
          `Rendered unit in ${part} no longer carries the element that owns a merged property`,
        )
      }
      const children = sorted.map(property => property.xml).join('')
      edits.push({
        start: spliceContainerInsertOffset(renderedParent),
        end: spliceContainerInsertOffset(renderedParent),
        text: `<w:${group.kind}>${children}</w:${group.kind}>`,
        order: 0,
        sequence: sequence++,
      })
      continue
    }
    /* v8 ignore start -- Writer-emitted property containers always carry their matching close tag. */
    if (renderedContainer.closeStart < 0) {
      const children = sorted.map(property => property.xml).join('')
      edits.push({
        start: renderedContainer.start,
        end: renderedContainer.end,
        text: `<${renderedContainer.qname}>${children}</${renderedContainer.qname}>`,
        order: 0,
        sequence: sequence++,
      })
      continue
    }
    /* v8 ignore stop */
    for (const property of sorted) {
      const rank = splicePropertyRank(group.kind, property.name)
      const offset = spliceChildInsertOffset(renderedContainer, group.kind, rank)
      edits.push({ start: offset, end: offset, text: property.xml, order: rank, sequence: sequence++ })
    }
  }
  const opening = renderedText.slice(renderedRoot.start, renderedRoot.openEnd)
  for (const attribute of rootAttributes) {
    const existing = new RegExp(`(?:^|\\s)${escapeRegExpInText(attribute.name)}\\s*=`, 'u')
    /* v8 ignore next -- Classification reports only attributes the rendered root does not reproduce. */
    if (existing.test(opening)) continue
    /* v8 ignore next -- Rendered paragraph, cell, and control roots are never self-closing. */
    const offset = renderedText[renderedRoot.openEnd - 2] === '/' ? renderedRoot.openEnd - 2 : renderedRoot.openEnd - 1
    edits.push({ start: offset, end: offset, text: ` ${attribute.xml}`, order: 0, sequence: sequence++ })
  }
  /* v8 ignore next -- Every merge group or root attribute contributes at least one edit. */
  if (edits.length === 0) return rendered
  edits.sort((left, right) => (
    left.start - right.start || left.order - right.order || left.sequence - right.sequence
  ))
  let merged = renderedText
  for (let index = edits.length - 1; index >= 0; index -= 1) {
    const edit = edits[index]
    /* v8 ignore next -- The loop index stays inside the edit array. */
    if (edit === undefined) continue
    merged = merged.slice(0, edit.start) + edit.text + merged.slice(edit.end)
  }
  return strToU8(merged)
}

/** One element span in a fragment being merged, with its parent chain. */
interface SpliceXmlNode {
  /** Qualified name as written. */
  readonly qname: string
  /** Local name after the namespace prefix. */
  readonly localName: string
  /** Inclusive offset of the opening `<`. */
  readonly start: number
  /** Exclusive offset immediately after the opening tag. */
  readonly openEnd: number
  /** Offset of the closing `</`, or `-1` for a self-closing element. */
  closeStart: number
  /** Exclusive offset after the element's last `>`, or `-1` while unclosed. */
  end: number
  /** Enclosing element, absent on the fragment root. */
  readonly parent: SpliceXmlNode | undefined
  /** Index among the parent's element children. */
  readonly siblingIndex: number
  /** Direct element children in source order. */
  readonly children: SpliceXmlNode[]
}

/**
 * Parse a balanced XML fragment into element spans.
 * @param text - Fragment text.
 * @returns The root element, or `undefined` when the fragment is unbalanced.
 */
function parseSpliceXml(text: string): SpliceXmlNode | undefined {
  const stack: SpliceXmlNode[] = []
  let root: SpliceXmlNode | undefined
  let index = 0
  while (index < text.length) {
    const open = text.indexOf('<', index)
    /* v8 ignore next 2 -- An anchored unit span ends with the root closing tag. */
    if (open < 0) break
    if (text.startsWith('<!--', open)) {
      const close = text.indexOf('-->', open + 4)
      /* v8 ignore next -- scanSpans rejects an unterminated comment before this parser runs. */
      if (close < 0) return undefined
      index = close + 3
      continue
    }
    if (text.startsWith('<![CDATA[', open)) {
      const close = text.indexOf(']]>', open + 9)
      /* v8 ignore next -- scanSpans rejects an unterminated CDATA section before this parser runs. */
      if (close < 0) return undefined
      index = close + 3
      continue
    }
    if (text[open + 1] === '?') {
      const close = text.indexOf('?>', open + 2)
      /* v8 ignore next -- scanSpans rejects an unterminated processing instruction before this parser runs. */
      if (close < 0) return undefined
      index = close + 2
      continue
    }
    /* v8 ignore next 6 -- XML declarations inside an anchored element are not valid WordprocessingML. */
    if (text[open + 1] === '!') {
      const close = declarationEndInText(text, open + 2)
      if (close < 0) return undefined
      index = close
      continue
    }
    const close = tagEndInText(text, open + 1)
    /* v8 ignore next -- scanSpans rejects an unterminated tag before this parser runs. */
    if (close < 0) return undefined
    if (text[open + 1] === '/') {
      const name = spliceTagName(text, open + 2, close)
      const node = stack.pop()
      /* v8 ignore next -- scanSpans rejects mismatched closing tags before this parser runs. */
      if (node === undefined || name === undefined || node.qname !== name) return undefined
      node.closeStart = open
      node.end = close + 1
      index = close + 1
      continue
    }
    const selfClosing = text[close - 1] === '/'
    const name = spliceTagName(text, open + 1, selfClosing ? close - 1 : close)
    /* v8 ignore next -- A tag scanned by scanSpans always carries a name. */
    if (name === undefined) return undefined
    const parent = stack.at(-1)
    /* v8 ignore next -- An anchored unit span holds exactly one root element. */
    if (parent === undefined && root !== undefined) return undefined
    const node: SpliceXmlNode = {
      qname: name,
      localName: spliceLocalName(name),
      start: open,
      openEnd: close + 1,
      closeStart: selfClosing ? close + 1 : -1,
      end: selfClosing ? close + 1 : -1,
      parent,
      siblingIndex: parent === undefined ? 0 : parent.children.length,
      children: [],
    }
    parent?.children.push(node)
    if (parent === undefined) root = node
    if (!selfClosing) stack.push(node)
    index = close + 1
  }
  /* v8 ignore next -- An anchored unit span is a balanced element. */
  if (stack.length > 0 || root === undefined) return undefined
  return root
}

/**
 * Escape a qualified name for a regular expression.
 * @param name - XML qualified name.
 * @returns The name with regular-expression metacharacters escaped.
 */
function escapeRegExpInText(name: string): string {
  return name.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function spliceTagName(text: string, from: number, limit: number): string | undefined {
  let end = from
  while (end < limit) {
    const character = text[end]
    if (character === undefined || character === '>' || character === '/' || isXmlTextWhitespace(character)) break
    end += 1
  }
  /* v8 ignore next -- scanSpans rejects an empty tag name before this parser runs. */
  return end > from ? text.slice(from, end) : undefined
}

function spliceLocalName(name: string): string {
  const colon = name.indexOf(':')
  return colon < 0 ? name : name.slice(colon + 1)
}

function findSpliceContainer(
  node: SpliceXmlNode,
  offset: number,
  kind: 'pPr' | 'rPr' | 'tcPr' | 'trPr',
): SpliceXmlNode | undefined {
  for (const child of node.children) {
    const match = findSpliceContainer(child, offset, kind)
    if (match !== undefined) return match
  }
  /* v8 ignore next -- Resolved containers come from a fully parsed element. */
  const end = node.end >= 0 ? node.end : node.openEnd
  return node.localName === kind && node.start < offset && offset < end ? node : undefined
}

function spliceNodePath(node: SpliceXmlNode): readonly SpliceXmlNode[] {
  const path: SpliceXmlNode[] = []
  let current: SpliceXmlNode | undefined = node
  while (current !== undefined) {
    path.unshift(current)
    current = current.parent
  }
  return path
}

function resolveSplicePath(
  root: SpliceXmlNode,
  path: readonly SpliceXmlNode[],
): SpliceXmlNode | undefined {
  /* v8 ignore next -- Every merge group path starts at the unit root. */
  if (path.length === 0) return undefined
  let current = root
  for (const [index, expected] of path.entries()) {
    if (current.localName !== expected.localName) return undefined
    if (index === path.length - 1) return current
    const next = current.children[expected.siblingIndex]
    if (next === undefined) return undefined
    current = next
  }
  /* v8 ignore next -- The loop returns on its final path entry. */
  return current
}

function spliceContainerParentAllowed(
  kind: 'pPr' | 'rPr' | 'tcPr' | 'trPr',
  parentLocalName: string,
): boolean {
  if (kind === 'pPr') return parentLocalName === 'p'
  if (kind === 'rPr') return parentLocalName === 'r'
  if (kind === 'tcPr') return parentLocalName === 'tc'
  return parentLocalName === 'tr'
}

function splicePropertyRank(kind: 'pPr' | 'rPr' | 'tcPr' | 'trPr', name: string): number {
  const order = SPLICE_PROPERTY_ORDER[kind]
  const rank = order.indexOf(spliceLocalName(name))
  return rank < 0 ? order.length : rank
}

function spliceChildInsertOffset(
  container: SpliceXmlNode,
  kind: 'pPr' | 'rPr' | 'tcPr' | 'trPr',
  rank: number,
): number {
  for (const child of container.children) {
    if (splicePropertyRank(kind, child.localName) > rank) return child.start
  }
  return container.closeStart
}

function spliceContainerInsertOffset(parent: SpliceXmlNode): number {
  return parent.openEnd
}

/**
 * Move an insertion that lands between table cells to the containing table's
 * start or end, where a block-level paragraph or table is legal.
 * @param state - Splice state with the part anchors.
 * @param operations - Plan operations before application.
 * @returns Operations with block-level insertion offsets and a per-part sort.
 */
function adjustSpliceInsertBoundaries(
  state: SpliceState,
  operations: readonly SpliceOperation[],
): readonly SpliceOperation[] {
  const adjusted = operations.map((operation) => {
    if (operation.kind !== 'insert') return operation
    const part = state.partsByName.get(operation.part)
    /* v8 ignore next -- Insert operations originate from a part in state.partsByName. */
    if (part === undefined) return operation
    const at = spliceBlockInsertOffset(part, operation.at)
    return at === operation.at ? operation : { ...operation, at }
  })
  return [...adjusted].sort(compareSpliceOperations)
}

function spliceBlockInsertOffset(part: SplicePart, at: number): number {
  const containers = part.spans.filter(span => span.kind === 'table' && span.start < at && at < span.end)
  if (containers.length === 0) return at
  /* v8 ignore start -- Units skip table interiors, so an insertion never lands strictly inside a table span. */
  const table = containers.reduce((outer, span) => (span.start < outer.start ? span : outer))
  if (part.anchors.some(anchor => anchor.end === at)) return table.end
  if (part.anchors.some(anchor => anchor.start === at)) return table.start
  return at
  /* v8 ignore stop */
}

function compareSpliceOperations(left: SpliceOperation, right: SpliceOperation): number {
  if (left.part !== right.part) return compareSpliceNames(left.part, right.part)
  return spliceOperationStart(left) - spliceOperationStart(right)
    || spliceOperationRank(left) - spliceOperationRank(right)
}

function spliceOperationStart(operation: SpliceOperation): number {
  if (operation.kind === 'insert' || operation.kind === 'append') return operation.at
  /* v8 ignore next -- A carry operation is never sorted against another operation in the same part. */
  if (operation.kind === 'carry') return 0
  return operation.start
}

const SPLICE_OPERATION_RANKS: Readonly<Record<SpliceOperation['kind'], number>> = {
  insert: 0,
  append: 1,
  replace: 2,
  delete: 3,
  carry: 4,
}

function spliceOperationRank(operation: SpliceOperation): number {
  return SPLICE_OPERATION_RANKS[operation.kind]
}

/**
 * Apply the plan's byte operations and save the source package.
 * @param state - Splice state with the source package.
 * @param operations - Sorted plan operations.
 * @param signal - Optional cancellation between parts.
 * @returns Complete output package bytes.
 */
function applySpliceOperations(
  state: SpliceState,
  operations: readonly SpliceOperation[],
  signal: AbortSignal | undefined,
): Uint8Array {
  const byPart = new Map<string, ApplySpliceOperation[]>()
  for (const operation of operations) {
    if (operation.kind === 'carry') continue
    const list = byPart.get(operation.part) ?? []
    switch (operation.kind) {
      case 'replace':
        list.push({ kind: 'replace', start: operation.start, end: operation.end, bytes: operation.bytes })
        break
      case 'insert':
        list.push({ kind: 'insert', at: operation.at, bytes: operation.bytes })
        break
      case 'append':
        list.push({ kind: 'append', at: operation.at, bytes: operation.bytes })
        break
      case 'delete':
        list.push({ kind: 'delete', start: operation.start, end: operation.end })
        break
    }
    byPart.set(operation.part, list)
  }
  for (const [part, partOperations] of byPart) {
    const sourceBytes = state.partsByName.get(part)?.bytes ?? sourcePartBytes(state.packageFile, part)
    /* v8 ignore next 3 -- Every operation part is a source part appended to the package. */
    if (sourceBytes === undefined) {
      throw new OfficeExchangeError('conversion-failed', `Splice operation names unknown part ${part}`)
    }
    const next = applySplices(sourceBytes, partOperations)
    try {
      SyntaxValidator.validate(strFromU8(next))
    } catch (error) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `Spliced ${part} is not well-formed XML: ${messageOf(error)}`,
      )
    }
    state.packageFile.replacePart(part, next)
    signal?.throwIfAborted()
  }
  assertSplicePackageGraph(state.packageFile)
  const bytes = state.packageFile.save()
  if (bytes.byteLength > state.bounds.maxFileBytes) {
    throw new OfficeExchangeError(
      'too-large',
      `Saved document is ${String(bytes.byteLength)} bytes; the limit is ${String(state.bounds.maxFileBytes)}`,
    )
  }
  return bytes
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
    throw new OfficeExchangeError('conversion-failed', `The saved document package is not a valid OPC graph: ${messageOf(error)}`)
  }
}

/**
 * Read the snapshot's opaque DOCX package state after validating its JSON shape.
 * @param data - Document snapshot, possibly produced by a Remote round-trip.
 * @returns The preserved package, or `undefined` when the snapshot carries none.
 * @throws {OfficeExchangeError} `conversion-failed` when the state is malformed.
 */
function preservedPackageOf(data: IDocumentData): DocxPreservedPackage | undefined {
  const extension: unknown = (data as OfficeExchangeDocumentData).officeExchange
  if (extension === undefined) return undefined
  if (!isPreservedExtension(extension)) {
    throw new OfficeExchangeError('conversion-failed', 'Document snapshot has malformed preserved package state')
  }
  return extension.docx
}

function isPreservedExtension(value: unknown): value is OfficeExchangeDocumentData['officeExchange'] {
  if (!isRecord(value)) return false
  const docx = value['docx']
  if (docx === undefined) return true
  return isPreservedPackage(docx)
}

function isPreservedPackage(value: unknown): value is DocxPreservedPackage {
  if (!isRecord(value)) return false
  const parts = value['parts']
  const relationships = value['relationships']
  const contentTypes = value['contentTypes']
  if (!Array.isArray(parts) || !parts.every(isPreservedPart)) return false
  if (!Array.isArray(contentTypes) || !contentTypes.every(item => typeof item === 'string')) return false
  return isPreservedRelationships(relationships)
}

function isPreservedPart(value: unknown): value is PreservedPart {
  return isRecord(value) && typeof value['name'] === 'string' && typeof value['base64'] === 'string'
}

function isStringArray(value: unknown): value is readonly string[] {
  return Array.isArray(value) && value.every(item => typeof item === 'string')
}

function isPreservedRelationships(value: unknown): value is PreservedRelationships {
  if (!isRecord(value)) return false
  const document = value['document']
  const stories = value['stories']
  if (document !== undefined && !isStringArray(document)) return false
  if (stories !== undefined) {
    if (!isRecord(stories)) return false
    if (!Object.values(stories).every(isStringArray)) return false
  }
  return true
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

/**
 * Create sequential relationship ids while skipping ids the source retains.
 * @param reserved - Relationship ids present in preserved source elements.
 * @returns An allocator that never returns a reserved id.
 */
function createRelationshipAllocator(reserved: ReadonlySet<string>): RelationshipAllocator {
  let serial = 0
  return {
    next() {
      let id = ''
      do {
        serial += 1
        id = `rId${String(serial)}`
      } while (reserved.has(id))
      return id
    },
  }
}

/**
 * Collect relationship ids from preserved element source strings.
 * @param elements - Complete `<Relationship .../>` elements.
 * @returns Ids those elements declare.
 */
function relationshipIdsOf(elements: readonly string[]): ReadonlySet<string> {
  const ids = new Set<string>()
  for (const element of elements) {
    const id = tagAttribute(element, 'Id')
    if (id !== undefined) ids.add(id)
  }
  return ids
}

function createReadContext(
  bounds: OfficeExchangeConfig,
  rawRuns?: RawRunSource,
  rawElements?: RawElementSource,
  anchored = false,
): ReadContext {
  return {
    parts: [],
    length: 0,
    paragraphs: [],
    textRuns: [],
    sectionBreaks: [],
    tables: [],
    tableSource: {},
    customRanges: [],
    rawRuns,
    rawElements,
    rawBlocks: [],
    bounds,
    anchored,
    syntheticParagraphs: new Set(),
    cells: 0,
    tableSerial: 0,
  }
}

function emptyNumbering(): NumberingCatalog {
  return { lists: {}, listTypeByNumId: new Map(), startByNumLevel: new Map() }
}

/**
 * Read `word/numbering.xml` into list definitions and paragraph bullet lookups.
 * @param xml - Decoded numbering part.
 * @returns One list definition per `w:num`, plus the level starts per instance.
 * @throws {OfficeExchangeError} `conversion-failed` when the part is not well formed.
 */
function readNumbering(xml: Uint8Array): NumberingCatalog {
  const root = parseRootElement(xml, 'numbering', 'numbering XML')
  const abstracts = new Map<string, readonly INestingLevel[]>()
  const instances: XmlElement[] = []
  for (const child of childElements(root)) {
    if (child.tag === 'abstractNum') {
      const id = attribute(child, '@_abstractNumId')
      if (id !== undefined) abstracts.set(id, readNestingLevels(child))
    }
    if (child.tag === 'num') instances.push(child)
  }
  const lists: ILists = {}
  const listTypeByNumId = new Map<string, string>()
  const startByNumLevel = new Map<string, number>()
  for (const instance of instances) {
    const numId = attribute(instance, '@_numId')
    if (numId === undefined) continue
    const abstractId = optionalAttribute(findChild(instance, 'abstractNumId'), '@_val')
    const levels = (abstractId === undefined ? undefined : abstracts.get(abstractId)) ?? [defaultNestingLevel(0)]
    const definition: INestingLevel[] = [...levels]
    for (const override of childElements(instance)) {
      if (override.tag !== 'lvlOverride') continue
      const level = numberAttribute(override, '@_ilvl') ?? 0
      const start = numberAttribute(findChild(override, 'startOverride'), '@_val')
      if (start === undefined) continue
      definition[level] = { ...(definition[level] ?? defaultNestingLevel(level)), startNumber: start }
      startByNumLevel.set(`${numId}:${level}`, start)
    }
    const listType = listKeyOf(numId)
    lists[listType] = { listType, nestingLevel: definition }
    listTypeByNumId.set(numId, listType)
  }
  return { lists, listTypeByNumId, startByNumLevel }
}

function readNestingLevels(abstractNum: XmlElement): readonly INestingLevel[] {
  const byLevel = new Map<number, INestingLevel>()
  let index = 0
  for (const child of childElements(abstractNum)) {
    if (child.tag !== 'lvl') continue
    const level = numberAttribute(child, '@_ilvl') ?? index
    byLevel.set(level, readNestingLevel(child, level))
    index += 1
  }
  if (byLevel.size === 0) return [defaultNestingLevel(0)]
  const last = Math.max(...byLevel.keys())
  const levels: INestingLevel[] = []
  for (let level = 0; level <= last; level += 1) levels.push(byLevel.get(level) ?? defaultNestingLevel(level))
  return levels
}

function readNestingLevel(element: XmlElement, level: number): INestingLevel {
  const formatName = optionalAttribute(findChild(element, 'numFmt'), '@_val')
  const glyphType = formatName === undefined
    ? ListGlyphType.DECIMAL
    : lookup(NUMBER_FORMATS, formatName) ?? ListGlyphType.CUSTOM
  const text = optionalAttribute(findChild(element, 'lvlText'), '@_val') ?? ''
  const nesting: INestingLevel = {
    bulletAlignment: BulletAlignment.START,
    glyphFormat: text.length > 0 ? text : `%${String(level + 1)}.`,
    glyphType,
    startNumber: numberAttribute(findChild(element, 'start'), '@_val') ?? 0,
  }
  if (glyphType === ListGlyphType.BULLET) {
    nesting.glyphFormat = ` %${String(level + 1)}`
    if (text.length > 0) nesting.glyphSymbol = text
  }
  if (isEnabled(element, 'isLgl') === true) nesting.isLegal = true
  const paragraph = findChild(element, 'pPr')
  const paragraphProperties = paragraph === undefined ? undefined : readParagraphStyle(paragraph)
  if (paragraphProperties !== undefined) nesting.paragraphProperties = paragraphProperties
  const run = findChild(element, 'rPr')
  const textStyle = run === undefined ? undefined : readTextStyle(run)
  if (textStyle !== undefined) nesting.textStyle = textStyle
  return nesting
}

function defaultNestingLevel(level: number): INestingLevel {
  return {
    bulletAlignment: BulletAlignment.START,
    glyphFormat: `%${String(level + 1)}.`,
    glyphType: ListGlyphType.DECIMAL,
    startNumber: 0,
  }
}

function listKeyOf(numId: string): string {
  return `numbering-${numId}`
}

function ensureList(catalog: NumberingCatalog, listType: string): IListData {
  const existing = catalog.lists[listType]
  if (existing !== undefined) return existing
  const created: IListData = { listType, nestingLevel: [defaultNestingLevel(0)] }
  catalog.lists[listType] = created
  return created
}

function readStoryRelationships(
  parts: ReadonlyMap<string, Uint8Array>,
): ReadonlyMap<string, DocumentRelationships> {
  const relationships = new Map<string, DocumentRelationships>()
  for (const [id, xml] of parts) relationships.set(id, readRelationships(xml))
  return relationships
}

function append(context: ReadContext, text: string): void {
  if (text.length === 0) return
  context.parts.push(text)
  context.length += text.length
}

function lastCharacter(context: ReadContext): string | undefined {
  const last = context.parts[context.parts.length - 1]
  return last === undefined || last.length === 0 ? undefined : last[last.length - 1]
}

function appendParagraphBreak(
  context: ReadContext,
  styleId?: string,
  formatting?: IParagraphStyle,
  bullet?: IBullet,
  synthetic = false,
): void {
  append(context, PARAGRAPH)
  const paragraph: IParagraph = {
    startIndex: context.length - 1,
    paragraphId: context.anchored
      ? `para_${String(context.paragraphs.length + 1)}`
      : `paragraph-${String(context.paragraphs.length + 1)}`,
  }
  if (synthetic) context.syntheticParagraphs.add(paragraph.startIndex)
  if (styleId !== undefined) paragraph.styleId = styleId
  const namedStyle = styleId === undefined ? undefined : NAMED_STYLE_BY_ID[styleId]
  if (namedStyle !== undefined) paragraph.paragraphStyle = { ...formatting, namedStyleType: namedStyle }
  else if (formatting !== undefined) paragraph.paragraphStyle = formatting
  if (bullet !== undefined) paragraph.bullet = bullet
  context.paragraphs.push(paragraph)
}

function appendParagraph(
  context: ReadContext,
  element: XmlElement,
  relationships?: DocumentRelationships,
  identifiers?: StoryPartIds,
  numbering?: NumberingCatalog,
): void {
  const properties = findChild(element, 'pPr')
  const style = properties === undefined ? undefined : findChild(properties, 'pStyle')
  const styleId = style === undefined ? undefined : attribute(style, '@_val')
  const formatting = properties === undefined ? undefined : readParagraphStyle(properties)
  const bullet = properties === undefined || numbering === undefined
    ? undefined
    : readBullet(properties, numbering)
  for (const child of childElements(element)) {
    if (child.tag === 'r') appendRun(context, child)
    if (child.tag === 'hyperlink') appendHyperlink(context, child, relationships)
    if (RAW_ELEMENT_TAGS.includes(child.tag)) appendRawElement(context, child)
  }
  appendParagraphBreak(context, styleId, formatting, bullet)
  if (properties === undefined || relationships === undefined || identifiers === undefined) return
  const section = findChild(properties, 'sectPr')
  if (section === undefined) return
  context.sectionBreaks.push({
    sectionId: `section-${context.sectionBreaks.length + 1}`,
    startIndex: context.length,
    ...readSectionSetup(section, relationships, identifiers),
  })
  append(context, SECTION_BREAK)
}

/**
 * Carry one paragraph-level preserved element verbatim.
 * @param context - Document stream being assembled.
 * @param element - Parsed `pict`, `sdt`, `fldSimple`, or bookmark marker.
 * @throws {OfficeExchangeError} `conversion-failed` when no source XML remains for the element.
 */
function appendRawElement(context: ReadContext, element: XmlElement): void {
  const xml = context.rawElements?.take(element)
  if (xml === undefined) {
    throw new OfficeExchangeError('conversion-failed', `Cannot preserve a ${element.tag} element`)
  }
  context.rawBlocks.push({ startIndex: context.length, xml })
}

/**
 * Append the unit placeholder for a body-level content control.
 *
 * A structured document tag's meaning depends on the content it wraps, so the
 * reader keeps it out of the editable model: the empty paragraph holds the
 * unit slot a neighbouring splice can address, and any edit of that unit is
 * refused as a coupled construct because the rendered paragraph cannot
 * reproduce the `w:sdt` root.
 * @param context - Document stream being assembled.
 */
function appendContentControl(context: ReadContext): void {
  appendParagraphBreak(context)
}

/**
 * Map `w:numPr` into the paragraph's list bullet.
 * @param properties - Paragraph properties containing `w:numPr`.
 * @param numbering - Numbering definitions from the package.
 * @returns The paragraph bullet, or `undefined` for an unnumbered paragraph.
 */function readBullet(properties: XmlElement, numbering: NumberingCatalog): IBullet | undefined {
  const numberingProperties = findChild(properties, 'numPr')
  if (numberingProperties === undefined) return undefined
  const numId = optionalAttribute(findChild(numberingProperties, 'numId'), '@_val')
  if (numId === undefined || numId === '0') return undefined
  const nestingLevel = numberAttribute(findChild(numberingProperties, 'ilvl'), '@_val') ?? 0
  const listType = numbering.listTypeByNumId.get(numId) ?? listKeyOf(numId)
  ensureList(numbering, listType)
  const bullet: IBullet = {
    listType,
    listId: `num-${numId}`,
    nestingLevel,
  }
  const start = numbering.startByNumLevel.get(`${numId}:${String(nestingLevel)}`)
  if (start !== undefined) bullet.startNumber = start
  return bullet
}

/**
 * Append a `w:hyperlink`'s runs inside custom-range tokens and record the range.
 * @param context - Document stream being assembled.
 * @param element - Hyperlink element.
 * @param relationships - Relationships that resolve an external target.
 */
function appendHyperlink(
  context: ReadContext,
  element: XmlElement,
  relationships: DocumentRelationships | undefined,
): void {
  const relationshipId = attribute(element, '@_id')
  const anchor = attribute(element, '@_anchor')
  const relationship = relationshipId === undefined ? undefined : relationships?.get(relationshipId)
  const url = relationship !== undefined
    && relationship.type.endsWith('/hyperlink')
    && relationship.target.length > 0
    ? relationship.target
    : undefined
  if (url === undefined && anchor === undefined) {
    for (const child of childElements(element)) {
      if (child.tag === 'r') appendRun(context, child)
    }
    return
  }
  const start = context.length
  append(context, CUSTOM_RANGE_START)
  for (const child of childElements(element)) {
    if (child.tag === 'r') appendRun(context, child)
  }
  append(context, CUSTOM_RANGE_END)
  const properties: IHyperlink = {}
  if (url !== undefined) properties.url = url
  if (anchor !== undefined) properties.bookmarkId = anchor
  context.customRanges.push({
    startIndex: start,
    endIndex: context.length - 1,
    rangeId: `hyperlink-${String(context.customRanges.length + 1)}`,
    rangeType: CustomRangeType.HYPERLINK,
    properties,
  })
}

/**
 * Map one `w:sectPr` element into section metadata.
 * @param element - Section properties element.
 * @param relationships - Document relationships resolving header and footer references.
 * @param identifiers - Header and footer part ids present in the package.
 * @returns Page size, orientation, margins, columns, break type, and header/footer references.
 */
function readSectionSetup(
  element: XmlElement,
  relationships: DocumentRelationships,
  identifiers: StoryPartIds,
): SectionSetup {
  const section: SectionSetup = {}
  readPageSize(findChild(element, 'pgSz'), section)
  readPageMargins(findChild(element, 'pgMar'), section)
  const columns = findChild(element, 'cols')
  if (columns !== undefined) readColumns(columns, section)
  const type = optionalAttribute(findChild(element, 'type'), '@_val')
  const sectionType = type === undefined ? undefined : lookup(SECTION_TYPES, type)
  if (sectionType !== undefined) section.sectionType = sectionType
  for (const child of childElements(element)) {
    if (child.tag !== 'headerReference' && child.tag !== 'footerReference') continue
    readStoryReference(child, relationships, identifiers, section)
  }
  if (isEnabled(element, 'titlePg') === true) section.useFirstPageHeaderFooter = BOOLEAN_TRUE
  return section
}

function readPageSize(element: XmlElement | undefined, section: SectionSetup): void {
  if (element === undefined) return
  const width = numberAttribute(element, '@_w')
  const height = numberAttribute(element, '@_h')
  if (width === undefined || height === undefined) return
  section.pageSize = { width: width / TWIPS_PER_PIXEL, height: height / TWIPS_PER_PIXEL }
  section.pageOrient = pageOrientationOf(attribute(element, '@_orient'), width, height)
}

function pageOrientationOf(orientation: string | undefined, width: number, height: number): PageOrientType {
  if (orientation === 'landscape') return PageOrientType.LANDSCAPE
  if (orientation === 'portrait') return PageOrientType.PORTRAIT
  return width > height ? PageOrientType.LANDSCAPE : PageOrientType.PORTRAIT
}

function readPageMargins(element: XmlElement | undefined, section: SectionSetup): void {
  if (element === undefined) return
  const top = numberAttribute(element, '@_top')
  const bottom = numberAttribute(element, '@_bottom')
  const left = numberAttribute(element, '@_left')
  const right = numberAttribute(element, '@_right')
  const header = numberAttribute(element, '@_header')
  const footer = numberAttribute(element, '@_footer')
  if (top !== undefined) section.marginTop = top / TWIPS_PER_PIXEL
  if (bottom !== undefined) section.marginBottom = bottom / TWIPS_PER_PIXEL
  if (left !== undefined) section.marginLeft = left / TWIPS_PER_PIXEL
  if (right !== undefined) section.marginRight = right / TWIPS_PER_PIXEL
  if (header !== undefined) section.marginHeader = header / TWIPS_PER_PIXEL
  if (footer !== undefined) section.marginFooter = footer / TWIPS_PER_PIXEL
}

function readColumns(element: XmlElement, section: SectionSetup): void {
  const separator = attribute(element, '@_sep')
  section.columnSeparatorType = separator === undefined || separator === '0' || separator === 'false' || separator === 'off'
    ? ColumnSeparatorType.NONE
    : ColumnSeparatorType.BETWEEN_EACH_COLUMN
  const columns: ISectionColumnProperties[] = []
  for (const child of childElements(element)) {
    if (child.tag !== 'col') continue
    const width = numberAttribute(child, '@_w')
    if (width === undefined) continue
    columns.push({
      width: width / TWIPS_PER_PIXEL,
      paddingEnd: (numberAttribute(child, '@_space') ?? 0) / TWIPS_PER_PIXEL,
    })
  }
  if (columns.length > 0) {
    section.columnProperties = columns
    return
  }
  const count = numberAttribute(element, '@_num')
  if (count === undefined || count <= 1) return
  const spacing = numberAttribute(element, '@_space') ?? DEFAULT_COLUMN_SPACING
  const widths = equalColumnWidths(section, count, spacing / TWIPS_PER_PIXEL)
  if (widths !== undefined) section.columnProperties = widths
}

function equalColumnWidths(
  section: SectionSetup,
  count: number,
  spacing: number,
): ISectionColumnProperties[] | undefined {
  const pageSize = section.pageSize
  if (pageSize?.width === undefined) return undefined
  const available = pageSize.width - (section.marginLeft ?? DEFAULT_PAGE_MARGIN)
    - (section.marginRight ?? DEFAULT_PAGE_MARGIN)
  const width = (available - (count - 1) * spacing) / count
  if (!Number.isFinite(width) || width <= 0) return undefined
  const columns: ISectionColumnProperties[] = []
  for (let index = 0; index < count; index += 1) {
    columns.push({ width, paddingEnd: index === count - 1 ? 0 : spacing })
  }
  return columns
}

function readStoryReference(
  element: XmlElement,
  relationships: DocumentRelationships,
  identifiers: StoryPartIds,
  section: SectionSetup,
): void {
  const relationshipId = attribute(element, '@_id')
  if (relationshipId === undefined) return
  const relationship = relationships.get(relationshipId)
  if (relationship === undefined || relationship.target.length === 0) return
  const id = partIdOf(relationship.target)
  const header = element.tag === 'headerReference'
  if (header ? relationship.type.endsWith('/header') : relationship.type.endsWith('/footer')) {
    if (header ? !identifiers.headers.has(id) : !identifiers.footers.has(id)) return
    section[referenceKeyOf(element.tag, variantOf(attribute(element, '@_type')))] = id
  }
}

function referenceKeyOf(tag: string, variant: StoryVariant): StoryReferenceKey {
  return tag === 'headerReference' ? HEADER_REFERENCE_KEYS[variant] : FOOTER_REFERENCE_KEYS[variant]
}

function variantOf(value: string | undefined): StoryVariant {
  if (value === 'first') return 'first'
  if (value === 'even') return 'even'
  return 'default'
}

function appendRun(context: ReadContext, element: XmlElement): void {
  if (containsPreservedContent(element)) {
    const xml = context.rawRuns?.take(element)
    if (xml === undefined) {
      throw new OfficeExchangeError('conversion-failed', 'Cannot preserve a drawing, picture, or note reference run')
    }
    context.rawBlocks.push({ startIndex: context.length, xml })
    return
  }
  const properties = findChild(element, 'rPr')
  const style = properties === undefined ? undefined : readTextStyle(properties)
  const start = context.length
  let text = ''
  for (const child of childElements(element)) {
    if (child.tag === 't') text += textContent(child)
    if (child.tag === 'tab') text += TAB
  }
  append(context, text)
  if (style !== undefined && text.length > 0) {
    context.textRuns.push({ st: start, ed: context.length, ts: style })
  }
}

/** Run children whose XML the writer re-emits instead of mapping into the stream. */
const PRESERVED_RUN_TAGS: ReadonlySet<string> = new Set([
  'drawing',
  'pict',
  'object',
  'fldChar',
  'instrText',
  'footnoteReference',
  'endnoteReference',
])

/** Source-run scan derived from {@link PRESERVED_RUN_TAGS}. */
const PRESERVED_RUN_PATTERN = new RegExp(
  `<w:(?:${[...PRESERVED_RUN_TAGS].join('|')})(?=[\\s/>])`,
  'u',
)

/**
 * Report whether a run carries content preserved only as raw XML.
 * @param element - Parsed run.
 * @returns `true` for an embedded drawing, picture, field marker, or note reference anywhere inside the run.
 */
function containsPreservedContent(element: XmlElement): boolean {
  for (const child of childElements(element)) {
    if (PRESERVED_RUN_TAGS.has(child.tag)) return true
    if (containsPreservedContent(child)) return true
  }
  return false
}

/**
 * Scan one part's source XML for runs that carry preserved content.
 * @param xml - Decoded part XML.
 * @returns A source whose runs are taken once, matched by identity first and source order second.
 */
function readRawRunCatalog(xml: string): RawRunSource {
  const records: RawXmlRecord[] = []
  for (const match of xml.matchAll(/<w:r(?=[\s/>])[^>]*>/gu)) {
    if (match[0].endsWith('/>')) continue
    const start = match.index
    const end = runEndOf(xml, start)
    /* v8 ignore next 2 -- A validated part ends every run it opens. */
    if (end === undefined) continue
    const run = xml.slice(start, end)
    if (!containsPreservedMarkup(run)) continue
    records.push({ xml: run, key: relationshipKeyOfXml(run), taken: false })
  }
  return rawXmlSourceOf(records, relationshipKeyOfElement)
}

/**
 * Scan one part's source XML for paragraph-level elements the writer re-emits.
 *
 * The catalog holds the outermost occurrence of each preserved element, so an
 * element nested inside another is carried with its container rather than
 * twice. Parsed children consume records by relationship identity first and
 * source order second.
 * @param xml - Decoded part XML.
 * @returns A source whose preserved elements are taken once.
 */
function readRawElementCatalog(xml: string): RawElementSource {
  const records: RawXmlRecord[] = []
  let coveredTo = -1
  for (const match of xml.matchAll(RAW_ELEMENT_PATTERN)) {
    const start = match.index
    if (start < coveredTo) continue
    const name = match[1]
    /* v8 ignore next -- The raw-element pattern always captures a name. */
    if (name === undefined) continue
    const end = match[0].endsWith('/>') ? start + match[0].length : elementEndOf(xml, start, name)
    /* v8 ignore next 2 -- A validated part ends every preserved element it opens. */
    if (end === undefined) continue
    coveredTo = end
    const element = xml.slice(start, end)
    records.push({ xml: element, key: relationshipKeyOfXml(element), taken: false })
  }
  return rawXmlSourceOf(records, relationshipKeyOfElement)
}

/**
 * Build a source that yields each captured element at most once.
 * @param records - Captured elements in source order.
 * @param keyOf - Identity used to match a parsed element before falling back to source order.
 * @returns A source whose `take` consumes the matching record or the next unconsumed one.
 */
function rawXmlSourceOf(records: readonly RawXmlRecord[], keyOf: (element: XmlElement) => string): RawXmlSource {
  const byKey = new Map<string, RawXmlRecord[]>()
  for (const record of records) {
    const queue = byKey.get(record.key)
    if (queue === undefined) byKey.set(record.key, [record])
    else queue.push(record)
  }
  let cursor = 0
  return {
    take(element) {
      const key = keyOf(element)
      for (const record of byKey.get(key) ?? []) {
        if (record.taken) continue
        record.taken = true
        return record.xml
      }
      while (cursor < records.length && records[cursor]?.taken === true) cursor += 1
      const fallback = records[cursor]
      if (fallback === undefined) return undefined
      fallback.taken = true
      cursor += 1
      return fallback.xml
    },
  }
}

/**
 * Find the end of the `w:` element named `name` that starts at `start`.
 * @param xml - Part XML.
 * @param start - Offset of the element's opening tag.
 * @param name - Local element name without its prefix.
 * @returns Exclusive offset after its closing tag, or `undefined` when it is unterminated.
 */
function elementEndOf(xml: string, start: number, name: string): number | undefined {
  const pattern = new RegExp(`<w:${name}(?=[\\s/>])[^>]*>|</w:${name}>`, 'gu')
  pattern.lastIndex = start
  let depth = 0
  for (let match = pattern.exec(xml); match !== null; match = pattern.exec(xml)) {
    if (match[0] === `</w:${name}>`) {
      depth -= 1
      if (depth === 0) return pattern.lastIndex
    } else if (!match[0].endsWith('/>')) {
      depth += 1
    }
  }
  /* v8 ignore next -- A validated part ends every element it opens. */
  return undefined
}

/**
 * Find the end of the `w:r` element that starts at `start`.
 * @param xml - Part XML.
 * @param start - Offset of the run's opening tag.
 * @returns Exclusive offset after its closing tag, or `undefined` when it is unterminated.
 */
function runEndOf(xml: string, start: number): number | undefined {
  const pattern = /<w:r(?=[\s/>])[^>]*>|<\/w:r>/gu
  pattern.lastIndex = start
  let depth = 0
  for (let match = pattern.exec(xml); match !== null; match = pattern.exec(xml)) {
    if (match[0] === '</w:r>') {
      depth -= 1
      if (depth === 0) return pattern.lastIndex
    } else if (!match[0].endsWith('/>')) {
      depth += 1
    }
  }
  /* v8 ignore next -- A validated part ends every element it opens. */
  return undefined
}

/**
 * Report whether source run XML carries preserved content.
 * @param xml - One source run.
 * @returns `true` when it contains a drawing, picture, object, field marker, or note reference.
 */
function containsPreservedMarkup(xml: string): boolean {
  return PRESERVED_RUN_PATTERN.test(xml)
}

/**
 * Collect relationship and note identity values from source run XML.
 * @param xml - One source run.
 * @returns Sorted attribute values joined as one match key.
 */
function relationshipKeyOfXml(xml: string): string {
  const values: string[] = []
  for (const match of xml.matchAll(/\b(?:embed|link|pict|id|dm|lo|qs|cs)="([^"]*)"/gu)) {
    const value = match[1]
    /* v8 ignore next -- The capture group is present whenever the pattern matches. */
    if (value !== undefined) values.push(value)
  }
  return values.sort().join('\u0000')
}

/**
 * Collect relationship and note identity values from a parsed run.
 * @param element - Parsed run.
 * @returns Sorted attribute values joined as one match key.
 */
function relationshipKeyOfElement(element: XmlElement): string {
  const values: string[] = []
  collectRelationshipValues(element, values)
  return values.sort().join('\u0000')
}

function collectRelationshipValues(element: XmlElement, values: string[]): void {
  for (const name of ['@_embed', '@_link', '@_pict', '@_id', '@_dm', '@_lo', '@_qs', '@_cs']) {
    const value = attribute(element, name)
    if (value !== undefined && value.length > 0) values.push(value)
  }
  for (const child of childElements(element)) collectRelationshipValues(child, values)
}

function readTextStyle(properties: XmlElement): ITextStyle | undefined {
  const style: ITextStyle = {}
  if (isEnabled(properties, 'b') === true) style.bl = BOOLEAN_TRUE
  if (isEnabled(properties, 'i') === true) style.it = BOOLEAN_TRUE
  if (isEnabled(properties, 'caps') === true) style.caps = true
  if (isEnabled(properties, 'smallCaps') === true) style.smallCaps = true
  const strike = readStrike(properties)
  if (strike !== undefined) style.st = strike
  const fonts = findChild(properties, 'rFonts')
  if (fonts !== undefined) {
    const latin = attribute(fonts, '@_ascii') ?? attribute(fonts, '@_hAnsi') ?? attribute(fonts, '@_cs')
    if (latin !== undefined) style.ff = latin
    const eastAsia = attribute(fonts, '@_eastAsia')
    if (eastAsia !== undefined) style.eastAsiaFontFamily = eastAsia
  }
  const size = readFontSize(properties)
  if (size !== undefined) style.fs = size
  const color = rgbOf(findChild(properties, 'color'))
  if (color !== undefined) style.cl = { rgb: color }
  const background = readHighlight(properties) ?? shadingColor(findChild(properties, 'shd'))
  if (background !== undefined) style.bg = { rgb: background }
  const underline = readUnderline(properties)
  if (underline !== undefined) style.ul = underline
  const vertical = readVerticalAlign(properties)
  if (vertical !== undefined) style.va = vertical
  const border = readRunBorder(properties)
  if (border !== undefined) style.bd = border
  return Object.keys(style).length === 0 ? undefined : style
}

function readParagraphStyle(properties: XmlElement): IParagraphStyle | undefined {
  const style: IParagraphStyle = {}
  const alignment = optionalAttribute(findChild(properties, 'jc'), '@_val')
  const horizontal = alignment === undefined ? undefined : lookup(HORIZONTAL_ALIGNMENTS, alignment)
  if (horizontal !== undefined) style.horizontalAlign = horizontal
  const spacing = findChild(properties, 'spacing')
  if (spacing !== undefined) readSpacing(spacing, style)
  const indent = findChild(properties, 'ind')
  if (indent !== undefined) readIndent(indent, style)
  const tabs = findChild(properties, 'tabs')
  if (tabs !== undefined) readTabStops(tabs, style)
  const borders = findChild(properties, 'pBdr')
  if (borders !== undefined) readParagraphBorders(borders, style)
  const outline = numberAttribute(findChild(properties, 'outlineLvl'), '@_val')
  if (outline !== undefined) style.outlineLevel = outline
  const shading = shadingColor(findChild(properties, 'shd'))
  if (shading !== undefined) style.shading = { backgroundColor: { rgb: shading } }
  return Object.keys(style).length === 0 ? undefined : style
}

function readSpacing(element: XmlElement, style: IParagraphStyle): void {
  const before = numberAttribute(element, '@_before')
  if (before !== undefined) style.spaceAbove = { v: before / TWIPS_PER_PIXEL }
  const after = numberAttribute(element, '@_after')
  if (after !== undefined) style.spaceBelow = { v: after / TWIPS_PER_PIXEL }
  const line = numberAttribute(element, '@_line')
  if (line === undefined) return
  const rule = attribute(element, '@_lineRule')
  if (rule === undefined || rule === 'auto') {
    style.lineSpacing = line / AUTO_LINE_SPACING_UNITS
    style.spacingRule = SpacingRule.AUTO
    return
  }
  style.lineSpacing = line / TWIPS_PER_PIXEL
  style.spacingRule = rule === 'atLeast' ? SpacingRule.AT_LEAST : SpacingRule.EXACT
}

function readIndent(element: XmlElement, style: IParagraphStyle): void {
  const start = numberAttribute(element, '@_start') ?? numberAttribute(element, '@_left')
  if (start !== undefined) style.indentStart = { v: start / TWIPS_PER_PIXEL }
  const end = numberAttribute(element, '@_end') ?? numberAttribute(element, '@_right')
  if (end !== undefined) style.indentEnd = { v: end / TWIPS_PER_PIXEL }
  const firstLine = numberAttribute(element, '@_firstLine')
  if (firstLine !== undefined) style.indentFirstLine = { v: firstLine / TWIPS_PER_PIXEL }
  const hanging = numberAttribute(element, '@_hanging')
  if (hanging !== undefined) style.hanging = { v: hanging / TWIPS_PER_PIXEL }
}

function readTabStops(element: XmlElement, style: IParagraphStyle): void {
  const stops: ITabStop[] = []
  for (const child of childElements(element)) {
    if (child.tag !== 'tab') continue
    const stop = readTabStop(child)
    if (stop !== undefined) stops.push(stop)
  }
  if (stops.length > 0) style.tabStops = stops
}

function readTabStop(element: XmlElement): ITabStop | undefined {
  const position = numberAttribute(element, '@_pos')
  if (position === undefined) return undefined
  const value = attribute(element, '@_val')
  const stop: ITabStop = {
    offset: position / TWIPS_PER_PIXEL,
    alignment: (value === undefined ? undefined : lookup(TAB_STOP_ALIGNMENTS, value)) ?? TabStopAlignment.START,
  }
  if (value === 'clear') stop.clear = true
  const leaderName = attribute(element, '@_leader')
  const leader = leaderName === undefined ? undefined : lookup(TAB_STOP_LEADERS, leaderName)
  if (leader !== undefined) stop.leader = leader
  return stop
}

function readParagraphBorders(element: XmlElement, style: IParagraphStyle): void {
  const top = readParagraphBorder(findChild(element, 'top'))
  if (top !== undefined) style.borderTop = top
  const bottom = readParagraphBorder(findChild(element, 'bottom'))
  if (bottom !== undefined) style.borderBottom = bottom
  const left = readParagraphBorder(findChild(element, 'left'))
  if (left !== undefined) style.borderLeft = left
  const right = readParagraphBorder(findChild(element, 'right'))
  if (right !== undefined) style.borderRight = right
  const between = readParagraphBorder(findChild(element, 'between'))
  if (between !== undefined) style.borderBetween = between
}

function readParagraphBorder(element: XmlElement | undefined): IParagraphBorder | undefined {
  if (element === undefined) return undefined
  const value = attribute(element, '@_val')
  if (value === undefined || value === 'none') return undefined
  const size = positiveNumber(attribute(element, '@_sz'))
  const space = numberAttribute(element, '@_space')
  const color = hexColor(attribute(element, '@_color')) ?? DEFAULT_BORDER_COLOR
  return {
    color: { rgb: `#${color}` },
    width: (size ?? DEFAULT_BORDER_SIZE) / BORDER_UNITS_PER_PIXEL,
    dashStyle: lookup(BORDER_DASH_STYLES, value) ?? DashStyleType.SOLID,
    padding: (space ?? 0) * PIXELS_PER_POINT,
  }
}

function readUnderline(properties: XmlElement): ITextDecoration | undefined {
  const underline = findChild(properties, 'u')
  if (underline === undefined) return undefined
  const value = attribute(underline, '@_val')
  if (value === 'none') return undefined
  const decoration = value === undefined ? undefined : lookup(UNDERLINE_VALUES, value)
  const style: ITextDecoration = {
    s: BOOLEAN_TRUE,
    t: decoration ?? TEXT_DECORATION_SINGLE,
  }
  const color = hexColor(attribute(underline, '@_color'))
  if (color !== undefined) style.cl = { rgb: `#${color}` }
  return style
}

function readStrike(properties: XmlElement): ITextDecoration | undefined {
  if (isEnabled(properties, 'dstrike') === true) {
    return { s: BOOLEAN_TRUE, t: TEXT_DECORATION_DOUBLE }
  }
  if (isEnabled(properties, 'strike') === true) {
    return { s: BOOLEAN_TRUE, t: TEXT_DECORATION_SINGLE }
  }
  return undefined
}

function readFontSize(properties: XmlElement): number | undefined {
  const size = findChild(properties, 'sz') ?? findChild(properties, 'szCs')
  if (size === undefined) return undefined
  const halfPoints = positiveNumber(attribute(size, '@_val'))
  return halfPoints === undefined ? undefined : halfPoints / HALF_POINTS_PER_POINT
}

function readHighlight(properties: XmlElement): string | undefined {
  const highlight = findChild(properties, 'highlight')
  if (highlight === undefined) return undefined
  const value = attribute(highlight, '@_val')
  if (value === undefined || value === 'none') return undefined
  const color = HIGHLIGHT_COLORS[value]
  return color === undefined ? undefined : `#${color}`
}

function readVerticalAlign(properties: XmlElement): BaselineOffset | undefined {
  const value = optionalAttribute(findChild(properties, 'vertAlign'), '@_val')
  if (value === 'superscript') return BaselineOffset.SUPERSCRIPT
  if (value === 'subscript') return BaselineOffset.SUBSCRIPT
  if (value === 'baseline') return BaselineOffset.NORMAL
  return undefined
}

function readRunBorder(properties: XmlElement): IBorderData | undefined {
  const border = findChild(properties, 'bdr')
  if (border === undefined) return undefined
  const side = readBorderStyle(border)
  return side === undefined ? undefined : { t: side, r: side, b: side, l: side }
}

function readBorderStyle(element: XmlElement): IBorderStyleData | undefined {
  const value = attribute(element, '@_val')
  if (value === undefined || value === 'none') return undefined
  const color = hexColor(attribute(element, '@_color')) ?? DEFAULT_BORDER_COLOR
  return {
    s: lookup(BORDER_WEIGHTS, value) ?? BorderStyleTypes.THIN,
    cl: { rgb: `#${color}` },
  }
}

function isEnabled(properties: XmlElement, tag: string): boolean | undefined {
  const element = findChild(properties, tag)
  if (element === undefined) return undefined
  const value = attribute(element, '@_val')
  return value === undefined || (value !== '0' && value !== 'false' && value !== 'off')
}

/**
 * Read the document defaults declared by `word/styles.xml`.
 * @param stylesXml - Decoded styles part, when the package contains one.
 * @returns Document-level text and paragraph defaults; empty when the part is absent.
 * @throws {OfficeExchangeError} `conversion-failed` when the styles XML is not well formed.
 */
function readDocumentStyle(stylesXml: Uint8Array | undefined): IDocumentStyle {
  if (stylesXml === undefined) return {}
  const root = parseStylesDocument(stylesXml)
  if (root === undefined) return {}
  const defaults = findChild(root, 'docDefaults')
  const defaultText = defaults === undefined ? undefined : defaultsTextStyle(defaults)
  const defaultParagraph = defaults === undefined ? undefined : defaultsParagraphStyle(defaults)
  const normal = findNormalStyle(root)
  const textStyle = mergeStyles(
    defaultText,
    normal === undefined ? undefined : readElementProperties(normal, 'rPr', readTextStyle),
  )
  const paragraphStyle = mergeStyles(
    defaultParagraph,
    normal === undefined ? undefined : readElementProperties(normal, 'pPr', readParagraphStyle),
  )
  const documentStyle: IDocumentStyle = {}
  if (textStyle !== undefined) documentStyle.textStyle = textStyle
  if (paragraphStyle !== undefined) documentStyle.defaultParagraphStyle = paragraphStyle
  return documentStyle
}

function defaultsTextStyle(defaults: XmlElement): ITextStyle | undefined {
  const container = findChild(defaults, 'rPrDefault')
  return container === undefined ? undefined : readElementProperties(container, 'rPr', readTextStyle)
}

function defaultsParagraphStyle(defaults: XmlElement): IParagraphStyle | undefined {
  const container = findChild(defaults, 'pPrDefault')
  return container === undefined ? undefined : readElementProperties(container, 'pPr', readParagraphStyle)
}

function readElementProperties<Properties>(
  element: XmlElement,
  tag: string,
  read: (properties: XmlElement) => Properties | undefined,
): Properties | undefined {
  const properties = findChild(element, tag)
  return properties === undefined ? undefined : read(properties)
}

function mergeStyles<Value extends object>(
  base: Value | undefined,
  override: Value | undefined,
): Value | undefined {
  if (base === undefined) return override
  if (override === undefined) return base
  return { ...base, ...override }
}

function findNormalStyle(root: XmlElement): XmlElement | undefined {
  for (const child of childElements(root)) {
    if (child.tag !== 'style') continue
    if (attribute(child, '@_type') !== 'paragraph') continue
    if (attribute(child, '@_default') === '1' || attribute(child, '@_styleId') === 'Normal') return child
  }
  return undefined
}

function parseStylesDocument(stylesXml: Uint8Array): XmlElement | undefined {
  const xml = strFromU8(stylesXml)
  let parsed: readonly ParserNode[]
  try {
    SyntaxValidator.validate(xml)
    parsed = new XMLParser(XML_PARSER_OPTIONS).parse(xml) as readonly ParserNode[]
  } catch (error) {
    throw new OfficeExchangeError('conversion-failed', `Cannot read styles XML: ${messageOf(error)}`)
  }
  return parsed.map(toElement).find(element => element?.tag === 'styles')
}

function appendTable(
  context: ReadContext,
  element: XmlElement,
  relationships?: DocumentRelationships,
  numbering?: NumberingCatalog,
): void {
  if (lastCharacter(context) !== PARAGRAPH) appendParagraphBreak(context, undefined, undefined, undefined, true)
  context.tableSerial += 1
  const tableId = `table-${String(context.tableSerial)}`
  const startIndex = context.length
  append(context, TABLE_START)
  const rows: ITableRow[] = []
  const gridDescriptors: CellGridDescriptor[][] = []
  const gridWidths = readGridWidths(element)
  let columnCount = 1
  for (const row of childElements(element)) {
    if (row.tag !== 'tr') continue
    append(context, TABLE_ROW_START)
    const cells: ITableCell[] = []
    const rowDescriptors: CellGridDescriptor[] = []
    let gridColumn = 0
    for (const cell of childElements(row)) {
      if (cell.tag !== 'tc') continue
      countCell(context)
      append(context, TABLE_CELL_START)
      for (const block of childElements(cell)) {
        if (block.tag === 'p') appendParagraph(context, block, relationships, undefined, numbering)
        if (block.tag === 'tbl') appendTable(context, block, relationships, numbering)
      }
      if (lastCharacter(context) !== PARAGRAPH) appendParagraphBreak(context, undefined, undefined, undefined, true)
      context.sectionBreaks.push({
        sectionId: `section-${context.sectionBreaks.length + 1}`,
        startIndex: context.length,
      })
      append(context, `${SECTION_BREAK}${TABLE_CELL_END}`)
      const properties = findChild(cell, 'tcPr')
      const columnSpan = readColumnSpan(properties)
      const tableCell = readTableCell(properties)
      if (columnSpan > 1) tableCell.columnSpan = columnSpan
      rowDescriptors.push({ gridStart: gridColumn, columnSpan, vMerge: readVerticalMerge(properties) })
      gridColumn += columnSpan
      cells.push(tableCell)
    }
    append(context, TABLE_ROW_END)
    columnCount = Math.max(columnCount, gridColumn)
    rows.push({ tableCells: cells, ...readRowProperties(row, rows.length === 0) })
    gridDescriptors.push(rowDescriptors)
  }
  append(context, TABLE_END)
  applyVerticalMerges(rows, gridDescriptors)
  context.tables.push({ startIndex, endIndex: context.length, tableId })
  context.tableSource[tableId] = createTable(tableId, rows, columnCount, gridWidths, readTableProperties(element))
}

function countCell(context: ReadContext): void {
  context.cells += 1
  if (context.cells > context.bounds.maxCells) {
    throw new OfficeExchangeError('too-large', `Document has more than ${context.bounds.maxCells} table cells`)
  }
}

/** One cell's table-grid coverage and vertical-merge role. */
interface CellGridDescriptor {
  /** Index of the first grid column the cell occupies. */
  readonly gridStart: number
  /** Number of grid columns the cell spans. */
  readonly columnSpan: number
  /** Vertical-merge role declared by `w:vMerge`. */
  readonly vMerge: 'restart' | 'continue' | undefined
}

/** Table-level layout read from `w:tblPr`. */
interface TableSetup {
  /** Preferred table width from `w:tblW`. */
  size?: IWidthInTableSize
  /** Table alignment from `w:jc`. */
  align?: TableAlignmentType
  /** Table indentation from `w:tblInd`. */
  indent?: INumberUnit
  /** Fixed or automatic fit from `w:tblLayout`. */
  layout?: TableLayoutType
}

function readColumnSpan(properties: XmlElement | undefined): number {
  if (properties === undefined) return 1
  const span = numberAttribute(findChild(properties, 'gridSpan'), '@_val')
  return span === undefined || span <= 1 ? 1 : Math.floor(span)
}

function readVerticalMerge(properties: XmlElement | undefined): 'restart' | 'continue' | undefined {
  if (properties === undefined) return undefined
  const merge = findChild(properties, 'vMerge')
  if (merge === undefined) return undefined
  return attribute(merge, '@_val') === 'restart' ? 'restart' : 'continue'
}

/**
 * Resolve `w:vMerge` chains into model spans: the restart cell carries the
 * row count and each covered continuation cell is marked with `rowSpan` zero.
 * @param rows - Rows whose cells receive their spans.
 * @param descriptors - Grid coverage and merge role of every cell, in row order.
 */
function applyVerticalMerges(
  rows: readonly ITableRow[],
  descriptors: readonly (readonly CellGridDescriptor[])[],
): void {
  for (const [rowIndex, rowDescriptors] of descriptors.entries()) {
    for (const [cellIndex, descriptor] of rowDescriptors.entries()) {
      if (descriptor.vMerge !== 'restart') continue
      let span = 1
      for (let coveredRow = rowIndex + 1; coveredRow < descriptors.length; coveredRow += 1) {
        /* v8 ignore next -- descriptors and rows are built together for the same table. */
        const candidates = descriptors[coveredRow] ?? []
        const coveredIndex = candidates.findIndex(candidate => (
          candidate.vMerge === 'continue' && candidate.gridStart === descriptor.gridStart
        ))
        if (coveredIndex < 0) break
        const coveredCell = rows[coveredRow]?.tableCells[coveredIndex]
        if (coveredCell !== undefined) coveredCell.rowSpan = 0
        span += 1
      }
      if (span > 1) {
        const startCell = rows[rowIndex]?.tableCells[cellIndex]
        /* v8 ignore next -- Every vertical-merge chain starts at a descriptor row. */
        if (startCell !== undefined) startCell.rowSpan = span
      }
    }
  }
}

function readTableCell(properties: XmlElement | undefined): ITableCell {
  const cell: ITableCell = { margin: readCellMargin(properties) }
  const width = readCellWidth(properties)
  if (width !== undefined) cell.size = width
  if (properties === undefined) return cell
  const background = shadingColor(findChild(properties, 'shd'))
  if (background !== undefined) cell.backgroundColor = { rgb: background }
  const vertical = readCellVerticalAlign(properties)
  if (vertical !== undefined) cell.vAlign = vertical
  const borders = findChild(properties, 'tcBorders')
  if (borders !== undefined) readCellBorders(borders, cell)
  return cell
}

function readCellWidth(properties: XmlElement | undefined): IWidthInTableSize | undefined {
  if (properties === undefined) return undefined
  const width = findChild(properties, 'tcW')
  if (width === undefined) return undefined
  const twips = positiveNumber(attribute(width, '@_w'))
  if (twips === undefined) return undefined
  return { type: TABLE_SIZE_SPECIFIED, width: { v: twips / TWIPS_PER_PIXEL } }
}

function readCellBorders(element: XmlElement, cell: ITableCell): void {
  const top = readCellBorder(findChild(element, 'top') ?? findChild(element, 'insideH'))
  if (top !== undefined) cell.borderTop = top
  const bottom = readCellBorder(findChild(element, 'bottom') ?? findChild(element, 'insideH'))
  if (bottom !== undefined) cell.borderBottom = bottom
  const start = readCellBorder(
    findChild(element, 'start') ?? findChild(element, 'left') ?? findChild(element, 'insideV'),
  )
  if (start !== undefined) cell.borderLeft = start
  const end = readCellBorder(
    findChild(element, 'end') ?? findChild(element, 'right') ?? findChild(element, 'insideV'),
  )
  if (end !== undefined) cell.borderRight = end
}

function readCellBorder(element: XmlElement | undefined): ITableCellBorder | undefined {
  if (element === undefined) return undefined
  const value = attribute(element, '@_val')
  if (value === undefined || value === 'none' || value === 'nil') return undefined
  const size = positiveNumber(attribute(element, '@_sz'))
  const color = hexColor(attribute(element, '@_color')) ?? DEFAULT_BORDER_COLOR
  return {
    color: { rgb: `#${color}` },
    width: { v: (size ?? DEFAULT_BORDER_SIZE) / BORDER_UNITS_PER_PIXEL },
    dashStyle: lookup(BORDER_DASH_STYLES, value) ?? DashStyleType.SOLID,
  }
}

function readCellMargin(properties: XmlElement | undefined): ITableCellMargin {
  const defaults = cellMargin()
  if (properties === undefined) return defaults
  const margins = findChild(properties, 'tcMar')
  if (margins === undefined) return defaults
  return {
    start: readMarginSide(margins, ['start', 'left']) ?? defaults.start,
    end: readMarginSide(margins, ['end', 'right']) ?? defaults.end,
    top: readMarginSide(margins, ['top']) ?? defaults.top,
    bottom: readMarginSide(margins, ['bottom']) ?? defaults.bottom,
  }
}

function readMarginSide(element: XmlElement, tags: readonly string[]): INumberUnit | undefined {
  for (const tag of tags) {
    const twips = numberAttribute(findChild(element, tag), '@_w')
    if (twips !== undefined) return { v: twips / TWIPS_PER_PIXEL }
  }
  return undefined
}

function readCellVerticalAlign(properties: XmlElement): VerticalAlignmentType | undefined {
  const value = optionalAttribute(findChild(properties, 'vAlign'), '@_val')
  if (value === 'top') return VerticalAlignmentType.TOP
  if (value === 'center') return VerticalAlignmentType.CENTER
  if (value === 'bottom') return VerticalAlignmentType.BOTTOM
  return undefined
}

function readRowProperties(
  element: XmlElement,
  first: boolean,
): Pick<ITableRow, 'trHeight' | 'cantSplit' | 'isFirstRow' | 'repeatHeaderRow'> {
  const properties = findChild(element, 'trPr')
  const row: Pick<ITableRow, 'trHeight' | 'cantSplit' | 'isFirstRow' | 'repeatHeaderRow'> = {
    trHeight: readRowHeight(properties),
  }
  if (first) row.isFirstRow = BOOLEAN_TRUE
  if (properties === undefined) return row
  if (isEnabled(properties, 'cantSplit') === true) row.cantSplit = BOOLEAN_TRUE
  if (isEnabled(properties, 'tblHeader') === true) row.repeatHeaderRow = BOOLEAN_TRUE
  return row
}

function readRowHeight(properties: XmlElement | undefined): ITableRowSize {
  const height = properties === undefined ? undefined : findChild(properties, 'trHeight')
  const twips = numberAttribute(height, '@_val')
  if (twips === undefined) return { val: { v: DEFAULT_ROW_HEIGHT }, hRule: TABLE_ROW_HEIGHT_AUTO }
  /* v8 ignore next -- twips is read from height, so a defined twips implies a defined height. */
  const rule = height === undefined ? undefined : attribute(height, '@_hRule')
  const hRule = rule === 'exact'
    ? TABLE_ROW_HEIGHT_EXACT
    : rule === 'atLeast' ? TABLE_ROW_HEIGHT_AT_LEAST : TABLE_ROW_HEIGHT_AUTO
  return { val: { v: twips / TWIPS_PER_PIXEL }, hRule }
}

function readTableProperties(element: XmlElement): TableSetup {
  const properties = findChild(element, 'tblPr')
  if (properties === undefined) return {}
  const setup: TableSetup = {}
  const size = readTableWidth(properties)
  if (size !== undefined) setup.size = size
  const align = readTableAlignment(properties)
  if (align !== undefined) setup.align = align
  const indent = numberAttribute(findChild(properties, 'tblInd'), '@_w')
  if (indent !== undefined) setup.indent = { v: indent / TWIPS_PER_PIXEL }
  const layout = optionalAttribute(findChild(properties, 'tblLayout'), '@_type')
  if (layout === 'fixed') setup.layout = TABLE_LAYOUT_FIXED
  else if (layout === 'autofit') setup.layout = TABLE_LAYOUT_AUTO
  return setup
}

function readTableWidth(properties: XmlElement): IWidthInTableSize | undefined {
  const width = findChild(properties, 'tblW')
  const value = numberAttribute(width, '@_w')
  if (value === undefined || value <= 0) return undefined
  /* v8 ignore next -- value is read from width, so a defined value implies a defined width. */
  const type = width === undefined ? undefined : attribute(width, '@_type')
  if (type === 'pct') {
    return { type: TABLE_SIZE_SPECIFIED, width: { v: value / 50, u: NumberUnitType.PERCENT } }
  }
  if (type === 'dxa' || type === undefined) {
    return { type: TABLE_SIZE_SPECIFIED, width: { v: value / TWIPS_PER_PIXEL } }
  }
  return undefined
}

function readTableAlignment(properties: XmlElement): TableAlignmentType | undefined {
  const value = optionalAttribute(findChild(properties, 'jc'), '@_val')
  if (value === 'center') return TABLE_ALIGN_CENTER
  if (value === 'end' || value === 'right') return TABLE_ALIGN_END
  if (value === 'start' || value === 'left') return TABLE_ALIGN_START
  return undefined
}

function readGridWidths(element: XmlElement): readonly number[] {
  const grid = findChild(element, 'tblGrid')
  if (grid === undefined) return []
  const widths: number[] = []
  for (const column of childElements(grid)) {
    if (column.tag !== 'gridCol') continue
    const twips = positiveNumber(attribute(column, '@_w'))
    if (twips !== undefined) widths.push(twips / TWIPS_PER_PIXEL)
  }
  return widths
}

function createTable(
  tableId: string,
  rows: readonly ITableRow[],
  columnCount: number,
  gridWidths: readonly number[],
  setup: TableSetup,
): ITable {
  const widths: number[] = []
  for (let index = 0; index < columnCount; index += 1) {
    widths.push(gridWidths[index] ?? DEFAULT_COLUMN_WIDTH)
  }
  const totalWidth = widths.reduce((total, width) => total + width, 0)
  const table: ITable = {
    tableRows: [...rows],
    tableColumns: widths.map(width => ({
      size: { type: TABLE_SIZE_SPECIFIED, width: { v: width } },
    })),
    tableId,
    align: setup.align ?? TABLE_ALIGN_START,
    indent: setup.indent ?? { v: 0 },
    textWrap: TABLE_WRAP_NONE,
    position: {
      positionH: { relativeFrom: RELATIVE_H_PAGE, posOffset: 0 },
      positionV: { relativeFrom: RELATIVE_V_PAGE, posOffset: 0 },
    },
    dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
    cellMargin: cellMargin(),
    size: setup.size ?? { type: TABLE_SIZE_SPECIFIED, width: { v: totalWidth } },
  }
  if (setup.layout !== undefined) table.layout = setup.layout
  return table
}

function cellMargin(): ITableCellMargin {
  return {
    start: { v: 10 },
    end: { v: 10 },
    top: { v: 5 },
    bottom: { v: 5 },
  }
}

function finalizeBody(context: ReadContext, section?: SectionSetup): void {
  if (lastCharacter(context) !== PARAGRAPH) appendParagraphBreak(context, undefined, undefined, undefined, true)
  context.tables.sort((left, right) => left.startIndex - right.startIndex)
  context.sectionBreaks.push({
    sectionId: `section-${context.sectionBreaks.length + 1}`,
    startIndex: context.length,
    ...section,
  })
  append(context, SECTION_BREAK)
}

/**
 * Read header or footer parts into their segment bodies.
 * @param parts - Part identities mapped to their decoded XML.
 * @param root - Root element name: `hdr` for headers, `ftr` for footers.
 * @param bounds - Bounds applied while parsing tables.
 * @param numbering - Numbering definitions shared with the document body.
 * @param relationships - Relationship maps for each header and footer part.
 * @param build - Builds one segment entry from its identity, body, and table metadata.
 * @param collect - Receives each part's bytes and body so a read can anchor them.
 * @returns Segment identity mapped to its entry.
 */
function readStoryParts<Entry>(
  parts: ReadonlyMap<string, Uint8Array>,
  root: 'hdr' | 'ftr',
  bounds: OfficeExchangeConfig,
  numbering: NumberingCatalog,
  relationships: ReadonlyMap<string, DocumentRelationships>,
  build: (id: string, body: IDocumentBody, tableSource: ITables) => Entry,
  collect?: StoryAnchorCollector,
): Map<string, Entry> {
  const stories = new Map<string, Entry>()
  for (const [id, xml] of parts) {
    const element = parseRootElement(xml, root, `${root === 'hdr' ? 'header' : 'footer'} XML`)
    const context = createReadContext(
      bounds,
      readRawRunCatalog(strFromU8(xml)),
      readRawElementCatalog(strFromU8(xml)),
      collect !== undefined,
    )
    const storyRelationships = relationships.get(id)
    for (const child of childElements(element)) {
      if (child.tag === 'p') appendParagraph(context, child, storyRelationships, undefined, numbering)
      if (child.tag === 'tbl') appendTable(context, child, storyRelationships, numbering)
      if (child.tag === 'sdt') appendContentControl(context)
    }
    finalizeBody(context)
    const body: IDocumentBody = {
      dataStream: context.parts.join(''),
      textRuns: context.textRuns,
      paragraphs: context.paragraphs,
      sectionBreaks: context.sectionBreaks,
      tables: context.tables,
      customRanges: context.customRanges,
    }
    if (context.rawBlocks.length > 0) body.docxRawBlocks = context.rawBlocks
    collect?.({
      part: `word/${id}.xml`,
      bytes: xml,
      body,
      tableSource: context.tableSource,
      syntheticParagraphs: context.syntheticParagraphs,
    })
    stories.set(id, build(id, body, context.tableSource))
  }
  return stories
}

/** One table's model metadata with its cell content ranges resolved in the stream. */
interface TableRecord {
  /** Custom table record naming the stream range and model table id. */
  readonly table: ICustomTable
  /** Cell content ranges in row-major order. */
  readonly ranges: ReadonlyArray<ReadonlyArray<ParsedCell>>
}

/**
 * Build the sidecar for a document read when every editable unit resolves to a
 * unique, non-overlapping source span.
 * @param bytes - Complete source package bytes.
 * @param read - Parsed document with its story parts.
 * @param format - Source format recorded in the sidecar.
 * @param bounds - Validated anchor count and payload limits.
 * @returns The validated sidecar, or `undefined` when any part cannot be
 * anchored unambiguously or the payload exceeds a budget.
 */
function buildDocumentSourceState(
  bytes: Uint8Array,
  read: DocumentRead,
  format: OfficeSourceFormat,
  bounds: OfficeExchangeConfig,
): OfficeSourceState | undefined {
  let packageFile: OfficePackage
  try {
    packageFile = open(bytes, { maxEntries: MAX_SOURCE_ENTRIES })
  } catch {
    return undefined
  }
  const anchors: UnitAnchor[] = []
  const sourceParts: string[] = []
  for (const story of read.storySources) {
    const partAnchors = buildStoryAnchors(story)
    if (partAnchors === undefined) return undefined
    sourceParts.push(story.part)
    anchors.push(...partAnchors)
  }
  const partByteLengths = new Map<string, number>()
  for (const entry of packageFile.source.entries) {
    /* v8 ignore next -- source.entries never contains directory records. */
    if (!entry.name.endsWith('/')) partByteLengths.set(entry.name, entry.uncompressedSize)
  }
  try {
    return parseSourceState({
      schemaVersion: 1,
      format,
      packageDigest: packageFile.source.digest,
      byteLength: bytes.byteLength,
      anchorMode: 'per-unit',
      anchors,
      sourceParts,
    }, {
      partByteLengths,
      maxSourceAnchors: bounds.maxSourceAnchors,
      maxSourceAnchorBytes: bounds.maxSourceAnchorBytes,
    })
  } catch {
    return undefined
  }
}

/**
 * Anchor every editable unit in one part.
 *
 * Top-level paragraphs and top-level table cells are paired with scanned spans
 * by source order. A body-level content control takes the unit slot of the
 * placeholder paragraph the reader inserted, and its span is the whole `w:sdt`
 * element so sibling edits never rewrite it. Spans nested inside a preserved
 * container are excluded because their XML is carried with that container. A
 * count mismatch, a malformed stream, or a scanner refusal returns `undefined`
 * so the caller omits the sidecar rather than splicing from a partial map.
 * @param source - Part bytes plus the body and table metadata read from them.
 * @returns Anchors in model order, or `undefined` when the part cannot be
 * anchored unambiguously.
 */
function buildStoryAnchors(source: StoryAnchorSource): readonly UnitAnchor[] | undefined {
  const scan = scanSpans(source.bytes, source.part, SOURCE_ANCHOR_KINDS)
  if (scan.failure !== undefined) return undefined
  const stream = source.body.dataStream
  /* v8 ignore next -- A reader-produced body always carries its paragraph array. */
  const paragraphs = source.body.paragraphs ?? []
  const orderedParagraphs = sortedParagraphs(paragraphs)
  /* v8 ignore next -- A reader-produced body always carries its text-run array. */
  const runs = source.body.textRuns ?? []
  /* v8 ignore next -- A reader-produced body always carries its table array. */
  const tables = [...(source.body.tables ?? [])].sort((left, right) => left.startIndex - right.startIndex)
  const records: TableRecord[] = []
  for (const table of tables) {
    const model = source.tableSource[table.tableId]
    /* v8 ignore next -- Reader table records always carry model metadata. */
    if (model === undefined) return undefined
    const ranges = readTableCellRanges(stream, table, model)
    /* v8 ignore next -- Reader table streams align with their model. */
    if (ranges === undefined) return undefined
    records.push({ table, ranges })
  }
  const preservedSpans = scan.spans.filter(span => span.kind === 'preserved')
  const controlSpans = scan.spans.filter(span => (
    span.kind === 'contentControl' && !containedByAny(span, preservedSpans)
  ))
  const tableSpans = scan.spans.filter(span => (
    span.kind === 'table'
    && !containedByAny(span, preservedSpans)
    && !containedByAny(span, controlSpans)
  ))
  const bodyParagraphs = orderedParagraphs
    .filter(paragraph => !paragraphInsideAnyCell(records, paragraph.startIndex))
    .filter(paragraph => !source.syntheticParagraphs.has(paragraph.startIndex))
  const paragraphSpans = scan.spans.filter(span => (
    span.kind === 'paragraph'
    && !containedByAny(span, preservedSpans)
    && !containedByAny(span, controlSpans)
    && !containedByAny(span, tableSpans)
  ))
  const blockSpans = [...paragraphSpans, ...controlSpans].sort((left, right) => left.start - right.start)
  if (blockSpans.length !== bodyParagraphs.length) return undefined
  const anchors: UnitAnchor[] = []
  for (const [index, paragraph] of bodyParagraphs.entries()) {
    const span = blockSpans[index]
    /* v8 ignore next -- Reader paragraph spans and records are produced in lockstep. */
    if (span === undefined) return undefined
    anchors.push({
      kind: span.kind === 'contentControl' ? 'contentControl' : 'paragraph',
      part: source.part,
      key: paragraph.paragraphId,
      start: span.start,
      end: span.end,
      digest: digestParagraph(paragraphUnitOf(stream, paragraph, runs)),
    })
  }
  const topRecords = topLevelTables(records)
  const topTableSpans = tableSpans.filter(span => (
    !tableSpans.some(other => other !== span && span.start >= other.start && span.end <= other.end)
  ))
  /* v8 ignore next -- Reader table spans and records are produced in lockstep. */
  if (topTableSpans.length !== topRecords.length) return undefined
  for (const [tableIndex, record] of topRecords.entries()) {
    const tableSpan = topTableSpans[tableIndex]
    /* v8 ignore next -- Every reader table record has a scanned span. */
    if (tableSpan === undefined) return undefined
    const nestedSpans = tableSpans.filter(span => (
      span.start > tableSpan.start && span.end < tableSpan.end
    ))
    const cellSpans = scan.spans.filter(span => (
      span.kind === 'tableCell'
      && span.start > tableSpan.start
      && span.end < tableSpan.end
      && !nestedSpans.some(nested => span.start >= nested.start && span.end <= nested.end)
    ))
    const flatRanges = record.ranges.flat()
    /* v8 ignore next -- Reader table cells and scanned cell spans are produced in lockstep. */
    if (cellSpans.length !== flatRanges.length) return undefined
    let cellIndex = 0
    for (const [rowIndex, row] of record.ranges.entries()) {
      for (let column = 0; column < row.length; column += 1) {
        const span = cellSpans[cellIndex]
        cellIndex += 1
        /* v8 ignore next -- The loop index stays inside the scanned cell spans. */
        if (span === undefined) return undefined
        const unit = cellUnitOf(
          stream,
          orderedParagraphs,
          runs,
          records,
          source.tableSource,
          record.table.tableId,
          rowIndex,
          column,
        )
        /* v8 ignore next -- Reader table cells always resolve to a unit. */
        if (unit === undefined) return undefined
        anchors.push({
          kind: 'tableCell',
          part: source.part,
          key: `${record.table.tableId}/${String(rowIndex)}/${String(column)}`,
          start: span.start,
          end: span.end,
          digest: digestTableCell(unit),
        })
      }
    }
  }
  return anchors
}

/**
 * Resolve one table's cell content ranges in the stream.
 * @param stream - Document data stream.
 * @param table - Custom table record with its stream range.
 * @param model - Table metadata naming the row and cell counts.
 * @returns Cell ranges in row-major order, or `undefined` when the tokens and
 * metadata disagree.
 */
function readTableCellRanges(
  stream: string,
  table: ICustomTable,
  model: ITable,
): ReadonlyArray<ReadonlyArray<ParsedCell>> | undefined {
  if (stream[table.startIndex] !== TABLE_START) return undefined
  const rows: ParsedCell[][] = []
  let index = table.startIndex + 1
  for (const row of model.tableRows) {
    if (stream[index] !== TABLE_ROW_START) return undefined
    index += 1
    const cells: ParsedCell[] = []
    for (let cell = 0; cell < row.tableCells.length; cell += 1) {
      if (stream[index] !== TABLE_CELL_START) return undefined
      index += 1
      const start = index
      const end = findCellEnd(stream, start)
      if (end === undefined) return undefined
      cells.push({ start, end })
      index = end + 1
    }
    if (stream[index] !== TABLE_ROW_END) return undefined
    index += 1
    rows.push(cells)
  }
  if (stream[index] !== TABLE_END) return undefined
  if (index + 1 !== table.endIndex) return undefined
  return rows
}

function paragraphInsideAnyCell(records: readonly TableRecord[], startIndex: number): boolean {
  return records.some(record => record.ranges.some(row => (
    row.some(cell => startIndex >= cell.start && startIndex < cell.end)
  )))
}

function containedByAny(
  span: { readonly start: number; readonly end: number },
  containers: readonly ElementSpan[],
): boolean {
  return containers.some(container => span.start >= container.start && span.end <= container.end)
}

function topLevelTables(records: readonly TableRecord[]): readonly TableRecord[] {
  return records.filter(record => !records.some(other => (
    other !== record && other.table.startIndex < record.table.startIndex && record.table.endIndex <= other.table.endIndex
  )))
}

function directNestedTables(records: readonly TableRecord[], parent: TableRecord): readonly TableRecord[] {
  const contained = records.filter(record => (
    record !== parent
    && record.table.startIndex > parent.table.startIndex
    && record.table.endIndex <= parent.table.endIndex
  ))
  return contained.filter(record => !contained.some(other => (
    other !== record && other.table.startIndex < record.table.startIndex && record.table.endIndex <= other.table.endIndex
  )))
}

/**
 * Project one paragraph for change detection.
 * @param stream - Document data stream.
 * @param paragraph - Paragraph record from the body.
 * @param runs - Every run in the body.
 * @returns The paragraph's stream slice, runs, and properties.
 */
function paragraphUnitOf(
  stream: string,
  paragraph: IParagraph,
  runs: readonly ITextRun[],
): ParagraphUnit {
  const start = paragraphContentStart(stream, paragraph.startIndex)
  const end = paragraph.startIndex + 1
  return {
    dataStream: stream.slice(start, end),
    textRuns: runs.filter(run => run.st >= start && run.ed <= end),
    paragraph,
  }
}

/**
 * Locate the first stream character of the paragraph that ends at `startIndex`.
 * @param stream - Document data stream.
 * @param startIndex - Offset of the paragraph's break token.
 * @returns The offset after the preceding paragraph break, section break, table
 * end, or cell start.
 */
function paragraphContentStart(stream: string, startIndex: number): number {
  let index = startIndex
  while (index > 0) {
    const previous = stream[index - 1]
    if (
      previous === PARAGRAPH
      || previous === SECTION_BREAK
      || previous === TABLE_END
      || previous === TABLE_CELL_START
    ) return index
    index -= 1
  }
  return index
}

function tableUnitOf(
  stream: string,
  paragraphs: readonly IParagraph[],
  runs: readonly ITextRun[],
  records: readonly TableRecord[],
  tableSource: ITables,
  tableId: string,
): TableUnit | undefined {
  const record = records.find(candidate => candidate.table.tableId === tableId)
  const model = tableSource[tableId]
  /* v8 ignore next -- Callers pass a table id from the same record set and model. */
  if (record === undefined || model === undefined) return undefined
  const rows: TableRowUnit[] = []
  for (const [rowIndex, row] of model.tableRows.entries()) {
    const cells: TableCellUnit[] = []
    for (let column = 0; column < row.tableCells.length; column += 1) {
      const unit = cellUnitOf(stream, paragraphs, runs, records, tableSource, tableId, rowIndex, column)
      /* v8 ignore next -- Nested rows mirror their model row cells. */
      if (unit === undefined) return undefined
      cells.push(unit)
    }
    rows.push({ row, cells })
  }
  return { table: model, rows }
}

function cellUnitOf(
  stream: string,
  paragraphs: readonly IParagraph[],
  runs: readonly ITextRun[],
  records: readonly TableRecord[],
  tableSource: ITables,
  tableId: string,
  rowIndex: number,
  column: number,
): TableCellUnit | undefined {
  const record = records.find(candidate => candidate.table.tableId === tableId)
  const model = tableSource[tableId]
  const range = record?.ranges[rowIndex]?.[column]
  const cell = model?.tableRows[rowIndex]?.tableCells[column]
  /* v8 ignore next -- Callers pass a row and column from the same record and model. */
  if (record === undefined || range === undefined || cell === undefined) return undefined
  const nested = directNestedTables(records, record)
    .filter(candidate => candidate.table.startIndex >= range.start && candidate.table.startIndex < range.end)
  const cellParagraphs = paragraphsStartingBetween(paragraphs, range.start, range.end)
    .filter(paragraph => !nested.some(candidate => (
      paragraph.startIndex >= candidate.table.startIndex
      && paragraph.startIndex < candidate.table.endIndex
    )))
  const paragraphUnits: ParagraphUnit[] = []
  for (const paragraph of cellParagraphs) paragraphUnits.push(paragraphUnitOf(stream, paragraph, runs))
  const tableUnits: TableUnit[] = []
  for (const candidate of nested) {
    const unit = tableUnitOf(stream, paragraphs, runs, records, tableSource, candidate.table.tableId)
    /* v8 ignore next -- Nested table records have model metadata in the same source. */
    if (unit === undefined) return undefined
    tableUnits.push(unit)
  }
  return { cell, paragraphs: paragraphUnits, tables: tableUnits }
}

/**
 * Sort paragraph records by their stream break offset.
 * @param paragraphs - Paragraph records in any order.
 * @returns A new array in ascending stream order.
 */
function sortedParagraphs(paragraphs: readonly IParagraph[]): readonly IParagraph[] {
  return [...paragraphs].sort((left, right) => left.startIndex - right.startIndex)
}

/**
 * Collect paragraph records whose break falls in a stream range.
 * @param paragraphs - Records sorted by `startIndex`.
 * @param start - Inclusive lower bound.
 * @param end - Exclusive upper bound.
 * @returns The records in stream order.
 */
function paragraphsStartingBetween(
  paragraphs: readonly IParagraph[],
  start: number,
  end: number,
): readonly IParagraph[] {
  let low = 0
  let high = paragraphs.length
  while (low < high) {
    const middle = (low + high) >>> 1
    const paragraph = paragraphs[middle]
    if (paragraph !== undefined && paragraph.startIndex < start) low = middle + 1
    else high = middle
  }
  const selected: IParagraph[] = []
  for (let index = low; index < paragraphs.length; index += 1) {
    const paragraph = paragraphs[index]
    if (paragraph === undefined || paragraph.startIndex >= end) break
    selected.push(paragraph)
  }
  return selected
}

/**
 * Assign list types to numbering ids and encode `word/numbering.xml`.
 * @param data - Document snapshot, including header and footer bullets.
 * @returns The part XML, or `undefined` when the snapshot declares no lists.
 */
function planNumbering(data: IDocumentData): NumberingPlan {
  const lists = data.lists ?? {}
  const listTypes = Object.keys(lists)
  for (const paragraph of allParagraphs(data)) {
    const listType = paragraph.bullet?.listType
    if (listType !== undefined && !listTypes.includes(listType)) listTypes.push(listType)
  }
  if (listTypes.length === 0) return { xml: undefined, numIdByListType: new Map() }
  const numIdByListType = new Map<string, number>()
  for (const [index, listType] of listTypes.entries()) numIdByListType.set(listType, index + 1)
  return { xml: numberingXml(listTypes, lists, listStarts(data)), numIdByListType }
}

/** Paragraph-level list starts keyed by list type, then nesting level. */
function listStarts(data: IDocumentData): ReadonlyMap<string, ReadonlyMap<number, number>> {
  const starts = new Map<string, Map<number, number>>()
  for (const paragraph of allParagraphs(data)) {
    const bullet = paragraph.bullet
    if (bullet?.startNumber === undefined) continue
    const levels = starts.get(bullet.listType) ?? new Map<number, number>()
    levels.set(bullet.nestingLevel, bullet.startNumber)
    starts.set(bullet.listType, levels)
  }
  return starts
}

/** Every paragraph in the document body, headers, and footers. */
function allParagraphs(data: IDocumentData): readonly IParagraph[] {
  const paragraphs: IParagraph[] = [...(data.body?.paragraphs ?? [])]
  for (const header of Object.values(data.headers ?? {})) paragraphs.push(...(header.body.paragraphs ?? []))
  for (const footer of Object.values(data.footers ?? {})) paragraphs.push(...(footer.body.paragraphs ?? []))
  return paragraphs
}

function numberingXml(
  listTypes: readonly string[],
  lists: ILists,
  starts: ReadonlyMap<string, ReadonlyMap<number, number>>,
): string {
  let abstractNumbers = ''
  let instances = ''
  for (const [index, listType] of listTypes.entries()) {
    const levels = lists[listType]?.nestingLevel ?? [defaultNestingLevel(0)]
    abstractNumbers += `<w:abstractNum w:abstractNumId="${String(index)}">`
      + '<w:multiLevelType w:val="hybridMultilevel"/>'
      + levels.map((level, levelIndex) => writeNestingLevel(level, levelIndex)).join('')
      + '</w:abstractNum>'
    instances += `<w:num w:numId="${String(index + 1)}">`
      + `<w:abstractNumId w:val="${String(index)}"/>`
    for (const [level, start] of [...(starts.get(listType) ?? new Map<number, number>()).entries()]) {
      instances += `<w:lvlOverride w:ilvl="${String(level)}"><w:startOverride w:val="${String(start)}"/></w:lvlOverride>`
    }
    instances += '</w:num>'
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + `${abstractNumbers}${instances}</w:numbering>`
}

function writeNestingLevel(level: INestingLevel, index: number): string {
  const glyphType = level.glyphType ?? ListGlyphType.DECIMAL
  const format = nameOf(NUMBER_FORMATS, glyphType) ?? 'decimal'
  const text = glyphType === ListGlyphType.BULLET
    ? level.glyphSymbol ?? level.glyphFormat
    : level.glyphFormat.length > 0
      ? level.glyphFormat
      : `%${String(index + 1)}.`
  let xml = `<w:lvl w:ilvl="${String(index)}">`
    + `<w:start w:val="${String(level.startNumber)}"/>`
    + `<w:numFmt w:val="${format}"/>`
    + `<w:lvlText w:val="${escapeXml(text)}"/>`
  if (level.isLegal === true) xml += '<w:isLgl/>'
  const paragraphProperties = level.paragraphProperties === undefined
    ? ''
    : writeParagraphFormatting(level.paragraphProperties)
  if (paragraphProperties.length > 0) xml += `<w:pPr>${paragraphProperties}</w:pPr>`
  const runProperties = writeRunProperties(level.textStyle)
  if (runProperties.length > 0) xml += runProperties
  return `${xml}</w:lvl>`
}

/**
 * Assign external hyperlink relationships for one part.
 * @param ranges - Custom ranges from the part's body.
 * @param allocator - Relationship id allocator for this part.
 * @returns The relationship declarations and a URL lookup.
 */
function planLinks(ranges: readonly ICustomRange[] | undefined, allocator: RelationshipAllocator): LinkPlan {
  const relIdByUrl = new Map<string, string>()
  const rels: LinkRel[] = []
  for (const range of ranges ?? []) {
    const url = hyperlinkPropertiesOf(range)?.url
    if (url === undefined || relIdByUrl.has(url)) continue
    const relId = allocator.next()
    relIdByUrl.set(url, relId)
    rels.push({ relId, target: url })
  }
  return { rels, relIdOf: url => relIdByUrl.get(url) }
}

/**
 * Read a hyperlink custom range's target.
 * @param range - Custom range from a document body.
 * @returns The URL or bookmark identity, or `undefined` for another range type.
 */
function hyperlinkPropertiesOf(range: ICustomRange): IHyperlink | undefined {
  const properties = range.properties
  if (range.rangeType !== HYPERLINK_RANGE_TYPE || properties === undefined) return undefined
  const link: IHyperlink = {}
  const url = stringPropertyOf(properties, 'url')
  if (url !== undefined && url.length > 0) link.url = url
  const bookmarkId = stringPropertyOf(properties, 'bookmarkId')
  if (bookmarkId !== undefined && bookmarkId.length > 0) link.bookmarkId = bookmarkId
  return link.url === undefined && link.bookmarkId === undefined ? undefined : link
}

function stringPropertyOf(properties: Readonly<Record<string, unknown>>, name: string): string | undefined {
  const value = properties[name]
  return typeof value === 'string' ? value : undefined
}

/**
 * Clip the hyperlink ranges that intersect one paragraph.
 * @param ranges - Custom ranges from the part's body.
 * @param start - Inclusive paragraph start offset.
 * @param end - Exclusive paragraph end offset.
 * @returns Ranges with their stream bounds and targets, sorted by start.
 */
function paragraphLinks(
  ranges: readonly ICustomRange[] | undefined,
  start: number,
  end: number,
): readonly ParagraphLink[] {
  const links: ParagraphLink[] = []
  for (const range of ranges ?? []) {
    const properties = hyperlinkPropertiesOf(range)
    if (properties === undefined) continue
    const linkStart = Math.max(start, range.startIndex)
    const linkEnd = Math.min(end, range.endIndex + 1)
    if (linkEnd <= linkStart) continue
    links.push({ start: linkStart, end: linkEnd, properties })
  }
  return links.sort((left, right) => left.start - right.start)
}

function writeBody(
  stream: string,
  paragraphs: ReadonlyMap<number, IParagraph>,
  story: StoryData,
  sections: SectionPlan,
  relIds: ReadonlyMap<string, string>,
  links: LinkPlan,
  numbering: NumberingPlan,
  bounds: OfficeExchangeConfig,
): string {
  const cells = { count: 0 }
  const rawBlocks = planRawBlocks(story.body)
  let xml = ''
  let cursor = 0
  let index = 0
  while (index < stream.length) {
    const token = stream[index]
    if (token === TABLE_START) {
      const table = readTableRange(stream, index)
      xml += writeTable(stream, index, table, paragraphs, story, bounds, cells, links, numbering, rawBlocks)
      index = table.end
      cursor = index
    } else if (token === PARAGRAPH) {
      const section = stream[index + 1] === SECTION_BREAK ? sections.breakAt.get(index + 1) : undefined
      const properties = section === undefined ? '' : writeSectionProperties(section, relIds)
      xml += writeParagraph(
        stream,
        cursor,
        index,
        paragraphs.get(index),
        story,
        properties,
        links,
        numbering,
        rawBlocks,
      )
      cursor = section === undefined ? index + 1 : index + 2
      index = cursor
    } else {
      index += 1
    }
  }
  const remainder = stream.slice(cursor).replaceAll(SECTION_BREAK, '')
  if (remainder.length > 0) {
    throw new OfficeExchangeError('conversion-failed', 'Document stream has text outside a paragraph')
  }
  assertRawBlocksEmitted(rawBlocks)
  return xml
}

/**
 * Index a body's raw XML by the stream offset where the writer must emit it.
 * @param body - Body whose `docxRawBlocks` are preserved.
 * @returns Positions in source order and an empty emitted set.
 */
function planRawBlocks(body: IDocumentBody | undefined): RawBlockPlan {
  const byIndex = new Map<number, string[]>()
  for (const block of body?.docxRawBlocks ?? []) {
    const existing = byIndex.get(block.startIndex)
    if (existing === undefined) byIndex.set(block.startIndex, [block.xml])
    else existing.push(block.xml)
  }
  return { byIndex, emitted: new Set() }
}

/**
 * Emit raw XML anchored at one offset exactly once.
 * @param plan - Raw-block positions for the body.
 * @param index - Stream offset to consume.
 * @returns Concatenated source XML, or an empty string when nothing is anchored there.
 */
function consumeRawBlocks(plan: RawBlockPlan, index: number): string {
  const blocks = plan.byIndex.get(index)
  if (blocks === undefined || plan.emitted.has(index)) return ''
  plan.emitted.add(index)
  return blocks.join('')
}

/**
 * Refuse to write when a preserved raw block has no paragraph or cell anchor.
 * @param plan - Raw-block positions for the body.
 * @throws {OfficeExchangeError} `conversion-failed` when any offset was not emitted.
 */
function assertRawBlocksEmitted(plan: RawBlockPlan): void {
  for (const index of plan.byIndex.keys()) {
    if (plan.emitted.has(index)) continue
    throw new OfficeExchangeError(
      'conversion-failed',
      `Preserved document content at stream offset ${String(index)} is outside a paragraph`,
    )
  }
}

/**
 * Write one header or footer body.
 * @param story - Header or footer segment content.
 * @param bounds - Table-cell bound applied while writing.
 * @param links - Hyperlink relationships this part declares.
 * @param numbering - Numbering ids shared with the document body.
 * @returns OOXML block content for the part root.
 */
function writeStoryBody(
  story: StoryData & { readonly body: IDocumentBody },
  bounds: OfficeExchangeConfig,
  links: LinkPlan,
  numbering: NumberingPlan,
): string {
  const paragraphs = new Map<number, IParagraph>()
  for (const paragraph of story.body.paragraphs ?? []) paragraphs.set(paragraph.startIndex, paragraph)
  return writeBody(
    story.body.dataStream,
    paragraphs,
    story,
    { breakAt: new Map(), finalXml: '' },
    new Map(),
    links,
    numbering,
    bounds,
  )
}

/**
 * Schedule header or footer parts for the writer.
 * @param entries - Segment bodies keyed by their identity.
 * @param kind - Whether the parts are headers or footers.
 * @param allocator - Document relationship id allocator shared with the body.
 * @param bounds - Table-cell bound applied while writing story bodies.
 * @param numbering - Numbering ids shared with the document body.
 * @param preserved - Source package state re-emitted with the document.
 * @returns The parts, their relationship ids, and the next free relationship number.
 */
function planStoryParts(
  entries: IHeaders | IFooters | undefined,
  kind: 'header' | 'footer',
  allocator: RelationshipAllocator,
  bounds: OfficeExchangeConfig,
  numbering: NumberingPlan,
  preserved: DocxPreservedPackage | undefined,
): StoryPlan {
  const ids = entries === undefined ? [] : Object.keys(entries).sort()
  const parts: StoryPart[] = []
  const relIds = new Map<string, string>()
  for (const [index, id] of ids.entries()) {
    const entry = entries?.[id]
    /* v8 ignore next -- ids came from this record's own keys, so every lookup resolves. */
    if (entry === undefined) continue
    const partName = `word/${kind}${String(index + 1)}.xml`
    const relId = allocator.next()
    const preservedRels = preserved?.relationships.stories?.[id] ?? []
    const storyAllocator = createRelationshipAllocator(relationshipIdsOf(preservedRels))
    const links = planLinks(entry.body.customRanges, storyAllocator)
    const relsXml = storyRelsXml(links.rels, preservedRels)
    parts.push({
      id,
      partName,
      relId,
      xml: storyXml(kind, writeStoryBody(entry, bounds, links, numbering)),
      ...(relsXml === undefined
        ? {}
        : { relsPartName: `word/_rels/${partName.slice('word/'.length)}.rels`, relsXml }),
    })
    relIds.set(id, relId)
  }
  return { parts, relIds }
}

function storyRelsXml(rels: readonly LinkRel[], preserved: readonly string[]): string | undefined {
  if (rels.length === 0 && preserved.length === 0) return undefined
  let relationships = preserved.join('')
  for (const rel of rels) {
    relationships += `<Relationship Id="${rel.relId}"`
      + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'
      + ` Target="${escapeXml(rel.target)}" TargetMode="External"/>`
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + `${relationships}</Relationships>`
}

function storyXml(kind: 'header' | 'footer', content: string): string {
  const root = kind === 'header' ? 'hdr' : 'ftr'
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<w:${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"${PRESERVED_NAMESPACES}>`
    + `${content}</w:${root}>`
}

/**
 * Split the snapshot's sections into paragraph-level breaks and the body-level final section.
 * @param sections - Section metadata in document order.
 * @param relIds - Segment identity mapped to its relationship id.
 * @returns Non-final sections by stream offset and the final `w:sectPr` XML.
 */
function planSections(
  sections: readonly ISectionBreak[],
  relIds: ReadonlyMap<string, string>,
): SectionPlan {
  const finalSection = lastSectionOf(sections)
  const breakAt = new Map<number, ISectionBreak>()
  for (const section of sections) {
    if (section !== finalSection) breakAt.set(section.startIndex, section)
  }
  return { breakAt, finalXml: writeSectionProperties(finalSection, relIds) }
}

function lastSectionOf(sections: readonly ISectionBreak[]): ISectionBreak | undefined {
  let last: ISectionBreak | undefined
  for (const section of sections) {
    if (last === undefined || section.startIndex > last.startIndex) last = section
  }
  return last
}

/**
 * Write one section's properties as a `w:sectPr` element.
 * @param section - Section metadata; absent writes an empty section.
 * @param relIds - Segment identity mapped to its relationship id.
 * @returns Section properties XML.
 */
function writeSectionProperties(
  section: ISectionBreak | undefined,
  relIds: ReadonlyMap<string, string>,
): string {
  if (section === undefined) return '<w:sectPr/>'
  let xml = ''
  xml += writeStoryReference('headerReference', 'default', section.defaultHeaderId, relIds)
  xml += writeStoryReference('headerReference', 'first', section.firstPageHeaderId, relIds)
  xml += writeStoryReference('headerReference', 'even', section.evenPageHeaderId, relIds)
  xml += writeStoryReference('footerReference', 'default', section.defaultFooterId, relIds)
  xml += writeStoryReference('footerReference', 'first', section.firstPageFooterId, relIds)
  xml += writeStoryReference('footerReference', 'even', section.evenPageFooterId, relIds)
  const sectionType = section.sectionType === undefined
    ? undefined
    : nameOf(SECTION_TYPES, section.sectionType)
  if (sectionType !== undefined) xml += `<w:type w:val="${sectionType}"/>`
  const size = section.pageSize
  if (size?.width !== undefined && size.height !== undefined) {
    const orientation = section.pageOrient === PageOrientType.LANDSCAPE ? ' w:orient="landscape"' : ''
    xml += `<w:pgSz w:w="${twipsOf(size.width)}" w:h="${twipsOf(size.height)}"${orientation}/>`
  }
  xml += writeSectionMargins(section)
  xml += writeSectionColumns(section)
  if (section.useFirstPageHeaderFooter === BOOLEAN_TRUE) xml += '<w:titlePg/>'
  return `<w:sectPr>${xml}</w:sectPr>`
}

function writeStoryReference(
  tag: 'headerReference' | 'footerReference',
  variant: StoryVariant,
  id: string | undefined,
  relIds: ReadonlyMap<string, string>,
): string {
  if (id === undefined) return ''
  const relId = relIds.get(id)
  if (relId === undefined) return ''
  return `<w:${tag} w:type="${variant}" r:id="${relId}"/>`
}

function writeSectionMargins(section: ISectionBreak): string {
  const attributes: string[] = []
  if (section.marginTop !== undefined) attributes.push(`w:top="${twipsOf(section.marginTop)}"`)
  if (section.marginRight !== undefined) attributes.push(`w:right="${twipsOf(section.marginRight)}"`)
  if (section.marginBottom !== undefined) attributes.push(`w:bottom="${twipsOf(section.marginBottom)}"`)
  if (section.marginLeft !== undefined) attributes.push(`w:left="${twipsOf(section.marginLeft)}"`)
  if (section.marginHeader !== undefined) attributes.push(`w:header="${twipsOf(section.marginHeader)}"`)
  if (section.marginFooter !== undefined) attributes.push(`w:footer="${twipsOf(section.marginFooter)}"`)
  return attributes.length === 0 ? '' : `<w:pgMar ${attributes.join(' ')}/>`
}

function writeSectionColumns(section: ISectionBreak): string {
  const columns = section.columnProperties
  const separator = section.columnSeparatorType === ColumnSeparatorType.BETWEEN_EACH_COLUMN
  const separatorAttribute = separator ? ' w:sep="1"' : ''
  if (columns === undefined || columns.length === 0) {
    return section.columnSeparatorType === undefined ? '' : `<w:cols w:num="1"${separatorAttribute}/>`
  }
  const spacing = Math.max(0, ...columns.map(column => column.paddingEnd))
  let xml = `<w:cols w:num="${String(columns.length)}" w:space="${twipsOf(spacing)}"${separatorAttribute}>`
  for (const column of columns) {
    xml += `<w:col w:w="${twipsOf(column.width)}" w:space="${twipsOf(column.paddingEnd)}"/>`
  }
  return `${xml}</w:cols>`
}

function writeSettingsXml(documentStyle: IDocumentStyle): string | undefined {
  if (documentStyle.evenAndOddHeaders !== BOOLEAN_TRUE) return undefined
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:evenAndOddHeaders/></w:settings>'
}

function readTableRange(stream: string, start: number): ParsedTable {
  let index = start + 1
  const rows: ParsedCell[][] = []
  while (stream[index] === TABLE_ROW_START) {
    index += 1
    const cells: ParsedCell[] = []
    while (stream[index] === TABLE_CELL_START) {
      index += 1
      const end = findCellEnd(stream, index)
      if (end === undefined) {
        throw new OfficeExchangeError('conversion-failed', 'Document stream has an unterminated table cell')
      }
      cells.push({ start: index, end })
      index = end + 1
    }
    if (stream[index] !== TABLE_ROW_END) {
      throw new OfficeExchangeError('conversion-failed', 'Document stream has an unterminated table row')
    }
    index += 1
    rows.push(cells)
  }
  if (stream[index] !== TABLE_END) {
    throw new OfficeExchangeError('conversion-failed', 'Document stream has an unterminated table')
  }
  return { rows, end: index + 1 }
}

function findCellEnd(stream: string, start: number): number | undefined {
  let depth = 0
  for (let index = start; index < stream.length; index += 1) {
    const token = stream[index]
    if (token === TABLE_CELL_START) depth += 1
    if (token !== TABLE_CELL_END) continue
    if (depth === 0) return index
    depth -= 1
  }
  return undefined
}

function writeTable(
  stream: string,
  start: number,
  table: ParsedTable,
  paragraphs: ReadonlyMap<number, IParagraph>,
  story: StoryData,
  bounds: OfficeExchangeConfig,
  cells: { count: number },
  links: LinkPlan,
  numbering: NumberingPlan,
  rawBlocks: RawBlockPlan,
): string {
  const source = tableSourceAt(story, start)
  const sourceColumns = source?.tableColumns ?? []
  const sourceRows = source?.tableRows ?? []
  const columnCount = Math.max(
    1,
    ...table.rows.map(row => row.length),
    sourceColumns.length,
    ...sourceRows.map(row => row.tableCells.reduce((count, cell) => count + cellSpan(cell), 0)),
  )
  const widths: number[] = []
  for (let index = 0; index < columnCount; index += 1) {
    /* v8 ignore next -- columnCount includes every span position a row can occupy. */
    widths.push(sourceColumns[index]?.size.width.v ?? DEFAULT_COLUMN_WIDTH)
  }
  let xml = `<w:tbl>${writeTableProperties(source)}<w:tblGrid>`
  for (const width of widths) {
    xml += `<w:gridCol w:w="${String(Math.max(1, Math.round(width * TWIPS_PER_PIXEL)))}"/>`
  }
  xml += '</w:tblGrid>'
  for (const [rowIndex, row] of table.rows.entries()) {
    xml += `<w:tr>${writeRowProperties(sourceRows[rowIndex])}`
    let gridColumn = 0
    for (const [cellIndex, cell] of row.entries()) {
      cells.count += 1
      if (cells.count > bounds.maxCells) {
        throw new OfficeExchangeError('too-large', `Document has more than ${bounds.maxCells} table cells`)
      }
      const sourceCell = sourceRows[rowIndex]?.tableCells[cellIndex]
      const span = cellSpan(sourceCell)
      const width = sourceCell?.size?.width.v ?? spanWidth(widths, gridColumn, span)
      xml += `<w:tc>${writeCellProperties(sourceCell, width, span)}`
        + `${writeCell(stream, cell, paragraphs, story, bounds, cells, links, numbering, rawBlocks)}</w:tc>`
      gridColumn += span
    }
    xml += '</w:tr>'
  }
  return `${xml}</w:tbl>`
}

/**
 * Write table-level width, alignment, indentation, and layout.
 * @param table - Table metadata; absent writes an automatic preferred width.
 * @returns `w:tblPr` children in schema order.
 */
function writeTableProperties(table: ITable | undefined): string {
  let xml = table === undefined ? '<w:tblW w:w="0" w:type="auto"/>' : writeTableWidth(table.size)
  if (table?.align === TABLE_ALIGN_CENTER) xml += '<w:jc w:val="center"/>'
  else if (table?.align === TABLE_ALIGN_END) xml += '<w:jc w:val="end"/>'
  if (table?.indent !== undefined && table.indent.v !== 0) {
    xml += `<w:tblInd w:w="${twipsOf(table.indent.v)}" w:type="dxa"/>`
  }
  xml += TABLE_BORDERS
  if (table?.layout === TABLE_LAYOUT_FIXED) xml += '<w:tblLayout w:type="fixed"/>'
  else if (table?.layout === TABLE_LAYOUT_AUTO) xml += '<w:tblLayout w:type="autofit"/>'
  return `<w:tblPr>${xml}</w:tblPr>`
}

function writeTableWidth(size: IWidthInTableSize): string {
  if (size.type === TABLE_SIZE_UNSPECIFIED) return '<w:tblW w:w="0" w:type="auto"/>'
  if (size.width.u === NumberUnitType.PERCENT) {
    return `<w:tblW w:w="${String(Math.max(1, Math.round(size.width.v * 50)))}" w:type="pct"/>`
  }
  return `<w:tblW w:w="${twipsOf(size.width.v)}" w:type="dxa"/>`
}

function cellSpan(cell: ITableCell | undefined): number {
  const span = cell?.columnSpan ?? 1
  return Number.isInteger(span) && span > 1 ? span : 1
}

function spanWidth(widths: readonly number[], start: number, span: number): number {
  let total = 0
  for (let index = start; index < start + span; index += 1) {
    /* v8 ignore next -- columnCount covers every column a cell span can reach. */
    total += widths[index] ?? DEFAULT_COLUMN_WIDTH
  }
  return total
}

/**
 * Write one row's height, break-prevention, and header-repeat properties.
 * @param row - Row metadata from the snapshot, when the stream start has one.
 * @returns A `w:trPr` element, or an empty string for a row without metadata.
 */
function writeRowProperties(row: ITableRow | undefined): string {
  if (row === undefined) return ''
  let xml = `<w:trHeight w:val="${twipsOf(row.trHeight.val.v)}"`
    + ` w:hRule="${rowHeightRuleName(row.trHeight.hRule)}"/>`
  if (row.cantSplit === BOOLEAN_TRUE) xml += '<w:cantSplit/>'
  if (row.repeatHeaderRow === BOOLEAN_TRUE) xml += '<w:tblHeader/>'
  return `<w:trPr>${xml}</w:trPr>`
}

function rowHeightRuleName(rule: TableRowHeightRule): string {
  if (rule === TABLE_ROW_HEIGHT_EXACT) return 'exact'
  if (rule === TABLE_ROW_HEIGHT_AT_LEAST) return 'atLeast'
  return 'auto'
}

/**
 * Write one cell's width, spans, borders, shading, margins, and vertical alignment.
 * @param cell - Cell metadata from the snapshot, when the stream start has one.
 * @param width - Preferred cell width in pixels.
 * @param span - Grid columns the cell occupies.
 * @returns The cell's `w:tcPr` element.
 */
function writeCellProperties(cell: ITableCell | undefined, width: number, span: number): string {
  let xml = '<w:tcPr>'
    + `<w:tcW w:w="${String(Math.max(1, Math.round(width * TWIPS_PER_PIXEL)))}" w:type="dxa"/>`
  if (span > 1) xml += `<w:gridSpan w:val="${String(span)}"/>`
  if (cell?.rowSpan === 0) xml += '<w:vMerge/>'
  else if (cell?.rowSpan !== undefined && cell.rowSpan > 1) xml += '<w:vMerge w:val="restart"/>'
  xml += writeCellBorders(cell)
  const background = cell?.backgroundColor === undefined ? undefined : hexColor(cell.backgroundColor.rgb)
  if (background !== undefined) xml += `<w:shd w:val="clear" w:color="auto" w:fill="${background}"/>`
  xml += writeCellMargin(cell?.margin)
  const vertical = cell?.vAlign === undefined ? undefined : verticalAlignmentName(cell.vAlign)
  if (vertical !== undefined) xml += `<w:vAlign w:val="${vertical}"/>`
  return `${xml}</w:tcPr>`
}

function writeCellBorders(cell: ITableCell | undefined): string {
  if (cell === undefined) return ''
  let xml = ''
  xml += writeCellBorder('top', cell.borderTop)
  xml += writeCellBorder('left', cell.borderLeft)
  xml += writeCellBorder('bottom', cell.borderBottom)
  xml += writeCellBorder('right', cell.borderRight)
  return xml.length === 0 ? '' : `<w:tcBorders>${xml}</w:tcBorders>`
}

function writeCellBorder(tag: string, border: ITableCellBorder | undefined): string {
  if (border === undefined) return ''
  const color = hexColor(border.color.rgb) ?? DEFAULT_BORDER_COLOR
  const dash = nameOf(BORDER_DASH_STYLES, border.dashStyle) ?? 'single'
  const width = border.width === undefined ? DEFAULT_BORDER_SIZE / BORDER_UNITS_PER_PIXEL : border.width.v
  const size = Math.max(1, Math.round(width * BORDER_UNITS_PER_PIXEL))
  return `<w:${tag} w:val="${dash}" w:sz="${String(size)}" w:space="0" w:color="${color}"/>`
}

function writeCellMargin(margin: ITableCellMargin | undefined): string {
  if (margin === undefined) return ''
  const defaults = cellMargin()
  if (
    margin.start.v === defaults.start.v
    && margin.end.v === defaults.end.v
    && margin.top.v === defaults.top.v
    && margin.bottom.v === defaults.bottom.v
  ) return ''
  return '<w:tcMar>'
    + `<w:top w:w="${twipsOf(margin.top.v)}" w:type="dxa"/>`
    + `<w:start w:w="${twipsOf(margin.start.v)}" w:type="dxa"/>`
    + `<w:bottom w:w="${twipsOf(margin.bottom.v)}" w:type="dxa"/>`
    + `<w:end w:w="${twipsOf(margin.end.v)}" w:type="dxa"/>`
    + '</w:tcMar>'
}

function verticalAlignmentName(vertical: VerticalAlignmentType): string | undefined {
  if (vertical === VerticalAlignmentType.TOP) return 'top'
  if (vertical === VerticalAlignmentType.CENTER) return 'center'
  if (vertical === VerticalAlignmentType.BOTTOM) return 'bottom'
  return undefined
}

function tableSourceAt(story: StoryData, startIndex: number): ITable | undefined {
  const range = story.body?.tables?.find(table => table.startIndex === startIndex)
  return range === undefined ? undefined : story.tableSource?.[range.tableId]
}

function writeCell(
  stream: string,
  cell: ParsedCell,
  paragraphs: ReadonlyMap<number, IParagraph>,
  story: StoryData,
  bounds: OfficeExchangeConfig,
  cells: { count: number },
  links: LinkPlan,
  numbering: NumberingPlan,
  rawBlocks: RawBlockPlan,
): string {
  let xml = ''
  let cursor = cell.start
  let index = cell.start
  while (index < cell.end) {
    const token = stream[index]
    if (token === PARAGRAPH) {
      xml += writeParagraph(stream, cursor, index, paragraphs.get(index), story, '', links, numbering, rawBlocks)
      cursor = index + 1
      index += 1
    } else if (token === TABLE_START) {
      const nested = readTableRange(stream, index)
      xml += writeTable(stream, index, nested, paragraphs, story, bounds, cells, links, numbering, rawBlocks)
      index = nested.end
      cursor = index
    } else {
      index += 1
    }
  }
  const remainder = stream.slice(cursor, cell.end).replaceAll(SECTION_BREAK, '')
  if (remainder.length > 0) {
    throw new OfficeExchangeError('conversion-failed', 'Table cell text must end with a paragraph break')
  }
  return xml.length === 0
    ? writeParagraph(stream, cell.start, cell.start, undefined, story, '', links, numbering, rawBlocks)
    : xml
}

function writeParagraph(
  stream: string,
  start: number,
  end: number,
  paragraph: IParagraph | undefined,
  story: StoryData,
  sectionProperties: string,
  links: LinkPlan,
  numbering: NumberingPlan,
  rawBlocks: RawBlockPlan,
): string {
  const properties = paragraphProperties(paragraph, sectionProperties, numbering)
  return `<w:p>${properties}${writeParagraphContent(stream, start, end, story, links, rawBlocks)}</w:p>`
}

function writeParagraphContent(
  stream: string,
  start: number,
  end: number,
  story: StoryData,
  links: LinkPlan,
  rawBlocks: RawBlockPlan,
): string {
  const ranges = paragraphLinks(story.body?.customRanges, start, end)
  const boundaries = new Set<number>([start, end])
  for (const range of ranges) {
    boundaries.add(range.start)
    boundaries.add(range.end)
  }
  for (const index of rawBlocks.byIndex.keys()) {
    if (index >= start && index <= end) boundaries.add(index)
  }
  const sorted = [...boundaries].sort((left, right) => left - right)
  let xml = ''
  for (let index = 0; index < sorted.length - 1; index += 1) {
    const sliceStart = sorted[index]
    const sliceEnd = sorted[index + 1]
    /* v8 ignore next -- Boundaries come in sorted pairs, so both offsets are always present. */
    if (sliceStart === undefined || sliceEnd === undefined) continue
    const content = consumeRawBlocks(rawBlocks, sliceStart)
      + writeTextRuns(stream, sliceStart, sliceEnd, story.body?.textRuns ?? [])
    const range = ranges.find(candidate => candidate.start <= sliceStart && candidate.end >= sliceEnd)
    xml += range === undefined ? content : writeHyperlink(range.properties, links, content)
  }
  return `${xml}${consumeRawBlocks(rawBlocks, end)}`
}

function writeTextRuns(stream: string, start: number, end: number, textRuns: readonly ITextRun[]): string {
  const runs = runSlices(textRuns, start, end)
  if (runs.length === 0) return writeRun(stream.slice(start, end), undefined)
  let xml = ''
  for (const run of runs) xml += writeRun(stream.slice(run.start, run.end), run.style)
  return xml
}

function writeHyperlink(properties: IHyperlink, links: LinkPlan, content: string): string {
  const attributes: string[] = []
  if (properties.url !== undefined) {
    const relId = links.relIdOf(properties.url)
    if (relId !== undefined) attributes.push(`r:id="${relId}"`)
  }
  if (properties.bookmarkId !== undefined) {
    attributes.push(`w:anchor="${escapeXml(properties.bookmarkId)}"`)
  }
  /* v8 ignore next -- Every planned hyperlink range declares a URL or bookmark anchor. */
  if (attributes.length === 0) return content
  return `<w:hyperlink ${attributes.join(' ')}>${content}</w:hyperlink>`
}

interface RunSlice {
  readonly start: number
  readonly end: number
  readonly style: ITextStyle | undefined
}

function runSlices(textRuns: readonly ITextRun[], start: number, end: number): readonly RunSlice[] {
  const boundaries = new Set<number>([start, end])
  const covered: RunSlice[] = []
  for (const run of textRuns) {
    const clippedStart = Math.max(start, run.st)
    const clippedEnd = Math.min(end, run.ed)
    if (clippedEnd <= clippedStart) continue
    boundaries.add(clippedStart)
    boundaries.add(clippedEnd)
    covered.push({ start: clippedStart, end: clippedEnd, style: run.ts })
  }
  if (covered.length === 0) return []
  const sorted = [...boundaries].sort((left, right) => left - right)
  const slices: RunSlice[] = []
  for (let index = 0; index < sorted.length - 1; index += 1) {
    const sliceStart = sorted[index]
    const sliceEnd = sorted[index + 1]
    /* v8 ignore next -- Boundaries come in sorted pairs, so both offsets are always present. */
    if (sliceStart === undefined || sliceEnd === undefined) continue
    const style = covered.find(run => run.start <= sliceStart && run.end >= sliceEnd)?.style
    slices.push({ start: sliceStart, end: sliceEnd, style })
  }
  return slices
}

function paragraphProperties(
  paragraph: IParagraph | undefined,
  sectionProperties: string,
  numbering: NumberingPlan,
): string {
  const styleId = paragraph?.styleId ?? styleIdForNamedStyle(paragraph?.paragraphStyle?.namedStyleType)
  const formatting = paragraph?.paragraphStyle === undefined
    ? ''
    : writeParagraphFormatting(paragraph.paragraphStyle)
  const numberProperties = writeNumberingProperties(paragraph?.bullet, numbering)
  if (
    styleId === undefined
    && formatting.length === 0
    && numberProperties.length === 0
    && sectionProperties.length === 0
  ) return ''
  const named = styleId === undefined ? '' : `<w:pStyle w:val="${escapeXml(styleId)}"/>`
  return `<w:pPr>${named}${numberProperties}${formatting}${sectionProperties}</w:pPr>`
}

function writeNumberingProperties(bullet: IBullet | undefined, numbering: NumberingPlan): string {
  if (bullet === undefined) return ''
  const numId = numbering.numIdByListType.get(bullet.listType)
  /* v8 ignore next -- planNumbering assigns an id to every bullet list type. */
  if (numId === undefined) return ''
  return `<w:numPr><w:ilvl w:val="${String(bullet.nestingLevel)}"/>`
    + `<w:numId w:val="${String(numId)}"/></w:numPr>`
}

function writeParagraphFormatting(style: IParagraphStyle): string {
  let properties = ''
  if (style.horizontalAlign !== undefined) {
    properties += `<w:jc w:val="${horizontalAlignName(style.horizontalAlign)}"/>`
  }
  properties += writeParagraphSpacing(style)
  properties += writeParagraphIndent(style)
  properties += writeTabStops(style.tabStops)
  properties += writeParagraphBorders(style)
  const shading = style.shading === undefined ? undefined : hexColor(style.shading.backgroundColor.rgb)
  if (shading !== undefined) properties += `<w:shd w:val="clear" w:color="auto" w:fill="${shading}"/>`
  if (style.outlineLevel !== undefined) properties += `<w:outlineLvl w:val="${String(style.outlineLevel)}"/>`
  return properties
}

function horizontalAlignName(alignment: HorizontalAlign): string {
  if (alignment === HorizontalAlign.CENTER) return 'center'
  if (alignment === HorizontalAlign.RIGHT) return 'right'
  if (alignment === HorizontalAlign.BOTH || alignment === HorizontalAlign.JUSTIFIED) return 'both'
  if (alignment === HorizontalAlign.DISTRIBUTED) return 'distribute'
  return 'left'
}

function writeParagraphSpacing(style: IParagraphStyle): string {
  const attributes: string[] = []
  if (style.spaceAbove !== undefined) {
    attributes.push(`w:before="${String(Math.round(style.spaceAbove.v * TWIPS_PER_PIXEL))}"`)
  }
  if (style.spaceBelow !== undefined) {
    attributes.push(`w:after="${String(Math.round(style.spaceBelow.v * TWIPS_PER_PIXEL))}"`)
  }
  if (style.lineSpacing !== undefined) {
    if (style.spacingRule === SpacingRule.AT_LEAST || style.spacingRule === SpacingRule.EXACT) {
      attributes.push(`w:line="${String(Math.round(style.lineSpacing * TWIPS_PER_PIXEL))}"`)
      attributes.push(`w:lineRule="${style.spacingRule === SpacingRule.AT_LEAST ? 'atLeast' : 'exact'}"`)
    } else {
      attributes.push(`w:line="${String(Math.round(style.lineSpacing * AUTO_LINE_SPACING_UNITS))}"`)
      attributes.push('w:lineRule="auto"')
    }
  }
  return attributes.length === 0 ? '' : `<w:spacing ${attributes.join(' ')}/>`
}

function writeParagraphIndent(style: IParagraphStyle): string {
  const attributes: string[] = []
  if (style.indentStart !== undefined) attributes.push(`w:left="${twipsOf(style.indentStart.v)}"`)
  if (style.indentEnd !== undefined) attributes.push(`w:right="${twipsOf(style.indentEnd.v)}"`)
  if (style.indentFirstLine !== undefined) attributes.push(`w:firstLine="${twipsOf(style.indentFirstLine.v)}"`)
  if (style.hanging !== undefined) attributes.push(`w:hanging="${twipsOf(style.hanging.v)}"`)
  return attributes.length === 0 ? '' : `<w:ind ${attributes.join(' ')}/>`
}

function writeTabStops(stops: readonly ITabStop[] | undefined): string {
  if (stops === undefined || stops.length === 0) return ''
  let xml = ''
  for (const stop of stops) {
    const value = stop.clear === true
      ? 'clear'
      : nameOf(TAB_STOP_ALIGNMENTS, stop.alignment) ?? 'left'
    xml += `<w:tab w:val="${value}" w:pos="${twipsOf(stop.offset)}"`
    const leader = stop.leader === undefined ? undefined : nameOf(TAB_STOP_LEADERS, stop.leader)
    if (leader !== undefined) xml += ` w:leader="${leader}"`
    xml += '/>'
  }
  return `<w:tabs>${xml}</w:tabs>`
}

function writeParagraphBorders(style: IParagraphStyle): string {
  let xml = ''
  xml += writeParagraphBorder('top', style.borderTop)
  xml += writeParagraphBorder('bottom', style.borderBottom)
  xml += writeParagraphBorder('left', style.borderLeft)
  xml += writeParagraphBorder('right', style.borderRight)
  xml += writeParagraphBorder('between', style.borderBetween)
  return xml.length === 0 ? '' : `<w:pBdr>${xml}</w:pBdr>`
}

function writeParagraphBorder(tag: string, border: IParagraphBorder | undefined): string {
  if (border === undefined) return ''
  const color = hexColor(border.color.rgb) ?? DEFAULT_BORDER_COLOR
  const dash = nameOf(BORDER_DASH_STYLES, border.dashStyle) ?? 'single'
  const size = Math.max(1, Math.round(border.width * BORDER_UNITS_PER_PIXEL))
  const space = Math.max(0, Math.round(border.padding * POINTS_PER_PIXEL))
  return `<w:${tag} w:val="${dash}" w:sz="${String(size)}" w:space="${String(space)}" w:color="${color}"/>`
}

function twipsOf(pixels: number): string {
  return String(Math.round(pixels * TWIPS_PER_PIXEL))
}

function styleIdForNamedStyle(namedStyle: NamedStyleType | undefined): string | undefined {
  if (namedStyle === undefined) return undefined
  return HEADING_STYLE_IDS[namedStyle - 4]
}

function writeRun(text: string, style: ITextStyle | undefined): string {
  const content = writeRunText(text)
  return content.length === 0 ? '' : `<w:r>${writeRunProperties(style)}${content}</w:r>`
}

function writeRunProperties(style: ITextStyle | undefined): string {
  if (style === undefined) return ''
  let properties = ''
  const family = style.ff ?? undefined
  const eastAsia = style.eastAsiaFontFamily ?? undefined
  if (family !== undefined || eastAsia !== undefined) {
    properties += '<w:rFonts'
    if (family !== undefined) properties += ` w:ascii="${escapeXml(family)}" w:hAnsi="${escapeXml(family)}"`
    if (eastAsia !== undefined) properties += ` w:eastAsia="${escapeXml(eastAsia)}"`
    properties += '/>'
  }
  if (style.bl === BOOLEAN_TRUE) properties += '<w:b/>'
  if (style.it === BOOLEAN_TRUE) properties += '<w:i/>'
  if (style.caps === true) properties += '<w:caps/>'
  if (style.smallCaps === true) properties += '<w:smallCaps/>'
  if (style.st?.s === BOOLEAN_TRUE) {
    properties += style.st.t === TEXT_DECORATION_DOUBLE ? '<w:dstrike/>' : '<w:strike/>'
  }
  const foreground = hexColor(style.cl?.rgb)
  if (foreground !== undefined) properties += `<w:color w:val="${foreground}"/>`
  if (style.fs !== undefined) {
    const halfPoints = String(Math.round(style.fs * HALF_POINTS_PER_POINT))
    properties += `<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>`
  }
  const background = hexColor(style.bg?.rgb)
  if (background !== undefined) {
    const highlight = highlightName(background)
    properties += highlight === undefined
      ? `<w:shd w:val="clear" w:color="auto" w:fill="${background}"/>`
      : `<w:highlight w:val="${highlight}"/>`
  }
  if (style.ul?.s === BOOLEAN_TRUE) {
    properties += `<w:u w:val="${underlineName(style.ul.t)}"`
    const underlineColor = hexColor(style.ul.cl?.rgb)
    if (underlineColor !== undefined) properties += ` w:color="${underlineColor}"`
    properties += '/>'
  }
  const vertical = style.va === null || style.va === undefined ? undefined : verticalAlignName(style.va)
  if (vertical !== undefined) properties += `<w:vertAlign w:val="${vertical}"/>`
  const border = style.bd?.t ?? style.bd?.r ?? style.bd?.b ?? style.bd?.l
  if (border !== null && border !== undefined) {
    const color = hexColor(border.cl.rgb) ?? DEFAULT_BORDER_COLOR
    properties += `<w:bdr w:val="${nameOf(BORDER_WEIGHTS, border.s) ?? 'single'}"`
      + ` w:sz="${String(RUN_BORDER_SIZE)}" w:space="1" w:color="${color}"/>`
  }
  return properties.length === 0 ? '' : `<w:rPr>${properties}</w:rPr>`
}

function underlineName(decoration: TextDecoration | undefined): string {
  return decoration === undefined ? 'single' : nameOf(UNDERLINE_VALUES, decoration) ?? 'single'
}

function verticalAlignName(vertical: BaselineOffset): string {
  if (vertical === BaselineOffset.SUPERSCRIPT) return 'superscript'
  if (vertical === BaselineOffset.SUBSCRIPT) return 'subscript'
  return 'baseline'
}

function highlightName(color: string): string | undefined {
  for (const [name, value] of Object.entries(HIGHLIGHT_COLORS)) {
    if (value === color) return name
  }
  return undefined
}

function writeRunText(text: string): string {
  let xml = ''
  let buffer = ''
  const flush = (): void => {
    if (buffer.length === 0) return
    xml += `<w:t xml:space="preserve">${escapeXml(buffer)}</w:t>`
    buffer = ''
  }
  for (const character of text) {
    if (character === TAB) {
      flush()
      xml += '<w:tab/>'
      continue
    }
    if (character.charCodeAt(0) >= 0x20) buffer += character
  }
  flush()
  return xml
}

function escapeXml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

/** XML parts a document read needs from the package. */
interface PackageParts {
  /** Decoded `word/document.xml`, when present. */
  readonly document: Uint8Array | undefined
  /** Decoded `word/styles.xml`, when present. */
  readonly styles: Uint8Array | undefined
  /** Decoded `word/settings.xml`, when present. */
  readonly settings: Uint8Array | undefined
  /** Decoded `word/numbering.xml`, when present. */
  readonly numbering: Uint8Array | undefined
  /** Decoded `word/_rels/document.xml.rels`, when present. */
  readonly relationships: Uint8Array | undefined
  /** Header parts keyed by their segment identity. */
  readonly headers: ReadonlyMap<string, Uint8Array>
  /** Footer parts keyed by their segment identity. */
  readonly footers: ReadonlyMap<string, Uint8Array>
  /** Header and footer relationship parts keyed by their segment identity. */
  readonly storyRelationships: ReadonlyMap<string, Uint8Array>
  /** Media entries keyed by their package entry name. */
  readonly media: ReadonlyMap<string, Uint8Array>
  /** Footnote and endnote parts keyed by their package entry name. */
  readonly notes: ReadonlyMap<string, Uint8Array>
  /** Footnote and endnote relationship parts keyed by their package entry name. */
  readonly noteRelationships: ReadonlyMap<string, Uint8Array>
  /** Decoded `[Content_Types].xml`, when present. */
  readonly contentTypes: Uint8Array | undefined
}

function readPackage(bytes: Uint8Array): PackageParts {
  let entryCount = 0
  let expandedBytes = 0
  let entries: Unzipped
  try {
    entries = unzipSync(bytes, {
      filter: (file) => {
        entryCount += 1
        if (entryCount > ENTITY_COUNT_LIMIT) {
          throw new OfficeExchangeError('too-large', `Document package contains more than ${ENTITY_COUNT_LIMIT} entries`)
        }
        expandedBytes += file.originalSize
        /* v8 ignore next 6 -- A 200 MB expansion is not testable; the entry guard covers this failure class. */
        if (expandedBytes > EXPANSION_BYTE_LIMIT) {
          throw new OfficeExchangeError(
            'too-large',
            `Document package expands beyond ${EXPANSION_BYTE_LIMIT} bytes`,
          )
        }
        const path = file.name.toLowerCase()
        return file.name === DOCUMENT_XML
          || file.name === STYLES_XML
          || file.name === SETTINGS_XML
          || file.name === NUMBERING_XML
          || file.name === DOCUMENT_RELS
          || file.name === CONTENT_TYPES_XML
          || MEDIA_PART.test(path)
          || NOTE_PART.test(path)
          || NOTE_RELS_PART.test(path)
          || HEADER_PART.test(path)
          || FOOTER_PART.test(path)
          || STORY_RELS_PART.test(path)
      },
    })
  } catch (error) {
    if (error instanceof OfficeExchangeError) throw error
    throw new OfficeExchangeError('conversion-failed', `Cannot read OOXML package: ${messageOf(error)}`)
  }
  const headers = new Map<string, Uint8Array>()
  const footers = new Map<string, Uint8Array>()
  const storyRelationships = new Map<string, Uint8Array>()
  const media = new Map<string, Uint8Array>()
  const notes = new Map<string, Uint8Array>()
  const noteRelationships = new Map<string, Uint8Array>()
  for (const [name, content] of Object.entries(entries)) {
    const path = name.toLowerCase()
    if (HEADER_PART.test(path)) headers.set(partIdOf(name), content)
    if (FOOTER_PART.test(path)) footers.set(partIdOf(name), content)
    if (STORY_RELS_PART.test(path)) storyRelationships.set(storyIdOf(name), content)
    if (MEDIA_PART.test(path)) media.set(name, content)
    if (NOTE_PART.test(path)) notes.set(name, content)
    if (NOTE_RELS_PART.test(path)) noteRelationships.set(name, content)
  }
  return {
    document: entries[DOCUMENT_XML],
    styles: entries[STYLES_XML],
    settings: entries[SETTINGS_XML],
    numbering: entries[NUMBERING_XML],
    relationships: entries[DOCUMENT_RELS],
    headers,
    footers,
    storyRelationships,
    media,
    notes,
    noteRelationships,
    contentTypes: entries[CONTENT_TYPES_XML],
  }
}

/** Header/footer identity from its relationship part name (`word/_rels/header1.xml.rels`). */
function storyIdOf(partName: string): string {
  const name = partName.slice(partName.lastIndexOf('/') + 1)
  return name.slice(0, -'.xml.rels'.length)
}

/**
 * Collect package state the editable model cannot address but the writer must re-emit.
 * @param parts - Decoded package entries read for one document.
 * @returns Preserved parts, relationships, and content types, or `undefined` when none exist.
 */
function collectPreservedPackage(parts: PackageParts): DocxPreservedPackage | undefined {
  const preservedParts: PreservedPart[] = []
  for (const [name, content] of [...parts.media, ...parts.notes, ...parts.noteRelationships]) {
    preservedParts.push({ name, base64: encodeBase64(content) })
  }
  preservedParts.sort((left, right) => left.name < right.name ? -1 : 1)
  const documentRelationships = preservedRelationshipElements(parts.relationships, isImageOrNoteRelationship)
  const stories: Record<string, readonly string[]> = {}
  for (const [id, xml] of parts.storyRelationships) {
    const elements = preservedRelationshipElements(xml, isImageRelationship)
    if (elements.length > 0) stories[id] = elements
  }
  const contentTypes = preservedContentTypeElements(parts, preservedParts)
  if (preservedParts.length === 0 && documentRelationships.length === 0
    && Object.keys(stories).length === 0 && contentTypes.length === 0) {
    return undefined
  }
  return {
    parts: preservedParts,
    relationships: {
      ...(documentRelationships.length === 0 ? {} : { document: documentRelationships }),
      ...(Object.keys(stories).length === 0 ? {} : { stories }),
    },
    contentTypes,
  }
}

function isImageRelationship(element: string): boolean {
  return (tagAttribute(element, 'Type') ?? '').endsWith('/image')
}

function isImageOrNoteRelationship(element: string): boolean {
  const type = tagAttribute(element, 'Type') ?? ''
  return type.endsWith('/image') || type.endsWith('/footnotes') || type.endsWith('/endnotes')
}

/**
 * Extract relationship elements that match a predicate without re-encoding them.
 * @param xml - Decoded relationship part, when present.
 * @param keep - Predicate over one complete element.
 * @returns Matching element source strings in source order.
 */
function preservedRelationshipElements(
  xml: Uint8Array | undefined,
  keep: (element: string) => boolean,
): readonly string[] {
  if (xml === undefined) return []
  const elements: string[] = []
  for (const match of strFromU8(xml).matchAll(/<Relationship(?=[\s/>])[^>]*\/?>/gu)) {
    if (keep(match[0])) elements.push(match[0])
  }
  return elements
}

/**
 * Keep content-type declarations the generated manifest does not reproduce.
 * @param parts - Decoded source package.
 * @param preservedParts - Entries the writer re-emits.
 * @returns Declaration elements for preserved entries and their extensions.
 */
function preservedContentTypeElements(
  parts: PackageParts,
  preservedParts: readonly PreservedPart[],
): readonly string[] {
  if (parts.contentTypes === undefined) return []
  const names = new Set(preservedParts.map(part => `/${part.name}`))
  const extensions = new Set(preservedParts.map(extensionOf).filter(name => name.length > 0))
  const elements: string[] = []
  const text = strFromU8(parts.contentTypes)
  for (const match of text.matchAll(/<(Default|Override)(?=[\s/>])[^>]*\/?>/gu)) {
    const element = match[0]
    if (match[1] === 'Default') {
      const extension = tagAttribute(element, 'Extension')?.toLowerCase()
      if (extension !== undefined && extension !== 'rels' && extension !== 'xml'
        && extensions.has(extension)) {
        elements.push(element)
      }
      continue
    }
    const partName = tagAttribute(element, 'PartName')
    if (partName !== undefined && names.has(partName)) elements.push(element)
  }
  return elements
}

function extensionOf(part: PreservedPart): string {
  const name = part.name
  const dot = name.lastIndexOf('.')
  return dot < 0 ? '' : name.slice(dot + 1).toLowerCase()
}

const BASE64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'

/**
 * Encode bytes as base64 for a JSON-safe snapshot field.
 * @param bytes - Source bytes.
 * @returns Standard padded base64 text.
 */
function encodeBase64(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString('base64')
}

/**
 * Decode base64 carried by a snapshot back to bytes.
 * @param text - Standard padded base64 text.
 * @returns Decoded bytes.
 * @throws {OfficeExchangeError} `conversion-failed` when the text is not valid base64.
 */
function decodeBase64(text: string): Uint8Array {
  if (text.length % 4 !== 0) throw new OfficeExchangeError('conversion-failed', 'Preserved part is not valid base64')
  const bytes: number[] = []
  for (let index = 0; index < text.length; index += 4) {
    const values: [number, number, number, number] = [0, 0, 0, 0]
    let padding = 0
    for (let offset = 0; offset < 4; offset += 1) {
      const character = text.charAt(index + offset)
      if (character === '=') {
        padding += 1
        continue
      }
      const value = BASE64_ALPHABET.indexOf(character)
      if (value < 0 || padding > 0) {
        throw new OfficeExchangeError('conversion-failed', 'Preserved part is not valid base64')
      }
      values[offset] = value
    }
    bytes.push((values[0] << 2) | (values[1] >> 4))
    if (padding < 2) bytes.push(((values[1] & 0x0F) << 4) | (values[2] >> 2))
    if (padding < 1) bytes.push(((values[2] & 0x03) << 6) | values[3])
    if (padding > 0 && index + 4 !== text.length) {
      throw new OfficeExchangeError('conversion-failed', 'Preserved part is not valid base64')
    }
  }
  return Uint8Array.from(bytes)
}

/**
 * Parse one XML part and return its root element.
 * @param bytes - Encoded part XML.
 * @param tag - Expected local root element name.
 * @param description - Part name used in the failure message.
 * @returns The parsed root element.
 * @throws {OfficeExchangeError} `conversion-failed` when the XML is malformed or lacks the root element.
 */
function parseRootElement(bytes: Uint8Array, tag: string, description: string): XmlElement {
  const xml = strFromU8(bytes)
  let parsed: readonly ParserNode[]
  try {
    SyntaxValidator.validate(xml)
    parsed = new XMLParser(XML_PARSER_OPTIONS).parse(xml) as readonly ParserNode[]
  } catch (error) {
    throw new OfficeExchangeError('conversion-failed', `Cannot read ${description}: ${messageOf(error)}`)
  }
  const root = parsed.map(toElement).find(element => element?.tag === tag)
  if (root === undefined) {
    throw new OfficeExchangeError('conversion-failed', `${description} has no w:${tag} element`)
  }
  return root
}

function parseDocumentBody(documentXml: Uint8Array): XmlElement {
  const root = parseRootElement(documentXml, 'document', 'document XML')
  const body = findChild(root, 'body')
  if (body === undefined) {
    throw new OfficeExchangeError('conversion-failed', 'Document XML has no w:body element')
  }
  return body
}

/**
 * Read the document relationship part.
 * @param xml - Decoded `word/_rels/document.xml.rels`, when present.
 * @returns Relationship ids mapped to their type and target; empty without a part.
 */
function readRelationships(xml: Uint8Array | undefined): DocumentRelationships {
  const relationships = new Map<string, DocumentRelationship>()
  if (xml === undefined) return relationships
  const text = strFromU8(xml)
  for (const match of text.matchAll(/<Relationship(?=[\s/>])[^>]*>/gu)) {
    const id = tagAttribute(match[0], 'Id')
    const type = tagAttribute(match[0], 'Type')
    const target = tagAttribute(match[0], 'Target')
    if (id !== undefined && type !== undefined && target !== undefined) {
      relationships.set(id, { type, target })
    }
  }
  return relationships
}

function tagAttribute(tag: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(tag)?.[1]
}

/**
 * Report whether `word/settings.xml` enables even and odd page headers.
 * @param xml - Decoded settings part, when present.
 * @returns `true` when the on/off element is present and not disabled.
 */
function readEvenAndOddHeaders(xml: Uint8Array | undefined): boolean {
  if (xml === undefined) return false
  const match = /<(?:w:)?evenAndOddHeaders(?=[\s/>])[^>]*>/u.exec(strFromU8(xml))
  if (match === null) return false
  const value = tagAttribute(match[0], 'w:val') ?? tagAttribute(match[0], 'val')
  return value === undefined || (value !== '0' && value !== 'false' && value !== 'off')
}

function partIdOf(partName: string): string {
  const name = partName.slice(partName.lastIndexOf('/') + 1)
  return name.slice(0, -'.xml'.length)
}

function toElement(node: ParserNode): XmlElement | undefined {
  for (const key of Object.keys(node)) {
    if (key === ':@' || key === '#text') continue
    const children = node[key]
    /* v8 ignore next -- Every non-attribute parser key maps to an element's child array. */
    if (!isParserNodeArray(children)) continue
    const attributes = node[':@']
    return { tag: key, children, attributes: isParserNode(attributes) ? attributes : {} }
  }
  return undefined
}

function childElements(element: XmlElement): readonly XmlElement[] {
  const elements: XmlElement[] = []
  for (const child of element.children) {
    const converted = toElement(child)
    if (converted !== undefined) elements.push(converted)
  }
  return elements
}

function findChild(element: XmlElement, tag: string): XmlElement | undefined {
  for (const child of element.children) {
    const converted = toElement(child)
    if (converted?.tag === tag) return converted
  }
  return undefined
}

function textContent(element: XmlElement): string {
  let text = ''
  for (const child of element.children) {
    const value = child['#text']
    /* v8 ignore next -- A w:t element contains only text children. */
    if (typeof value === 'string') text += value
  }
  return text
}

function attribute(element: XmlElement, name: string): string | undefined {
  const value = element.attributes[name]
  return typeof value === 'string' ? value : undefined
}

function optionalAttribute(element: XmlElement | undefined, name: string): string | undefined {
  return element === undefined ? undefined : attribute(element, name)
}

function numberAttribute(element: XmlElement | undefined, name: string): number | undefined {
  const value = optionalAttribute(element, name)
  if (value === undefined) return undefined
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : undefined
}

function lookup<Value>(table: readonly (readonly [string, Value])[], name: string): Value | undefined {
  for (const [key, value] of table) {
    if (key === name) return value
  }
  return undefined
}

function nameOf<Value>(table: readonly (readonly [string, Value])[], value: Value): string | undefined {
  for (const [key, candidate] of table) {
    if (candidate === value) return key
  }
  return undefined
}

function hexColor(value: string | null | void): string | undefined {
  if (typeof value !== 'string') return undefined
  const hex = value.startsWith('#') ? value.slice(1) : value
  return /^[0-9A-Fa-f]{6}$/u.test(hex) ? hex.toUpperCase() : undefined
}

function rgbOf(element: XmlElement | undefined): string | undefined {
  const color = hexColor(optionalAttribute(element, '@_val'))
  return color === undefined ? undefined : `#${color}`
}

function shadingColor(element: XmlElement | undefined): string | undefined {
  const color = hexColor(optionalAttribute(element, '@_fill'))
  return color === undefined ? undefined : `#${color}`
}

function isParserNode(value: unknown): value is ParserNode {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function isParserNodeArray(value: unknown): value is readonly ParserNode[] {
  return Array.isArray(value) && value.every(isParserNode)
}

function positiveNumber(value: string | undefined): number | undefined {
  if (value === undefined) return undefined
  const parsed = Number(value)
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined
}

function documentXml(body: string, section: string): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    + `${PRESERVED_NAMESPACES}>`
    + `<w:body>${body}${section}</w:body>`
    + '</w:document>'
}

function contentTypesXml(
  headers: readonly StoryPart[],
  footers: readonly StoryPart[],
  settings: boolean,
  numbering: boolean,
  preserved: DocxPreservedPackage | undefined,
): string {
  let overrides = ''
  if (settings) {
    overrides += '<Override PartName="/word/settings.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
  }
  if (numbering) {
    overrides += '<Override PartName="/word/numbering.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
  }
  for (const part of headers) {
    overrides += `<Override PartName="/${part.partName}"`
      + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
  }
  for (const part of footers) {
    overrides += `<Override PartName="/${part.partName}"`
      + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
  }
  const preservedDeclarations = preserved?.contentTypes ?? []
  const noteOverrides = preservedNoteOverrides(preservedDeclarations, preserved)
  const mediaDefaults = preservedContentTypeDeclarations(preservedDeclarations, preserved)
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + mediaDefaults
    + '<Override PartName="/word/document.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + `${overrides}${noteOverrides}</Types>`
}

function preservedNoteOverrides(
  declarations: readonly string[],
  preserved: DocxPreservedPackage | undefined,
): string {
  let overrides = ''
  for (const [name, contentType] of Object.entries(NOTE_CONTENT_TYPES)) {
    const present = preserved?.parts.some(part => part.name === name) === true
    if (!present) continue
    const declaration = declarations.find(candidate => tagAttribute(candidate, 'PartName') === `/${name}`)
    overrides += declaration ?? `<Override PartName="/${name}" ContentType="${contentType}"/>`
  }
  return overrides
}

function preservedContentTypeDeclarations(
  declarations: readonly string[],
  preserved: DocxPreservedPackage | undefined,
): string {
  const declared = new Set<string>()
  let defaults = ''
  for (const declaration of declarations) {
    const extension = tagAttribute(declaration, 'Extension')
    if (extension === undefined) {
      const partName = tagAttribute(declaration, 'PartName')
      if (partName !== undefined && !notePartNameOf(partName)) defaults += declaration
      continue
    }
    declared.add(extension.toLowerCase())
    defaults += declaration
  }
  const missing = new Map<string, string>()
  for (const part of preserved?.parts ?? []) {
    const extension = extensionOf(part)
    const contentType = IMAGE_CONTENT_TYPES[extension]
    if (contentType !== undefined && !declared.has(extension)) missing.set(extension, contentType)
  }
  for (const [extension, contentType] of [...missing].sort()) {
    defaults += `<Default Extension="${extension}" ContentType="${contentType}"/>`
  }
  return defaults
}

function notePartNameOf(partName: string): boolean {
  return Object.keys(NOTE_CONTENT_TYPES).some(name => `/${name}` === partName)
}

function rootRelsXml(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1"'
    + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
    + ' Target="word/document.xml"/>'
    + '</Relationships>'
}

function documentRelsXml(
  stylesRelId: string,
  settingsRelId: string | undefined,
  numberingRelId: string | undefined,
  headers: readonly StoryPart[],
  footers: readonly StoryPart[],
  links: readonly LinkRel[],
  preserved: readonly string[],
): string {
  let relationships = `<Relationship Id="${stylesRelId}"`
    + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
    + ' Target="styles.xml"/>'
  if (settingsRelId !== undefined) {
    relationships += `<Relationship Id="${settingsRelId}"`
      + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'
      + ' Target="settings.xml"/>'
  }
  if (numberingRelId !== undefined) {
    relationships += `<Relationship Id="${numberingRelId}"`
      + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"'
      + ' Target="numbering.xml"/>'
  }
  for (const part of headers) {
    relationships += `<Relationship Id="${part.relId}"`
      + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"'
      + ` Target="${part.partName.slice('word/'.length)}"/>`
  }
  for (const part of footers) {
    relationships += `<Relationship Id="${part.relId}"`
      + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"'
      + ` Target="${part.partName.slice('word/'.length)}"/>`
  }
  for (const rel of links) {
    relationships += `<Relationship Id="${rel.relId}"`
      + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'
      + ` Target="${escapeXml(rel.target)}" TargetMode="External"/>`
  }
  relationships += preserved.join('')
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + `${relationships}</Relationships>`
}

function stylesXml(documentStyle: IDocumentStyle): string {
  const headingStyles = HEADING_STYLE_IDS.map((styleId, index) => {
    const level = index + 1
    return `<w:style w:type="paragraph" w:styleId="${styleId}">`
      + `<w:name w:val="heading ${String(level)}"/>`
      + '<w:basedOn w:val="Normal"/>'
      + '<w:next w:val="Normal"/>'
      + '<w:qFormat/>'
      + `<w:pPr><w:outlineLvl w:val="${String(index)}"/></w:pPr>`
      + '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>'
      + '</w:style>'
  }).join('')
  const textStyle = writeRunProperties(documentStyle.textStyle ?? DEFAULT_TEXT_STYLE)
  const paragraphStyle = documentStyle.defaultParagraphStyle === undefined
    ? ''
    : writeParagraphFormatting(documentStyle.defaultParagraphStyle)
  const paragraphDefault = paragraphStyle.length === 0
    ? ''
    : `<w:pPrDefault><w:pPr>${paragraphStyle}</w:pPr></w:pPrDefault>`
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + `<w:docDefaults><w:rPrDefault>${textStyle}</w:rPrDefault>${paragraphDefault}</w:docDefaults>`
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    + headingStyles
    + '</w:styles>'
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
