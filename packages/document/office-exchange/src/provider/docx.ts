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
 * PRESERVED OPAQUELY: each run that carries a `w:drawing` or `w:pict` and each
 * run that carries a `w:footnoteReference` or `w:endnoteReference` is captured
 * byte-for-byte in `body.docxRawBlocks`, together with the source media parts,
 * note parts, relationships, and content-type declarations. The writer
 * re-emits them unchanged, so images and note bodies survive an edit-and-save
 * without becoming editable or rendered content. A raw block anchors at the
 * stream offset before the run and adds no stream character, so paragraph and
 * run offsets are unchanged.
 *
 * UNSUPPORTED: charts, drawing parts, shapes, comments, tracked changes,
 * macros, page borders and page numbering, theme styles and theme-derived
 * colors, named style definitions other than the heading styles above,
 * floating objects with their own drawing parts, and property values the
 * Univer model has no field for (for example a run border's line size or a
 * decimal tab stop's alignment). A read/write round-trip does not preserve
 * these. Macro-bearing packages are read without executing or retaining
 * macros.
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
  readonly rawBlocks: IDocxRawBlock[]
  readonly bounds: OfficeExchangeConfig
  cells: number
  tableSerial: number
}

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
  }
}

/** One XML run captured from a source part for lossless re-emission. */
interface RawRunRecord {
  /** Exact source XML from `<w:r` through `</w:r>`. */
  readonly xml: string
  /** Relationship and note identity values used to match the parsed run. */
  readonly key: string
  taken: boolean
}

/** Source runs that a parsed run may consume once, matched by identity or order. */
interface RawRunSource {
  /**
   * Take the source XML for one parsed run.
   * @param element - Parsed run that carries a drawing, picture, or note reference.
   * @returns The exact source run XML, or `undefined` when none remains.
   */
  take(element: XmlElement): string | undefined
}

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
 * hyperlinks, and opaque image and note preservation are mapped; the
 * unsupported features listed in the module note are not.
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
  return Promise.resolve().then(() => readDocumentData(bytes, config))
}

function readDocumentData(bytes: Uint8Array, config?: OfficeExchangeConfig): IDocumentData {
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
  const context = createReadContext(bounds, readRawRunCatalog(strFromU8(parts.document)))
  for (const child of childElements(body)) {
    if (child.tag === 'p') appendParagraph(context, child, relationships, identifiers, numbering)
    if (child.tag === 'tbl') appendTable(context, child, relationships, numbering)
  }
  const finalSection = findChild(body, 'sectPr')
  finalizeBody(context, finalSection === undefined ? undefined : readSectionSetup(finalSection, relationships, identifiers))
  const storyRelationships = readStoryRelationships(parts.storyRelationships)
  const headers = readStoryParts(parts.headers, 'hdr', bounds, numbering, storyRelationships, (id, body, tableSource) => (
    { headerId: id, body, tableSource }
  ))
  const footers = readStoryParts(parts.footers, 'ftr', bounds, numbering, storyRelationships, (id, body, tableSource) => (
    { footerId: id, body, tableSource }
  ))
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
  return snapshot
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

function createReadContext(bounds: OfficeExchangeConfig, rawRuns?: RawRunSource): ReadContext {
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
    rawBlocks: [],
    bounds,
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
): void {
  append(context, PARAGRAPH)
  const paragraph: IParagraph = {
    startIndex: context.length - 1,
    paragraphId: `paragraph-${context.paragraphs.length + 1}`,
  }
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
 * Map `w:numPr` into the paragraph's list bullet.
 * @param properties - Paragraph properties containing `w:numPr`.
 * @param numbering - Numbering definitions from the package.
 * @returns The paragraph bullet, or `undefined` for an unnumbered paragraph.
 */
function readBullet(properties: XmlElement, numbering: NumberingCatalog): IBullet | undefined {
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
  'footnoteReference',
  'endnoteReference',
])

/**
 * Report whether a run carries content preserved only as raw XML.
 * @param element - Parsed run.
 * @returns `true` for a drawing, picture, or footnote/endnote reference anywhere inside the run.
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
  const records: RawRunRecord[] = []
  const byKey = new Map<string, RawRunRecord[]>()
  for (const match of xml.matchAll(/<w:r(?=[\s/>])[^>]*>/gu)) {
    if (match[0].endsWith('/>')) continue
    const start = match.index
    const end = runEndOf(xml, start)
    if (end === undefined) continue
    const run = xml.slice(start, end)
    if (!containsPreservedMarkup(run)) continue
    const key = relationshipKeyOfXml(run)
    const record: RawRunRecord = { xml: run, key, taken: false }
    records.push(record)
    const queue = byKey.get(key)
    if (queue === undefined) byKey.set(key, [record])
    else queue.push(record)
  }
  let cursor = 0
  return {
    take(element) {
      const key = relationshipKeyOfElement(element)
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
  return undefined
}

/**
 * Report whether source run XML carries preserved content.
 * @param xml - One source run.
 * @returns `true` when it contains a drawing, picture, or note reference.
 */
function containsPreservedMarkup(xml: string): boolean {
  return /<w:(?:drawing|pict|footnoteReference|endnoteReference)(?=[\s/>])/u.test(xml)
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
  if (lastCharacter(context) !== PARAGRAPH) appendParagraphBreak(context)
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
      if (lastCharacter(context) !== PARAGRAPH) appendParagraphBreak(context)
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
  if (lastCharacter(context) !== PARAGRAPH) appendParagraphBreak(context)
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
 * @returns Segment identity mapped to its entry.
 */
function readStoryParts<Entry>(
  parts: ReadonlyMap<string, Uint8Array>,
  root: 'hdr' | 'ftr',
  bounds: OfficeExchangeConfig,
  numbering: NumberingCatalog,
  relationships: ReadonlyMap<string, DocumentRelationships>,
  build: (id: string, body: IDocumentBody, tableSource: ITables) => Entry,
): Map<string, Entry> {
  const stories = new Map<string, Entry>()
  for (const [id, xml] of parts) {
    const element = parseRootElement(xml, root, `${root === 'hdr' ? 'header' : 'footer'} XML`)
    const context = createReadContext(bounds, readRawRunCatalog(strFromU8(xml)))
    const storyRelationships = relationships.get(id)
    for (const child of childElements(element)) {
      if (child.tag === 'p') appendParagraph(context, child, storyRelationships, undefined, numbering)
      if (child.tag === 'tbl') appendTable(context, child, storyRelationships, numbering)
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
    stories.set(id, build(id, body, context.tableSource))
  }
  return stories
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
    const color = hexColor(border.cl?.rgb) ?? DEFAULT_BORDER_COLOR
    properties += `<w:bdr w:val="${nameOf(BORDER_WEIGHTS, border.s) ?? 'single'}"`
      + ` w:sz="${String(RUN_BORDER_SIZE)}" w:space="1" w:color="${color}"/>`
  }
  return properties.length === 0 ? '' : `<w:rPr>${properties}</w:rPr>`
}

function underlineName(decoration: TextDecoration | undefined): string {
  return decoration === undefined ? 'single' : nameOf(UNDERLINE_VALUES, decoration) ?? 'single'
}

function verticalAlignName(vertical: BaselineOffset): string | undefined {
  if (vertical === BaselineOffset.SUPERSCRIPT) return 'superscript'
  if (vertical === BaselineOffset.SUBSCRIPT) return 'subscript'
  if (vertical === BaselineOffset.NORMAL) return 'baseline'
  return undefined
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
  preservedParts.sort((left, right) => left.name < right.name ? -1 : left.name > right.name ? 1 : 0)
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
  let text = ''
  for (let index = 0; index < bytes.length; index += 3) {
    const first = bytes[index] ?? 0
    const second = bytes[index + 1]
    const third = bytes[index + 2]
    text += BASE64_ALPHABET[first >> 2] ?? ''
    text += BASE64_ALPHABET[((first & 0x03) << 4) | ((second ?? 0) >> 4)] ?? ''
    text += second === undefined ? '=' : BASE64_ALPHABET[((second & 0x0F) << 2) | ((third ?? 0) >> 6)] ?? ''
    text += third === undefined ? '=' : BASE64_ALPHABET[third & 0x3F] ?? ''
  }
  return text
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
    const values = [0, 0, 0, 0]
    let padding = 0
    for (let offset = 0; offset < 4; offset += 1) {
      const character = text[index + offset]
      if (character === '=') {
        padding += 1
        continue
      }
      const value = character === undefined ? -1 : BASE64_ALPHABET.indexOf(character)
      if (value < 0 || padding > 0) {
        throw new OfficeExchangeError('conversion-failed', 'Preserved part is not valid base64')
      }
      values[offset] = value
    }
    bytes.push(((values[0] ?? 0) << 2) | ((values[1] ?? 0) >> 4))
    if (padding < 2) bytes.push((((values[1] ?? 0) & 0x0F) << 4) | ((values[2] ?? 0) >> 2))
    if (padding < 1) bytes.push((((values[2] ?? 0) & 0x03) << 6) | (values[3] ?? 0))
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
    const declared = declarations.some(declaration => tagAttribute(declaration, 'PartName') === `/${name}`)
    const present = preserved?.parts.some(part => part.name === name) === true
    if (!present) continue
    overrides += declared
      ? declarations.find(declaration => tagAttribute(declaration, 'PartName') === `/${name}`) ?? ''
      : `<Override PartName="/${name}" ContentType="${contentType}"/>`
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
  const missing = new Set<string>()
  for (const part of preserved?.parts ?? []) {
    const extension = extensionOf(part)
    if (IMAGE_CONTENT_TYPES[extension] !== undefined && !declared.has(extension)) missing.add(extension)
  }
  for (const extension of [...missing].sort()) {
    defaults += `<Default Extension="${extension}" ContentType="${IMAGE_CONTENT_TYPES[extension] ?? 'application/octet-stream'}"/>`
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
