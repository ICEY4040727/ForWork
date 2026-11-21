/** OOXML package inspection for features the editable model cannot preserve. */
import { strFromU8, strToU8, unzipSync } from 'fflate'
import type { Unzipped } from 'fflate'
import type { OfficeExchangeConfig } from './config.ts'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from './config.ts'
import { OfficeExchangeError } from './errors.ts'
import type {
  CoupledConstructMarker,
  PreservedConstructKind,
  PreservedConstructs,
  UnsupportedFeature,
  UnsupportedFeatureKind,
} from './types.ts'

const DOCUMENT_XML = 'word/document.xml'

const MACRO_PATHS: ReadonlySet<string> = new Set([
  'word/vbaproject.bin',
  'xl/vbaproject.bin',
  'ppt/vbaproject.bin',
])

const DRAWING_PREFIXES: readonly string[] = [
  'word/drawings/',
  'ppt/drawings/',
]

/** Workbook drawing parts inspected for constructs the image mapping does not rebuild. */
const WORKSHEET_DRAWING_PREFIX = 'xl/drawings/'

/** Drawing elements that are not cell pictures the sheet-drawing mapping preserves. */
const UNMAPPED_DRAWING_ELEMENT = /<(?:[A-Za-z0-9]+:)?(?:sp|cxnSp|grpSp|graphicFrame|absoluteAnchor)(?=[\s/>])/u

const CHART_PREFIXES: readonly string[] = ['xl/charts/', 'ppt/charts/']

const EMBEDDING_PREFIXES: readonly string[] = [
  'word/embeddings/',
  'xl/embeddings/',
  'ppt/embeddings/',
]

const UNSUPPORTED_KIND_ORDER: readonly UnsupportedFeatureKind[] = [
  'macro',
  'tracked-changes',
  'comment',
  'chart',
  'drawing',
  'shape',
  'embedded-object',
]

/** Construct classes reported by {@link detectPreserved}, in report order. */
const PRESERVED_KIND_ORDER: readonly PreservedConstructKind[] = [
  'macro',
  'chart',
  'pivot-table',
  'excel-table',
  'defined-names',
  'calc-chain',
  'comment',
  'embedded-object',
  'drawing',
  'shape',
  'theme',
  'font-table',
  'web-settings',
  'custom-xml',
  'external-link',
  'field',
  'bookmark',
  'content-control',
  'document-properties',
]

/** Chart parts across containers; a chart's workbook is carried with it. */
const CHART_PART_PREFIXES: readonly string[] = [
  'word/charts/',
  'xl/charts/',
  'ppt/charts/',
]

/** Drawing parts across containers. */
const DRAWING_PART_PREFIXES: readonly string[] = [
  'word/drawings/',
  'xl/drawings/',
  'ppt/drawings/',
]

/** Theme parts across containers. */
const THEME_PART_PREFIXES: readonly string[] = [
  'word/theme/',
  'xl/theme/',
  'ppt/theme/',
]

/** Part trees that declare an external workbook link. */
const EXTERNAL_LINK_PREFIXES: readonly string[] = [
  'xl/externallinks/',
  'word/externallinks/',
]

/** WordprocessingML revision elements; any one marks a tracked change. */
const TRACKED_CHANGE_TAGS: readonly string[] = [
  'w:ins',
  'w:del',
  'w:moveFrom',
  'w:moveTo',
  'w:rPrChange',
  'w:pPrChange',
  'w:tblPrChange',
  'w:trPrChange',
  'w:tcPrChange',
  'w:sectPrChange',
]

/** Comment anchors in a story part; an edit over one orphanes the comment. */
const COMMENT_RANGE_TAGS: readonly string[] = [
  'w:commentRangeStart',
  'w:commentRangeEnd',
  'w:commentReference',
]

/** Field markup whose result text the field instruction owns. */
const FIELD_TAGS: readonly string[] = [
  'w:fldSimple',
  'w:fldChar',
]

/** Bookmark anchors, carried verbatim when the edit does not remove them. */
const BOOKMARK_TAGS: readonly string[] = [
  'w:bookmarkStart',
  'w:bookmarkEnd',
]

/** WordprocessingML story parts that carry body-level constructs. */
const WORD_STORY_PATHS: ReadonlySet<string> = new Set([
  'word/document.xml',
  'word/footnotes.xml',
  'word/endnotes.xml',
])

/** Matches the `t` attribute that selects a shared or array formula. */
const FORMULA_TYPE = /\bt\s*=\s*["'](shared|array)["']/u

/** Byte value of `>`, ending a start tag. */
const GREATER_THAN = 0x3e

/** Byte value of `/`, marking a self-closing start tag. */
const SLASH = 0x2f

/** Package entries a plausible document can contain before the archive is treated as hostile. */
const MAX_ENTRY_COUNT = 10_000

/** Bytes a plausible package can expand to, guarding against decompression bombs. */
const MAX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024

/** One decoded OOXML part inspected for construct-level losses. */
interface PackagePart {
  /** Entry name as stored in the archive. */
  readonly name: string
  /** Decoded part XML. */
  readonly xml: string
}

/** Entry inventory plus the decoded entries a scan selected. */
interface PackageScan {
  /** Entry names in central-directory order. */
  readonly entryNames: readonly string[]
  /** Decoded selected entries, keyed by their stored name. */
  readonly files: Unzipped
}

/** Selection predicate over one entry's stored name and its lowercased path. */
type EntrySelector = (name: string, path: string) => boolean

/** Everything one package scan classifies for the editor. */
export interface OfficePackageDetection {
  /** At most one unsupported feature per kind, in the fixed kind order. */
  readonly unsupported: readonly UnsupportedFeature[]
  /** Construct classes carried verbatim, in the fixed preserved kind order. */
  readonly preserved: ReadonlySet<PreservedConstructKind>
  /** Coupled-construct occurrences, sorted by part name then offset. */
  readonly coupled: readonly CoupledConstructMarker[]
}

/**
 * Detect unsupported features and preserved construct classes in one package read.
 *
 * The Exchange read path classifies both reports from a single package scan, so
 * an open inflates the source once for detection instead of running a second
 * classification pass.
 * @param bytes - Complete OOXML file bytes (`.docx`, `.docm`, `.xlsx`, or `.xlsm`).
 * @param config - Bounds applied before parsing; defaults when omitted.
 * @returns The unsupported features, preserved classes, and coupled markers.
 * @throws OfficeExchangeError `too-large` when a bound is exceeded, `conversion-failed` when the input is not a readable zip archive.
 */
export function detectOfficePackage(
  bytes: Uint8Array,
  config: OfficeExchangeConfig = DEFAULT_OFFICE_EXCHANGE_CONFIG,
): OfficePackageDetection {
  assertPackageSize(bytes, config)
  const { entryNames, files } = scanPackage(
    bytes,
    (_name, path) => isPreservedScanPart(path) || isDrawingPart(path),
  )
  return {
    unsupported: unsupportedFeaturesOf(entryNames, files),
    ...preservedConstructsOf(entryNames, files),
  }
}

/**
 * Detect features in an OOXML package that an edit-and-save would drop.
 * @param bytes - Complete OOXML file bytes (`.docx`, `.xlsx`, or `.pptx`).
 * @param config - Bounds applied before parsing; defaults when omitted.
 * @returns At most one feature per kind, ordered by the fixed kind order.
 * @throws OfficeExchangeError `too-large` when a bound is exceeded, `conversion-failed` when the input is not a readable zip archive.
 */
export function detectUnsupportedFeatures(
  bytes: Uint8Array,
  config: OfficeExchangeConfig = DEFAULT_OFFICE_EXCHANGE_CONFIG,
): readonly UnsupportedFeature[] {
  assertPackageSize(bytes, config)
  const { entryNames, files } = scanPackage(
    bytes,
    (name, path) => name === DOCUMENT_XML || isDrawingPart(path) || isHeaderOrFooterPart(path),
  )
  return unsupportedFeaturesOf(entryNames, files)
}

/**
 * Classify an OOXML package for an in-place save: which construct classes are
 * carried verbatim, and where coupled constructs sit so the caller can test
 * them against edited unit spans. Presence in either set never refuses a save.
 * @param bytes - Complete OOXML file bytes (`.docx`, `.docm`, `.xlsx`, or `.xlsm`).
 * @param config - Bounds applied before parsing; defaults when omitted.
 * @returns Preserved construct classes and coupled markers, each reported once per occurrence.
 * @throws OfficeExchangeError `too-large` when a bound is exceeded, `conversion-failed` when the input is not a readable zip archive.
 */
export function detectPreserved(
  bytes: Uint8Array,
  config: OfficeExchangeConfig = DEFAULT_OFFICE_EXCHANGE_CONFIG,
): PreservedConstructs {
  assertPackageSize(bytes, config)
  const { entryNames, files } = scanPackage(bytes, (_name, path) => isPreservedScanPart(path))
  return preservedConstructsOf(entryNames, files)
}

/**
 * Enforce the source byte bound every detection entry point shares.
 * @param bytes - Complete package bytes.
 * @param config - Detection bounds.
 * @throws OfficeExchangeError `too-large` when the archive exceeds the bound.
 */
function assertPackageSize(bytes: Uint8Array, config: OfficeExchangeConfig): void {
  if (bytes.byteLength > config.maxFileBytes) {
    throw new OfficeExchangeError(
      'too-large',
      `Office file is ${bytes.byteLength} bytes, exceeding maxFileBytes ${config.maxFileBytes}`,
    )
  }
}

/**
 * Report the features a whole-file regeneration would drop from one scanned package.
 * @param entryNames - Entry inventory in central-directory order.
 * @param files - Entries the scan decoded.
 * @returns At most one feature per kind, ordered by the fixed kind order.
 */
function unsupportedFeaturesOf(
  entryNames: readonly string[],
  files: Unzipped,
): readonly UnsupportedFeature[] {
  const details = new Map<UnsupportedFeatureKind, string>()
  const record = (kind: UnsupportedFeatureKind, detail: string): void => {
    if (!details.has(kind)) details.set(kind, detail)
  }

  for (const name of entryNames) {
    const path = name.toLowerCase()
    if (MACRO_PATHS.has(path)) record('macro', name)
    if (isCommentPart(path)) record('comment', name)
    if (hasPrefix(path, CHART_PREFIXES)) record('chart', name)
    if (hasPrefix(path, DRAWING_PREFIXES)) record('drawing', name)
    if (path.includes('vmldrawing')) record('shape', name)
    if (hasPrefix(path, EMBEDDING_PREFIXES)) record('embedded-object', name)
  }

  const storyParts: PackagePart[] = []
  const drawingParts: PackagePart[] = []
  for (const [name, content] of Object.entries(files)) {
    if (name === DOCUMENT_XML) continue
    const path = name.toLowerCase()
    const xml = strFromU8(content)
    if (isDrawingPart(path)) drawingParts.push({ name, xml })
    else if (isHeaderOrFooterPart(path)) storyParts.push({ name, xml })
  }

  const documentXml = files[DOCUMENT_XML]
  if (documentXml !== undefined) detectDocumentFeatures(strFromU8(documentXml), DOCUMENT_XML, record)
  for (const part of storyParts) detectDocumentFeatures(part.xml, part.name, record)
  for (const part of drawingParts) detectDrawingFeatures(part, record)

  const features: UnsupportedFeature[] = []
  for (const kind of UNSUPPORTED_KIND_ORDER) {
    const detail = details.get(kind)
    if (detail !== undefined) features.push({ kind, detail })
  }
  return features
}

/**
 * Classify the preserved construct classes and coupled markers of one scanned package.
 * @param entryNames - Entry inventory in central-directory order.
 * @param files - Entries the scan decoded.
 * @returns Preserved classes in the fixed order and coupled markers sorted by part then offset.
 */
function preservedConstructsOf(
  entryNames: readonly string[],
  files: Unzipped,
): PreservedConstructs {
  const found = new Set<PreservedConstructKind>()
  const coupled: CoupledConstructMarker[] = []

  for (const name of entryNames) classifyPreservedEntry(name.toLowerCase(), found)

  for (const [name, content] of Object.entries(files)) {
    const path = name.toLowerCase()
    if (path === 'xl/workbook.xml') {
      if (findTagOffsets(content, 'definedName').length > 0) found.add('defined-names')
      continue
    }
    if (path === 'xl/sharedstrings.xml') {
      for (const offset of richTextSharedStringOffsets(content)) {
        coupled.push({ kind: 'rich-text-shared-string', part: name, offset })
      }
      continue
    }
    if (isWordStoryPart(path)) {
      scanWordStory(name, content, found, coupled)
      continue
    }
    if (isWorksheetPart(path)) scanWorksheet(name, content, coupled)
  }

  const preserved = new Set<PreservedConstructKind>()
  for (const kind of PRESERVED_KIND_ORDER) {
    if (found.has(kind)) preserved.add(kind)
  }
  coupled.sort(compareMarkers)
  return { preserved, coupled }
}

/**
 * Record document-body constructs the document writer drops. The writer maps
 * paragraph text, run and paragraph formatting, heading styles, section page
 * setup, header and footer text, flow tables with merged cells, cell formatting
 * and nested tables, lists, and hyperlinks, and preserves images and note
 * references opaquely; tracked changes are the remaining document-body loss.
 * @param xml - Decoded `word/document.xml` or a header/footer part.
 * @param detail - Part name reported with each detected kind.
 * @param record - First-detail-per-kind recorder.
 */
function detectDocumentFeatures(
  xml: string,
  detail: string,
  record: (kind: UnsupportedFeatureKind, detail: string) => void,
): void {
  if (xml.includes('<w:ins ') || xml.includes('<w:del ')) record('tracked-changes', detail)
}

/**
 * Record workbook drawings whose anchors are not pictures. Cell pictures are
 * mapped to sheet drawings; shapes, charts, groups, and absolute anchors are
 * not, so a package containing them still refuses an edit-and-save.
 * @param part - Decoded `xl/drawings/*.xml` part.
 * @param record - First-detail-per-kind recorder.
 */
function detectDrawingFeatures(
  part: PackagePart,
  record: (kind: UnsupportedFeatureKind, detail: string) => void,
): void {
  if (UNMAPPED_DRAWING_ELEMENT.test(part.xml)) record('drawing', part.name)
}

/**
 * Enumerate every entry and decode only the ones a caller selects, enforcing
 * the entry-count and expanded-size bounds before any part is retained.
 * @param bytes - Complete package bytes.
 * @param select - Predicate over the stored name and its lowercased path.
 * @returns The entry inventory in central-directory order plus the decoded selection.
 * @throws OfficeExchangeError `too-large` when a bound is exceeded, `conversion-failed` when the archive cannot be read.
 */
function scanPackage(bytes: Uint8Array, select: EntrySelector): PackageScan {
  const entryNames: string[] = []
  let totalUncompressedBytes = 0

  let files: Unzipped
  try {
    files = unzipSync(bytes, {
      filter: (file) => {
        entryNames.push(file.name)
        if (entryNames.length > MAX_ENTRY_COUNT) {
          throw new OfficeExchangeError(
            'too-large',
            `Office package contains more than ${MAX_ENTRY_COUNT} entries`,
          )
        }
        totalUncompressedBytes += file.originalSize
        if (totalUncompressedBytes > MAX_UNCOMPRESSED_BYTES) {
          throw new OfficeExchangeError(
            'too-large',
            `Office package expands beyond ${MAX_UNCOMPRESSED_BYTES} bytes`,
          )
        }
        return select(file.name, file.name.toLowerCase())
      },
    })
  } catch (error) {
    if (error instanceof OfficeExchangeError) throw error
    const reason = error instanceof Error ? error.message : String(error)
    throw new OfficeExchangeError('conversion-failed', `Cannot read OOXML package: ${reason}`)
  }
  return { entryNames, files }
}

/**
 * Record the construct classes named by package entries. Part presence is the
 * whole evidence for these classes: the save carries the entry verbatim.
 * @param path - Lowercased entry name.
 * @param preserved - Accumulator of detected classes.
 */
function classifyPreservedEntry(path: string, preserved: Set<PreservedConstructKind>): void {
  if (MACRO_PATHS.has(path)) preserved.add('macro')
  if (hasPrefix(path, CHART_PART_PREFIXES)) preserved.add('chart')
  if (hasPrefix(path, ['xl/pivottables/', 'xl/pivotcache/'])) preserved.add('pivot-table')
  if (path.startsWith('xl/tables/')) preserved.add('excel-table')
  if (path === 'xl/calcchain.xml') preserved.add('calc-chain')
  if (isCommentEntry(path)) preserved.add('comment')
  if (hasPrefix(path, EMBEDDING_PREFIXES)) preserved.add('embedded-object')
  if (hasPrefix(path, DRAWING_PART_PREFIXES)) preserved.add('drawing')
  if (path.includes('vmldrawing') || path.endsWith('.vml')) preserved.add('shape')
  if (hasPrefix(path, THEME_PART_PREFIXES)) preserved.add('theme')
  const base = baseNameOf(path)
  if (base === 'fonttable.xml') preserved.add('font-table')
  if (base === 'websettings.xml') preserved.add('web-settings')
  if (path.startsWith('customxml/')) preserved.add('custom-xml')
  if (hasPrefix(path, EXTERNAL_LINK_PREFIXES)) preserved.add('external-link')
  if (path.startsWith('docprops/')) preserved.add('document-properties')
}

/**
 * Record preserved classes and coupled markers carried by one WordprocessingML
 * story part.
 * @param part - Part name reported with each marker.
 * @param bytes - Uncompressed part bytes.
 * @param preserved - Accumulator of detected classes.
 * @param coupled - Accumulator of coupled markers.
 */
function scanWordStory(
  part: string,
  bytes: Uint8Array,
  preserved: Set<PreservedConstructKind>,
  coupled: CoupledConstructMarker[],
): void {
  recordMarkers(coupled, 'tracked-change', part, tagOffsets(bytes, TRACKED_CHANGE_TAGS))
  recordMarkers(coupled, 'comment-range', part, tagOffsets(bytes, COMMENT_RANGE_TAGS))

  const controls = tagOffsets(bytes, ['w:sdt'])
  if (controls.length > 0) preserved.add('content-control')
  recordMarkers(coupled, 'content-control-wrap', part, controls)

  const fields = tagOffsets(bytes, FIELD_TAGS)
  if (fields.length > 0) preserved.add('field')
  recordMarkers(coupled, 'field-owned-text', part, fields)

  if (tagOffsets(bytes, BOOKMARK_TAGS).length > 0) preserved.add('bookmark')
  recordMarkers(coupled, 'ambiguous-alternate-content', part, tagOffsets(bytes, ['mc:AlternateContent']))
}

/**
 * Record shared and array formulas carried by one worksheet part. A cell edit
 * over either cannot be expressed as one cell-span splice.
 * @param part - Worksheet part name reported with each marker.
 * @param bytes - Uncompressed part bytes.
 * @param coupled - Accumulator of coupled markers.
 */
function scanWorksheet(
  part: string,
  bytes: Uint8Array,
  coupled: CoupledConstructMarker[],
): void {
  const { shared, array } = formulaOffsets(bytes)
  recordMarkers(coupled, 'shared-formula', part, shared)
  recordMarkers(coupled, 'array-formula', part, array)
}

/**
 * Append one marker per reported offset.
 * @param coupled - Accumulator of coupled markers.
 * @param kind - Construct class of the reported occurrences.
 * @param part - Part name carrying the occurrences.
 * @param offsets - Ascending byte offsets of the occurrences.
 */
function recordMarkers(
  coupled: CoupledConstructMarker[],
  kind: CoupledConstructMarker['kind'],
  part: string,
  offsets: readonly number[],
): void {
  for (const offset of offsets) coupled.push({ kind, part, offset })
}

/** Order markers by part name, then by byte offset, for deterministic output. */
function compareMarkers(left: CoupledConstructMarker, right: CoupledConstructMarker): number {
  if (left.part !== right.part) return left.part < right.part ? -1 : 1
  return left.offset - right.offset
}

/**
 * Locate `t="shared"` and `t="array"` formula elements in worksheet bytes.
 * The start tag ends at the first `>`; attribute values and formula text follow
 * it, so a formula containing `>` cannot extend the inspected tag.
 * @param bytes - Uncompressed worksheet part bytes.
 * @returns Ascending offsets grouped by formula kind.
 */
function formulaOffsets(bytes: Uint8Array): { shared: readonly number[]; array: readonly number[] } {
  const shared: number[] = []
  const array: number[] = []
  const needle = strToU8('<f')
  let from = 0
  while (from < bytes.length) {
    const start = indexOfBytes(bytes, needle, from)
    if (start < 0) break
    const boundary = bytes[start + needle.length]
    if (boundary === undefined || !isTagBoundary(boundary)) {
      from = start + needle.length
      continue
    }
    const end = indexOfByte(bytes, GREATER_THAN, start + needle.length)
    if (end < 0) break
    const match = FORMULA_TYPE.exec(strFromU8(bytes.subarray(start, end + 1), true))
    if (match?.[1] === 'shared') shared.push(start)
    else if (match?.[1] === 'array') array.push(start)
    from = end + 1
  }
  return { shared, array }
}

/**
 * Locate shared-string items that hold rich text runs (`<r>`). Changing such a
 * cell's text cannot rewrite the shared item without changing every referrer.
 * @param bytes - Uncompressed `xl/sharedStrings.xml` bytes.
 * @returns Ascending byte offsets of rich-text `<si>` elements.
 */
function richTextSharedStringOffsets(bytes: Uint8Array): readonly number[] {
  const itemStarts = findTagOffsets(bytes, 'si')
  const runStarts = findTagOffsets(bytes, 'r')
  const closeNeedle = strToU8('</si>')
  const offsets: number[] = []
  // Runs are ordered, so each item advances the cursor once instead of
  // rescanning every run from the first item.
  let runIndex = 0
  for (const start of itemStarts) {
    const openEnd = indexOfByte(bytes, GREATER_THAN, start)
    if (openEnd < 0) break
    if (bytes[openEnd - 1] === SLASH) continue
    const close = indexOfBytes(bytes, closeNeedle, openEnd + 1)
    if (close < 0) continue
    let run = runStarts[runIndex]
    while (run !== undefined && run < start) {
      runIndex += 1
      run = runStarts[runIndex]
    }
    if (run !== undefined && run < close) offsets.push(start)
  }
  return offsets
}

/**
 * Find every start tag with the exact element name, requiring a tag boundary
 * after the name so `<w:sdtPr>` is not reported as `<w:sdt>`.
 * @param bytes - Part bytes to scan.
 * @param tag - Qualified element name without angle brackets.
 * @returns Ascending byte offsets of matching start tags.
 */
function findTagOffsets(bytes: Uint8Array, tag: string): readonly number[] {
  const needle = strToU8(`<${tag}`)
  const offsets: number[] = []
  let from = 0
  while (from < bytes.length) {
    const index = indexOfBytes(bytes, needle, from)
    if (index < 0) break
    const boundary = bytes[index + needle.length]
    if (boundary !== undefined && isTagBoundary(boundary)) offsets.push(index)
    from = index + needle.length
  }
  return offsets
}

/**
 * Collect and order the offsets of several element names.
 * @param bytes - Part bytes to scan.
 * @param tags - Qualified element names without angle brackets.
 * @returns Ascending byte offsets across every named element.
 */
function tagOffsets(bytes: Uint8Array, tags: readonly string[]): readonly number[] {
  const offsets: number[] = []
  for (const tag of tags) offsets.push(...findTagOffsets(bytes, tag))
  return offsets.sort((left, right) => left - right)
}

/** Offset of the first occurrence of a byte sequence, or `-1`. */
function indexOfBytes(haystack: Uint8Array, needle: Uint8Array, from: number): number {
  const last = haystack.length - needle.length
  scan: for (let index = from; index <= last; index += 1) {
    for (let offset = 0; offset < needle.length; offset += 1) {
      if (haystack[index + offset] !== needle[offset]) continue scan
    }
    return index
  }
  return -1
}

/** Offset of the first occurrence of one byte, or `-1`. */
function indexOfByte(haystack: Uint8Array, byte: number, from: number): number {
  for (let index = from; index < haystack.length; index += 1) {
    if (haystack[index] === byte) return index
  }
  return -1
}

/** Whether a byte can follow an element name in a start tag. */
function isTagBoundary(byte: number): boolean {
  return byte === 0x20 || byte === 0x09 || byte === 0x0a || byte === 0x0d || byte === 0x2f || byte === 0x3e
}

function baseNameOf(path: string): string {
  const slash = path.lastIndexOf('/')
  return slash < 0 ? path : path.slice(slash + 1)
}

function isPreservedScanPart(path: string): boolean {
  return path === 'xl/workbook.xml'
    || path === 'xl/sharedstrings.xml'
    || isWordStoryPart(path)
    || isWorksheetPart(path)
}

function isWordStoryPart(path: string): boolean {
  return WORD_STORY_PATHS.has(path) || isHeaderOrFooterPart(path)
}

function isWorksheetPart(path: string): boolean {
  return path.startsWith('xl/worksheets/') && path.endsWith('.xml')
}

function isCommentEntry(path: string): boolean {
  return /^word\/comments[^/]*\.xml$/u.test(path)
    || /^xl\/comments[^/]*\.xml$/u.test(path)
    || path.startsWith('xl/threadedcomments/')
    || path.startsWith('ppt/comments/')
}

function isDrawingPart(path: string): boolean {
  return path.startsWith(WORKSHEET_DRAWING_PREFIX) && path.endsWith('.xml')
}

function hasPrefix(path: string, prefixes: readonly string[]): boolean {
  return prefixes.some(prefix => path.startsWith(prefix))
}

function isCommentPart(path: string): boolean {
  return path === 'word/comments.xml'
    || (path.startsWith('xl/comments') && path.endsWith('.xml'))
    || path.startsWith('ppt/comments/')
}

function isHeaderOrFooterPart(path: string): boolean {
  return (path.startsWith('word/header') || path.startsWith('word/footer')) && path.endsWith('.xml')
}
