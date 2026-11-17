/** OOXML package inspection for features the editable model cannot preserve. */
import { strFromU8, unzipSync } from 'fflate'
import type { Unzipped } from 'fflate'
import type { OfficeExchangeConfig } from './config.ts'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from './config.ts'
import { OfficeExchangeError } from './errors.ts'
import type { UnsupportedFeature, UnsupportedFeatureKind } from './types.ts'

const DOCUMENT_XML = 'word/document.xml'
const NUMBERING_XML = 'word/numbering.xml'

const MACRO_PATHS: ReadonlySet<string> = new Set([
  'word/vbaproject.bin',
  'xl/vbaproject.bin',
  'ppt/vbaproject.bin',
])

/** Worksheet parts whose own elements the workbook writer does not rebuild. */
const WORKSHEET_PREFIX = 'xl/worksheets/'

/** Media parts an image relationship needs, and a rewrite drops. */
const MEDIA_PREFIXES: readonly string[] = ['word/media/', 'xl/media/']

const DRAWING_PREFIXES: readonly string[] = [
  'word/drawings/',
  'xl/drawings/',
  'ppt/drawings/',
]

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
  'header-footer',
  'embedded-object',
  'image',
  'hyperlink',
  'numbering',
  'run-formatting',
  'merged-cells',
  'nested-table',
  'conditional-formatting',
  'data-validation',
  'freeze-panes',
  'tab-color',
]

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

interface PackageContents {
  /** Entry names in central-directory order. */
  readonly entryNames: readonly string[]
  /** Decoded `word/document.xml`, when the package contains it. */
  readonly documentXml: Uint8Array | undefined
  /** Decoded `word/numbering.xml`, when the package declares numbering. */
  readonly numberingXml: Uint8Array | undefined
  /** Decoded worksheet parts, in central-directory order. */
  readonly worksheetParts: readonly PackagePart[]
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
  if (bytes.byteLength > config.maxFileBytes) {
    throw new OfficeExchangeError(
      'too-large',
      `Office file is ${bytes.byteLength} bytes, exceeding maxFileBytes ${config.maxFileBytes}`,
    )
  }

  const { entryNames, documentXml, numberingXml, worksheetParts } = readPackage(bytes)
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
    if (isHeaderOrFooterPart(path)) record('header-footer', name)
    if (hasPrefix(path, EMBEDDING_PREFIXES)) record('embedded-object', name)
    if (hasPrefix(path, MEDIA_PREFIXES)) record('image', name)
  }

  if (documentXml !== undefined) detectDocumentFeatures(strFromU8(documentXml), record)
  if (numberingXml !== undefined && hasNumberingDefinitions(strFromU8(numberingXml))) {
    record('numbering', NUMBERING_XML)
  }
  for (const part of worksheetParts) detectWorksheetFeatures(part, record)

  const features: UnsupportedFeature[] = []
  for (const kind of UNSUPPORTED_KIND_ORDER) {
    const detail = details.get(kind)
    if (detail !== undefined) features.push({ kind, detail })
  }
  return features
}

/**
 * Record document-body constructs that the document writer drops, because it
 * writes back only paragraph text, bold/italic/underline runs, heading styles,
 * and flow tables.
 * @param xml - Decoded `word/document.xml`.
 * @param record - First-detail-per-kind recorder.
 */
function detectDocumentFeatures(
  xml: string,
  record: (kind: UnsupportedFeatureKind, detail: string) => void,
): void {
  if (xml.includes('<w:ins ') || xml.includes('<w:del ')) record('tracked-changes', DOCUMENT_XML)
  if (/<w:hyperlink(?=[\s/>])/u.test(xml)) record('hyperlink', DOCUMENT_XML)
  if (/<w:numPr(?=[\s/>])/u.test(xml)) record('numbering', DOCUMENT_XML)
  if (hasUnsupportedTextFormatting(xml)) record('run-formatting', DOCUMENT_XML)
  if (hasMergedCellSpan(xml)) record('merged-cells', DOCUMENT_XML)
  if (hasNestedTable(xml)) record('nested-table', DOCUMENT_XML)
}

/**
 * Record worksheet elements that the workbook writer does not rebuild.
 * @param part - Decoded `xl/worksheets/*.xml` part.
 * @param record - First-detail-per-kind recorder.
 */
function detectWorksheetFeatures(
  part: PackagePart,
  record: (kind: UnsupportedFeatureKind, detail: string) => void,
): void {
  if (/<conditionalFormatting(?=[\s/>])/u.test(part.xml)) record('conditional-formatting', part.name)
  if (/<dataValidation(?=[\s/>])/u.test(part.xml)) record('data-validation', part.name)
  if (/<pane\b[^>]*\bstate="frozen(?:Split)?"/u.test(part.xml)) record('freeze-panes', part.name)
  if (/<tabColor(?=[\s/>])/u.test(part.xml)) record('tab-color', part.name)
  if (/<hyperlink(?=[\s/>])/u.test(part.xml)) record('hyperlink', part.name)
}

/**
 * Report whether a numbering part declares any numbering definition.
 * A part with no `w:abstractNum` or `w:num` element defines nothing a save
 * could lose.
 * @param xml - Decoded `word/numbering.xml`.
 * @returns `true` when the part declares a numbering definition.
 */
function hasNumberingDefinitions(xml: string): boolean {
  return /<w:(?:abstractNum|num)(?=[\s/>])/u.test(xml)
}

/**
 * Report whether a paragraph or run carries text formatting the writer drops.
 * Bold, italic, and underline round-trip, so a run holding only those is not a
 * loss.
 * @param xml - Decoded `word/document.xml`.
 * @returns `true` when a run property names a font, size, or color, or a paragraph property sets alignment.
 */
function hasUnsupportedTextFormatting(xml: string): boolean {
  for (const block of xml.matchAll(/<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>/gu)) {
    if (/<w:(?:rFonts|sz|szCs|color)(?=[\s/>])/u.test(block[0])) return true
  }
  for (const block of xml.matchAll(/<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>/gu)) {
    if (/<w:jc(?=[\s/>])/u.test(block[0])) return true
  }
  return false
}

/**
 * Report whether a table cell declares a horizontal or vertical span.
 * A `gridSpan` of one declares no span.
 * @param xml - Decoded `word/document.xml`.
 * @returns `true` when a cell has a multi-column span or a vertical merge.
 */
function hasMergedCellSpan(xml: string): boolean {
  if (/<w:vMerge(?=[\s/>])/u.test(xml)) return true
  for (const match of xml.matchAll(/<w:gridSpan\b[^>]*>/gu)) {
    const value = /\bw:val="([^"]*)"/u.exec(match[0])?.[1]
    if (value !== undefined && Number(value) > 1) return true
  }
  return false
}

/**
 * Report whether a table structurally contains another table's start.
 * `word/document.xml` also lists the root table and sibling tables, so the scan
 * counts open `w:tc` elements instead of matching a `w:tbl` tag anywhere.
 * @param xml - Decoded `word/document.xml`.
 * @returns `true` when a `w:tbl` starts inside a table cell.
 */
function hasNestedTable(xml: string): boolean {
  let cellDepth = 0
  for (const match of xml.matchAll(/<(\/?)w:(tc|tbl)(?=[\s/>])[^>]*>/gu)) {
    const closing = match[1] === '/'
    const tag = match[2]
    const selfClosing = match[0].endsWith('/>')
    if (tag === 'tc') {
      if (closing) cellDepth = Math.max(0, cellDepth - 1)
      else if (!selfClosing) cellDepth += 1
      continue
    }
    if (!closing && cellDepth > 0) return true
  }
  return false
}

function readPackage(bytes: Uint8Array): PackageContents {
  const entryNames: string[] = []
  let totalUncompressedBytes = 0

  let entries: Unzipped
  try {
    entries = unzipSync(bytes, {
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
        const path = file.name.toLowerCase()
        return file.name === DOCUMENT_XML || file.name === NUMBERING_XML || isWorksheetPart(path)
      },
    })
  } catch (error) {
    if (error instanceof OfficeExchangeError) throw error
    const reason = error instanceof Error ? error.message : String(error)
    throw new OfficeExchangeError('conversion-failed', `Cannot read OOXML package: ${reason}`)
  }

  const worksheetParts: PackagePart[] = []
  for (const [name, content] of Object.entries(entries)) {
    if (name === DOCUMENT_XML || name === NUMBERING_XML) continue
    worksheetParts.push({ name, xml: strFromU8(content) })
  }
  return {
    entryNames,
    documentXml: entries[DOCUMENT_XML],
    numberingXml: entries[NUMBERING_XML],
    worksheetParts,
  }
}

function isWorksheetPart(path: string): boolean {
  return path.startsWith(WORKSHEET_PREFIX) && path.endsWith('.xml')
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
