/** OOXML package inspection for features the editable model cannot preserve. */
import { strFromU8, unzipSync } from 'fflate'
import type { Unzipped } from 'fflate'
import type { OfficeExchangeConfig } from './config.ts'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from './config.ts'
import { OfficeExchangeError } from './errors.ts'
import type { UnsupportedFeature, UnsupportedFeatureKind } from './types.ts'

const DOCUMENT_XML = 'word/document.xml'
const DOCUMENT_RELS = 'word/_rels/document.xml.rels'
const NUMBERING_XML = 'word/numbering.xml'

/** Parts that hold footnote and endnote definitions. */
const NOTE_PARTS: ReadonlySet<string> = new Set(['word/footnotes.xml', 'word/endnotes.xml'])

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
  'embedded-object',
  'image',
  'external-image',
  'hyperlink',
  'numbering',
  'borders',
  'footnotes',
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
  /** Decoded `word/_rels/document.xml.rels`, when the package declares relationships. */
  readonly documentRelsXml: Uint8Array | undefined
  /** Decoded `word/numbering.xml`, when the package declares numbering. */
  readonly numberingXml: Uint8Array | undefined
  /** Decoded header and footer parts, in central-directory order. */
  readonly storyParts: readonly PackagePart[]
  /** Decoded header and footer relationship parts, in central-directory order. */
  readonly storyRelsParts: readonly PackagePart[]
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

  const { entryNames, documentXml, documentRelsXml, numberingXml, storyParts, storyRelsParts, worksheetParts }
    = readPackage(bytes)
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
    if (hasPrefix(path, MEDIA_PREFIXES)) record('image', name)
    if (NOTE_PARTS.has(path)) record('footnotes', name)
  }

  if (documentRelsXml !== undefined && hasExternalImageRelationship(strFromU8(documentRelsXml))) {
    record('external-image', DOCUMENT_RELS)
  }
  if (documentXml !== undefined) detectDocumentFeatures(strFromU8(documentXml), DOCUMENT_XML, record)
  for (const part of storyParts) detectDocumentFeatures(part.xml, part.name, record)
  for (const part of storyRelsParts) {
    if (hasExternalImageRelationship(part.xml)) record('external-image', part.name)
  }
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
 * writes back paragraph text, run and paragraph formatting, heading styles,
 * section page setup, header and footer text, and flow tables.
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
  if (/<w:hyperlink(?=[\s/>])/u.test(xml)) record('hyperlink', detail)
  if (/<w:numPr(?=[\s/>])/u.test(xml)) record('numbering', detail)
  if (hasCellBorders(xml)) record('borders', detail)
  if (hasNoteReference(xml)) record('footnotes', detail)
  if (/<a:blip(?=[\s/>])[^>]*\br:link=/u.test(xml)) record('external-image', detail)
  if (hasMergedCellSpan(xml)) record('merged-cells', detail)
  if (hasNestedTable(xml)) record('nested-table', detail)
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
 * Report whether a table cell declares its own borders. Table borders
 * (`w:tblBorders`), paragraph borders (`w:pBdr`), and run borders (`w:bdr`)
 * are written by the provider and are not a loss.
 * @param xml - Decoded `word/document.xml`.
 * @returns `true` when a cell border element is present.
 */
function hasCellBorders(xml: string): boolean {
  return /<w:tcBorders(?=[\s/>])/u.test(xml)
}

/**
 * Report whether the body references a footnote or endnote.
 * @param xml - Decoded `word/document.xml`.
 * @returns `true` when a footnote or endnote reference is present.
 */
function hasNoteReference(xml: string): boolean {
  return /<w:(?:footnote|endnote)Reference(?=[\s/>])/u.test(xml)
}

/**
 * Report whether the document's relationships carry an image outside the
 * package. Hyperlink relationships also use an external target, so the
 * relationship type must name an image.
 * @param xml - Decoded `word/_rels/document.xml.rels` or a header/footer relationship part.
 * @returns `true` when an image relationship has `TargetMode="External"`.
 */
function hasExternalImageRelationship(xml: string): boolean {
  for (const match of xml.matchAll(/<Relationship(?=[\s/>])[^>]*>/gu)) {
    if (!/\bTargetMode="External"/u.test(match[0])) continue
    if (/\bType="[^"]*\/image"/u.test(match[0])) return true
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
        return file.name === DOCUMENT_XML
          || file.name === DOCUMENT_RELS
          || file.name === NUMBERING_XML
          || isWorksheetPart(path)
          || isHeaderOrFooterPart(path)
          || isStoryRelationshipPart(path)
      },
    })
  } catch (error) {
    if (error instanceof OfficeExchangeError) throw error
    const reason = error instanceof Error ? error.message : String(error)
    throw new OfficeExchangeError('conversion-failed', `Cannot read OOXML package: ${reason}`)
  }

  const worksheetParts: PackagePart[] = []
  const storyParts: PackagePart[] = []
  const storyRelsParts: PackagePart[] = []
  for (const [name, content] of Object.entries(entries)) {
    if (name === DOCUMENT_XML || name === DOCUMENT_RELS || name === NUMBERING_XML) continue
    const xml = strFromU8(content)
    if (isWorksheetPart(name.toLowerCase())) {
      worksheetParts.push({ name, xml })
      continue
    }
    if (isStoryRelationshipPart(name.toLowerCase())) storyRelsParts.push({ name, xml })
    else storyParts.push({ name, xml })
  }
  return {
    entryNames,
    documentXml: entries[DOCUMENT_XML],
    documentRelsXml: entries[DOCUMENT_RELS],
    numberingXml: entries[NUMBERING_XML],
    storyParts,
    storyRelsParts,
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

function isStoryRelationshipPart(path: string): boolean {
  return path.startsWith('word/_rels/') && path.endsWith('.rels')
}
