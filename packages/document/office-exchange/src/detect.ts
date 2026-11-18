/** OOXML package inspection for features the editable model cannot preserve. */
import { strFromU8, unzipSync } from 'fflate'
import type { Unzipped } from 'fflate'
import type { OfficeExchangeConfig } from './config.ts'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from './config.ts'
import { OfficeExchangeError } from './errors.ts'
import type { UnsupportedFeature, UnsupportedFeatureKind } from './types.ts'

const DOCUMENT_XML = 'word/document.xml'

const MACRO_PATHS: ReadonlySet<string> = new Set([
  'word/vbaproject.bin',
  'xl/vbaproject.bin',
  'ppt/vbaproject.bin',
])

/** Worksheet parts whose own elements the workbook writer does not rebuild. */
const WORKSHEET_PREFIX = 'xl/worksheets/'

/** Worksheet media parts the workbook writer does not rebuild. */
const WORKSHEET_MEDIA_PREFIX = 'xl/media/'

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
  'hyperlink',
  'conditional-formatting',
  'data-validation',
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
  /** Decoded header and footer parts, in central-directory order. */
  readonly storyParts: readonly PackagePart[]
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

  const { entryNames, documentXml, storyParts, worksheetParts } = readPackage(bytes)
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
    if (path.startsWith(WORKSHEET_MEDIA_PREFIX)) record('image', name)
  }

  if (documentXml !== undefined) detectDocumentFeatures(strFromU8(documentXml), DOCUMENT_XML, record)
  for (const part of storyParts) detectDocumentFeatures(part.xml, part.name, record)
  for (const part of worksheetParts) detectWorksheetFeatures(part, record)

  const features: UnsupportedFeature[] = []
  for (const kind of UNSUPPORTED_KIND_ORDER) {
    const detail = details.get(kind)
    if (detail !== undefined) features.push({ kind, detail })
  }
  return features
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
  if (/<hyperlink(?=[\s/>])/u.test(part.xml)) record('hyperlink', part.name)
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
          || isWorksheetPart(path)
          || isHeaderOrFooterPart(path)
      },
    })
  } catch (error) {
    if (error instanceof OfficeExchangeError) throw error
    const reason = error instanceof Error ? error.message : String(error)
    throw new OfficeExchangeError('conversion-failed', `Cannot read OOXML package: ${reason}`)
  }

  const worksheetParts: PackagePart[] = []
  const storyParts: PackagePart[] = []
  for (const [name, content] of Object.entries(entries)) {
    if (name === DOCUMENT_XML) continue
    const xml = strFromU8(content)
    if (isWorksheetPart(name.toLowerCase())) worksheetParts.push({ name, xml })
    else storyParts.push({ name, xml })
  }
  return {
    entryNames,
    documentXml: entries[DOCUMENT_XML],
    storyParts,
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
