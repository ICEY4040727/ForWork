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
]

/** Package entries a plausible document can contain before the archive is treated as hostile. */
const MAX_ENTRY_COUNT = 10_000

/** Bytes a plausible package can expand to, guarding against decompression bombs. */
const MAX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024

interface PackageContents {
  /** Entry names in central-directory order. */
  readonly entryNames: readonly string[]
  /** Decoded `word/document.xml`, when the package contains it. */
  readonly documentXml: Uint8Array | undefined
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

  const { entryNames, documentXml } = readPackage(bytes)
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
  }

  if (documentXml !== undefined) {
    const xml = strFromU8(documentXml)
    if (xml.includes('<w:ins ') || xml.includes('<w:del ')) record('tracked-changes', DOCUMENT_XML)
  }

  const features: UnsupportedFeature[] = []
  for (const kind of UNSUPPORTED_KIND_ORDER) {
    const detail = details.get(kind)
    if (detail !== undefined) features.push({ kind, detail })
  }
  return features
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
        return file.name === DOCUMENT_XML
      },
    })
  } catch (error) {
    if (error instanceof OfficeExchangeError) throw error
    const reason = error instanceof Error ? error.message : String(error)
    throw new OfficeExchangeError('conversion-failed', `Cannot read OOXML package: ${reason}`)
  }

  return { entryNames, documentXml: entries[DOCUMENT_XML] }
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
