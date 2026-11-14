/**
 * Legacy binary Office conversion (`.doc`, `.xls`, `.ppt`) to the editable
 * formats. Conversion runs through the prebuilt LibreOffice engine in
 * `@deepseek-ai/libreoffice-kit`; the source file is opened read-only and the
 * result is written to a caller-owned output directory.
 */
import { mkdir, stat } from 'node:fs/promises'
import { basename, join, resolve } from 'node:path'
import { ConversionError, createConverter, type Converter } from '@deepseek-ai/libreoffice-kit'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from '../config.ts'
import { OfficeExchangeError } from '../errors.ts'

/** Editable format a legacy binary format converts to. */
type LegacyTargetFormat = 'docx' | 'xlsx' | 'pptx'

/**
 * Map a legacy Office path to its modern equivalent. Pure: the returned path
 * replaces only the final extension and keeps the original stem and directory.
 * @param path - Source path whose final extension decides the mapping.
 * @returns The converted path and format, or `undefined` for any other extension.
 */
export function legacyTarget(
  path: string,
): { targetPath: string; format: LegacyTargetFormat } | undefined {
  const dot = path.lastIndexOf('.')
  if (dot < 0) return undefined
  const suffix = path.slice(dot + 1).toLowerCase()
  switch (suffix) {
    case 'doc':
      return { targetPath: `${path.slice(0, dot)}.docx`, format: 'docx' }
    case 'xls':
      return { targetPath: `${path.slice(0, dot)}.xlsx`, format: 'xlsx' }
    case 'ppt':
      return { targetPath: `${path.slice(0, dot)}.pptx`, format: 'pptx' }
    default:
      return undefined
  }
}

/**
 * Convert one legacy binary Office file into `outputDir` without modifying the source.
 * @param inputPath - Existing `.doc`, `.xls`, or `.ppt` file.
 * @param outputDir - Directory receiving the converted file; created when absent.
 * @param signal - Optional cancellation; the converter stops its engine process.
 * @returns Absolute path of the converted, non-empty file.
 * @throws {OfficeExchangeError} `unsupported-format` for another extension and
 * `conversion-failed` when the engine is unavailable, the conversion fails, or
 * the output is missing or empty.
 */
export async function convertLegacyFile(
  inputPath: string,
  outputDir: string,
  signal?: AbortSignal,
): Promise<string> {
  signal?.throwIfAborted()
  const target = legacyTarget(inputPath)
  if (target === undefined) {
    throw new OfficeExchangeError(
      'unsupported-format',
      `Legacy conversion supports .doc, .xls, and .ppt, not ${inputPath}`,
    )
  }
  const source = resolve(inputPath)
  const directory = resolve(outputDir)
  const outputPath = join(directory, basename(target.targetPath))
  await mkdir(directory, { recursive: true })
  signal?.throwIfAborted()

  let converter: Converter | undefined
  try {
    converter = await createConverter({ timeoutMs: DEFAULT_OFFICE_EXCHANGE_CONFIG.maxConversionMs })
    signal?.throwIfAborted()
    await converter.convert({ inputPath: source, outputPath }, signal)
    signal?.throwIfAborted()
    const outputSize = await sizeOf(outputPath)
    if (outputSize === undefined || outputSize === 0) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `The converted file is missing or empty: ${outputPath}`,
      )
    }
    return outputPath
  } catch (cause) {
    signal?.throwIfAborted()
    if (cause instanceof OfficeExchangeError) throw cause
    if (cause instanceof ConversionError && cause.code === 'unavailable') {
      throw new OfficeExchangeError('conversion-failed', 'LibreOffice is not installed')
    }
    const detail = cause instanceof Error ? cause.message : String(cause)
    throw new OfficeExchangeError('conversion-failed', `LibreOffice conversion failed: ${detail}`)
  } finally {
    if (converter !== undefined) await Promise.allSettled([converter.dispose()])
  }
}

async function sizeOf(path: string): Promise<number | undefined> {
  try {
    const info = await stat(path)
    return info.isFile() ? info.size : undefined
  } catch (cause) {
    if (cause instanceof Error && 'code' in cause && cause.code === 'ENOENT') return undefined
    throw cause
  }
}
