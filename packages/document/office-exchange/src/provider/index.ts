/**
 * Local Office Exchange provider. Reads and writes Office files on disk through
 * the pure OOXML provider modules, reports unsupported features, and converts
 * legacy binary formats through the LibreOffice kit.
 * @module @forwork/office-exchange/provider
 */

import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Context } from '@forwork/cordis'
import z from '@forwork/schemastery'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG, resolveConfig, type OfficeExchangeConfig } from '../config.ts'
import { detectUnsupportedFeatures } from '../detect.ts'
import { OfficeExchangeError } from '../errors.ts'
import { formatOf, isEditableFormat, isLegacyFormat } from '../format.ts'
import { OfficeDocumentRemote } from '../remote.ts'
import { OfficeCommandBroker } from '../commands.ts'
import { OfficeExchange } from '../service.ts'
import type { ExchangeReadResult, UnitSnapshot } from '../snapshots.ts'
import type { OfficeFormat, UnsupportedFeature } from '../types.ts'
import { readDocument, writeDocument } from './docx.ts'
import { convertLegacyFile } from './legacy.ts'
import { readWorkbook, writeWorkbook } from './xlsx.ts'

/** Deployment bounds for local Office reads, writes, and legacy conversion. */
export interface Config {
  /** Largest accepted source or output file, in bytes; omitted uses the default. */
  maxFileBytes?: number
  /** Largest accepted spreadsheet cell or document table-cell count; omitted uses the default. */
  maxCells?: number
  /** Longest allowed single legacy conversion, in milliseconds; omitted uses the default. */
  maxConversionMs?: number
  /** Longest wait for a live editor to answer one document command, in milliseconds; omitted uses the default. */
  commandTimeoutMs?: number
}

/** Validated bounds applied before parsing, serializing, or converting a file. */
export const Config: z<Partial<Config>, Config> = z.object({
  maxFileBytes: z.natural().min(1).max(Number.MAX_SAFE_INTEGER)
    .default(DEFAULT_OFFICE_EXCHANGE_CONFIG.maxFileBytes),
  maxCells: z.natural().min(1).max(Number.MAX_SAFE_INTEGER)
    .default(DEFAULT_OFFICE_EXCHANGE_CONFIG.maxCells),
  maxConversionMs: z.natural().min(1).max(2_147_483_647)
    .default(DEFAULT_OFFICE_EXCHANGE_CONFIG.maxConversionMs),
  commandTimeoutMs: z.natural().min(1).max(2_147_483_647)
    .default(DEFAULT_OFFICE_EXCHANGE_CONFIG.commandTimeoutMs),
})

/**
 * Filesystem-backed Exchange provider for the deployment's local workspace.
 * Reads and writes do not modify the source path; `convertLegacy` writes its
 * result into the caller's directory, or a fresh per-call temporary directory
 * when none is supplied.
 */
export class LocalOfficeExchange extends OfficeExchange {
  static Config = Config
  private readonly config: OfficeExchangeConfig

  /**
   * @param ctx - Owning context whose service key is `officeExchange`.
   * @param config - Deployment bounds; omitted fields keep their defaults.
   * @throws RangeError When a configured bound is not a positive integer.
   */
  constructor(ctx: Context, config: Config = {}) {
    super(ctx)
    this.config = resolveConfig(config)
    ctx.plugin(OfficeDocumentRemote, { maxBytes: this.config.maxFileBytes })
    ctx.plugin(OfficeCommandBroker, { commandTimeoutMs: this.config.commandTimeoutMs })
  }

  /**
   * Read an editable Office file into a unit snapshot.
   * @param path - `.docx` or `.xlsx` file path.
   * @param signal - Optional cancellation; an abort reason propagates unchanged.
   * @returns The detected format, parsed snapshot, and unsupported features.
   * @throws {OfficeExchangeError} `unsupported-format` for preview-only or legacy
   * formats, `conversion-failed` when the file cannot be read or parsed.
   */
  async read(path: string, signal?: AbortSignal): Promise<ExchangeReadResult> {
    const format = requireFormat(path)
    if (!isEditableFormat(format)) {
      const reason = isLegacyFormat(format)
        ? 'convert it to an editable format with convertLegacy first'
        : 'PowerPoint files are preview-only'
      throw new OfficeExchangeError('unsupported-format', `Cannot read ${path}: ${reason}`)
    }
    const bytes = await this.readBytes(path, signal)
    if (format === 'xlsx') {
      return {
        format,
        snapshot: { kind: 'workbook', data: await readWorkbook(bytes, this.config) },
        unsupported: detectUnsupportedFeatures(bytes, this.config),
      }
    }
    return {
      format,
      snapshot: { kind: 'document', data: await readDocument(bytes, this.config) },
      unsupported: detectUnsupportedFeatures(bytes, this.config),
    }
  }

  /**
   * Serialize a unit snapshot to Office file bytes without writing them.
   * @param snapshot - Workbook or document content bound to the target format.
   * @param format - Target `.xlsx` or `.docx` format.
   * @param signal - Optional cancellation checked before and after serialization.
   * @returns Complete encoded file bytes.
   * @throws {OfficeExchangeError} `unsupported-format` for PowerPoint or legacy
   * targets, `conversion-failed` when the snapshot kind does not match the format
   * or serialization fails.
   */
  async write(snapshot: UnitSnapshot, format: OfficeFormat, signal?: AbortSignal): Promise<Uint8Array> {
    signal?.throwIfAborted()
    switch (format) {
      case 'xlsx': {
        if (snapshot.kind !== 'workbook') {
          throw new OfficeExchangeError('conversion-failed', `Cannot write a ${snapshot.kind} snapshot as .xlsx`)
        }
        const bytes = await writeWorkbook(snapshot.data, this.config)
        signal?.throwIfAborted()
        return bytes
      }
      case 'docx': {
        if (snapshot.kind !== 'document') {
          throw new OfficeExchangeError('conversion-failed', `Cannot write a ${snapshot.kind} snapshot as .docx`)
        }
        const bytes = await writeDocument(snapshot.data, this.config)
        signal?.throwIfAborted()
        return bytes
      }
      case 'pptx':
      case 'doc':
      case 'xls':
      case 'ppt':
        throw new OfficeExchangeError(
          'unsupported-format',
          `Cannot write .${format} files; only .docx and .xlsx are editable`,
        )
    }
  }

  /**
   * Report features an edit-and-save would drop, without parsing editable content.
   * @param path - `.docx`, `.xlsx`, or `.pptx` file path.
   * @param signal - Optional cancellation; an abort reason propagates unchanged.
   * @returns Detected unsupported features, empty when the file round-trips cleanly.
   * @throws {OfficeExchangeError} `unsupported-format` for a legacy file,
   * `conversion-failed` when the file cannot be read or is not an OOXML package.
   */
  async detectUnsupported(path: string, signal?: AbortSignal): Promise<readonly UnsupportedFeature[]> {
    const format = requireFormat(path)
    if (isLegacyFormat(format)) {
      throw new OfficeExchangeError(
        'unsupported-format',
        `Cannot inspect ${path}: convert it to an editable format with convertLegacy first`,
      )
    }
    const bytes = await this.readBytes(path, signal)
    return detectUnsupportedFeatures(bytes, this.config)
  }

  /**
   * Convert a legacy binary Office file into an editable format.
   * @param path - `.doc`, `.xls`, or `.ppt` file path.
   * @param signal - Optional cancellation passed to the conversion engine.
   * @param outputDir - Directory receiving the converted file; omitted uses a
   * fresh temporary directory outside any workspace.
   * @returns Absolute path of the converted file inside `outputDir` or the fresh
   * temporary directory.
   * @throws {OfficeExchangeError} `unsupported-format` for another extension,
   * `too-large` when the conversion exceeds `maxConversionMs`, and
   * `conversion-failed` when the engine is unavailable or conversion fails.
   */
  async convertLegacy(path: string, signal?: AbortSignal, outputDir?: string): Promise<string> {
    signal?.throwIfAborted()
    const request = { config: this.config, signal }
    if (outputDir !== undefined) return await convertLegacyFile(path, outputDir, request)
    const directory = await mkdtemp(join(tmpdir(), 'dsh-office-exchange-'))
    try {
      return await convertLegacyFile(path, directory, request)
    } catch (cause) {
      await Promise.allSettled([rm(directory, { recursive: true, force: true })])
      throw cause
    }
  }

  private async readBytes(path: string, signal?: AbortSignal): Promise<Uint8Array> {
    try {
      return await readFile(path, { signal })
    } catch (cause) {
      if (signal?.aborted === true) signal.throwIfAborted()
      /* v8 ignore next -- Node fs rejects with an Error; a non-Error rejection needs a mocked fs seam. */
      const detail = cause instanceof Error ? cause.message : String(cause)
      throw new OfficeExchangeError('conversion-failed', `Cannot read ${path}: ${detail}`)
    }
  }
}

function requireFormat(path: string): OfficeFormat {
  const format = formatOf(path)
  if (format === undefined) {
    throw new OfficeExchangeError('unsupported-format', `Unsupported Office file extension: ${path}`)
  }
  return format
}

export default LocalOfficeExchange
