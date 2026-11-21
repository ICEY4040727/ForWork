/**
 * Local Office Exchange provider. Reads and writes Office files on disk through
 * the pure OOXML provider modules, reports unsupported features, and converts
 * legacy binary formats through the LibreOffice kit.
 * @module @forwork/office-exchange/provider
 */

import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { IDocumentData } from '@univerjs/core'
import type { Context } from '@forwork/cordis'
import z from '@forwork/schemastery'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG, resolveConfig, type OfficeExchangeConfig } from '../config.ts'
import { detectOfficePackage, detectUnsupportedFeatures } from '../detect.ts'
import { OfficeExchangeError } from '../errors.ts'
import { formatOf, isEditableFormat, isLegacyFormat } from '../format.ts'
import type { SourcePackage } from '../package/types.ts'
import { OfficeDocumentRemote } from '../remote.ts'
import { OfficeCommandBroker } from '../commands.ts'
import { OfficeExchange } from '../service.ts'
import type { ExchangeReadResult, OfficeWriteResult, UnitSnapshot } from '../snapshots.ts'
import { SAVE_DISPOSITION_NAME_LIMIT } from '../source/save-disposition.ts'
import { parseSourceState, SourceStateError } from '../source/source-state.ts'
import type { OfficeFormat, SaveDisposition, SaveDispositionGroup, UnsupportedFeature } from '../types.ts'
import { readDocumentWithSource, writeDocument, writeDocumentFromSource } from './docx.ts'
import { convertLegacyFile } from './legacy.ts'
import { readWorkbookWithSource, writeWorkbook, writeWorkbookFromSource } from './xlsx.ts'

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
  /** Largest number of source unit anchors recorded for one opened file; omitted uses the default. */
  maxSourceAnchors?: number
  /** Largest byte length of the encoded source-anchor payload for one opened file; omitted uses the default. */
  maxSourceAnchorBytes?: number
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
  maxSourceAnchors: z.natural().min(1).max(Number.MAX_SAFE_INTEGER)
    .default(DEFAULT_OFFICE_EXCHANGE_CONFIG.maxSourceAnchors),
  maxSourceAnchorBytes: z.natural().min(1).max(Number.MAX_SAFE_INTEGER)
    .default(DEFAULT_OFFICE_EXCHANGE_CONFIG.maxSourceAnchorBytes),
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
   * @param path - `.docx`, `.docm`, `.xlsx`, or `.xlsm` file path.
   * @param signal - Optional cancellation; an abort reason propagates unchanged.
   * @returns The detected format, parsed snapshot, unsupported features, and
   * preserved construct classes.
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
    // One package scan classifies both reports the editor consumes.
    const detection = detectOfficePackage(bytes, this.config)
    const preserved = [...detection.preserved].sort(compareAscii)
    if (format === 'xlsx' || format === 'xlsm') {
      return {
        format,
        snapshot: { kind: 'workbook', data: await readWorkbookWithSource(bytes, format, this.config) },
        unsupported: detection.unsupported,
        preserved,
      }
    }
    return {
      format,
      snapshot: { kind: 'document', data: await readDocumentWithSource(bytes, format, this.config) },
      unsupported: detection.unsupported,
      preserved,
    }
  }

  /**
   * Serialize a unit snapshot to Office file bytes without writing them.
   *
   * A snapshot with `data.officeExchange.source` and a source package saves by
   * splicing only the edited units into that package; without either, the
   * legacy regeneration writer builds a fresh package and claims no span
   * identity. A snapshot with a sidecar but no source package keeps the legacy
   * path, so a caller that cannot supply the read-time bytes still saves.
   * Macro-enabled formats refuse that fallback: regeneration would emit the
   * macro-free content type and drop the VBA project.
   *
   * When the splice path runs, preserved base64 parts are never consulted: the
   * package read at the caller's expected version is authoritative.
   * @param snapshot - Workbook or document content bound to the target format.
   * @param format - Target `.docx`, `.docm`, `.xlsx`, or `.xlsm` format.
   * @param source - Source package read at the caller's expected version; when
   * provided, an attached source sidecar selects the splice writer.
   * @param signal - Optional cancellation checked before and after serialization.
   * @returns Complete encoded file bytes and the save disposition.
   * @throws {OfficeExchangeError} `unsupported-format` for PowerPoint or legacy
   * targets, `conversion-failed` when the snapshot kind does not match the format,
   * the source sidecar is malformed, or serialization fails, `too-large` when a
   * bound is exceeded, and `unsupported-feature` when a coupled construct
   * overlaps an edited unit or a macro-enabled target has no anchored source.
   */
  async write(
    snapshot: UnitSnapshot,
    format: OfficeFormat,
    source?: SourcePackage,
    signal?: AbortSignal,
  ): Promise<OfficeWriteResult> {
    signal?.throwIfAborted()
    const sourceState = sourceStateValueOf(snapshot.data)
    if (sourceState !== undefined) this.assertSourceState(sourceState, source)
    switch (format) {
      case 'xlsx':
      case 'xlsm': {
        if (snapshot.kind !== 'workbook') {
          throw new OfficeExchangeError('conversion-failed', `Cannot write a ${snapshot.kind} snapshot as .${format}`)
        }
        if (sourceState !== undefined && source !== undefined) {
          const written = await writeWorkbookFromSource(snapshot.data, source.bytes, this.config, signal)
          return { bytes: written.bytes, disposition: saveDispositionOf(written.disposition) }
        }
        if (format === 'xlsm') throw macroRegenerationRefusal(format)
        const bytes = await writeWorkbook(snapshot.data, this.config)
        signal?.throwIfAborted()
        return { bytes, disposition: regenerationDisposition() }
      }
      case 'docx':
      case 'docm': {
        if (snapshot.kind !== 'document') {
          throw new OfficeExchangeError('conversion-failed', `Cannot write a ${snapshot.kind} snapshot as .${format}`)
        }
        if (sourceState !== undefined && source !== undefined) {
          const written = await writeDocumentFromSource(snapshot.data, source.bytes, this.config, signal)
          return { bytes: written.bytes, disposition: saveDispositionOf(written.disposition) }
        }
        if (format === 'docm') throw macroRegenerationRefusal(format)
        const data = sourceState === undefined ? snapshot.data : withoutPreservedParts(snapshot.data)
        const bytes = await writeDocument(data, this.config)
        signal?.throwIfAborted()
        return { bytes, disposition: regenerationDisposition() }
      }
      case 'pptx':
      case 'doc':
      case 'xls':
      case 'ppt':
        throw new OfficeExchangeError(
          'unsupported-format',
          `Cannot write .${format} files; only .docx, .docm, .xlsx, and .xlsm are editable`,
        )
    }
  }

  /**
   * Require an attached source sidecar to satisfy the data-model rules.
   *
   * A sidecar whose recorded package digest does not describe the supplied
   * source is validated structurally only; its offsets will be rebuilt from
   * that source by the splice writer, so bounding them against the new part
   * lengths would reject a legitimate second save.
   * @param value - Untrusted `data.officeExchange.source` value.
   * @param source - Source package read at the caller's expected version; when
   * omitted, part names are bounded without checking them against a package.
   * @throws {OfficeExchangeError} `conversion-failed` when the sidecar is
   * malformed and `too-large` when it exceeds an anchor budget.
   */
  private assertSourceState(value: unknown, source: SourcePackage | undefined): void {
    const stale = source !== undefined && sourceStateDiffersFromPackage(value, source)
    const partByteLengths = source === undefined || stale
      ? inferredPartLengths(value)
      : sourcePartLengths(source)
    try {
      parseSourceState(value, {
        partByteLengths,
        maxSourceAnchors: this.config.maxSourceAnchors,
        maxSourceAnchorBytes: this.config.maxSourceAnchorBytes,
      })
    } catch (error) {
      /* v8 ignore start -- Validated bounds make any non-SourceStateError rejection unreachable here. */
      if (!(error instanceof SourceStateError)) throw error
      /* v8 ignore stop */
      const code = error.code === 'anchor-budget-exceeded' ? 'too-large' : 'conversion-failed'
      throw new OfficeExchangeError(code, `Snapshot source state is malformed: ${error.message}`)
    }
  }

  /**
   * Report features an edit-and-save would drop, without parsing editable content.
   * @param path - `.docx`, `.docm`, `.xlsx`, `.xlsm`, or `.pptx` file path.
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

/** A snapshot data object that may carry the Exchange source sidecar. */
interface OfficeExchangeCarrier {
  /** Provider-owned state stored inside the editable data object. */
  readonly officeExchange?: unknown
}

/** A document data object carrying the provider-owned sidecar fields. */
interface DocumentOfficeExchangeCarrier extends IDocumentData {
  /** Provider-owned state stored inside the editable data object. */
  officeExchange?: Record<string, unknown>
}

/**
 * Read the untrusted source sidecar from a snapshot.
 * @param data - Document or workbook data object.
 * @returns The sidecar value, or `undefined` when the snapshot is legacy.
 * @throws {OfficeExchangeError} `conversion-failed` when `officeExchange` is not an object.
 */
function sourceStateValueOf(data: UnitSnapshot['data']): unknown {
  const extension = (data as OfficeExchangeCarrier).officeExchange
  if (extension === undefined) return undefined
  if (!isRecord(extension)) {
    throw new OfficeExchangeError('conversion-failed', 'Snapshot officeExchange state is malformed')
  }
  return extension['source']
}

/**
 * Drop the legacy base64 preserved parts while keeping the source sidecar.
 * @param data - Document data that may carry both sidecars.
 * @returns The same object when it carries no preserved parts.
 */
function withoutPreservedParts(data: IDocumentData): IDocumentData {
  const extension = (data as DocumentOfficeExchangeCarrier).officeExchange
  if (extension === undefined || extension['docx'] === undefined) return data
  const rest: Record<string, unknown> = { ...extension }
  delete rest['docx']
  const result: DocumentOfficeExchangeCarrier = { ...data, officeExchange: rest }
  return result
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function sourcePartLengths(source: SourcePackage): ReadonlyMap<string, number> {
  const lengths = new Map<string, number>()
  for (const entry of source.entries) {
    if (!entry.name.endsWith('/')) lengths.set(entry.name, entry.uncompressedSize)
  }
  return lengths
}

/**
 * Whether a sidecar's recorded package identity no longer describes the source.
 *
 * A `true` answer means the splice writer must rebuild anchors from the source;
 * it is not a conflict, because the caller's version guard already proved the
 * file is the one it expects.
 * @param value - Untrusted sidecar value.
 * @param source - Source package read at the caller's expected version.
 * @returns `true` when the digest or byte length differs; malformed headers are
 * left for the sidecar parser to reject.
 */
function sourceStateDiffersFromPackage(value: unknown, source: SourcePackage): boolean {
  if (!isRecord(value)) return false
  return value['packageDigest'] !== source.digest || value['byteLength'] !== source.bytes.byteLength
}

/**
 * Bound each part a sidecar names without claiming a real package length.
 * Structural, identity, digest, overlap, and budget rules still apply, so a
 * malformed sidecar fails even when the caller supplied no source package.
 * @param value - Untrusted sidecar value.
 * @returns Every named part mapped to the largest safe offset.
 */
function inferredPartLengths(value: unknown): ReadonlyMap<string, number> {
  const names = new Set<string>()
  if (isRecord(value)) {
    for (const key of ['sourceParts', 'anchors', 'sheets']) {
      const entries = value[key]
      if (!Array.isArray(entries)) continue
      for (const entry of entries) {
        const name = typeof entry === 'string'
          ? entry
          : isRecord(entry) && typeof entry['part'] === 'string' ? entry['part'] : undefined
        if (name !== undefined) names.add(name)
      }
    }
  }
  return new Map([...names].map(name => [name, Number.MAX_SAFE_INTEGER]))
}

/**
 * Apply the save-disposition contract to one writer's plan-derived disposition.
 *
 * The writer computes the groups from its splice plan; this dispatch pass makes
 * the returned payload deterministic for both formats: every name list is
 * sorted ASCII and truncated to {@link SAVE_DISPOSITION_NAME_LIMIT}, while each
 * count stays exact.
 * @param disposition - disposition produced by a splice writer from its plan.
 * @returns the normalized disposition with the fixed name limit.
 */
export function saveDispositionOf(disposition: SaveDisposition): SaveDisposition {
  return {
    schemaVersion: 1,
    verbatimCount: disposition.verbatimCount,
    spliced: saveDispositionGroup(disposition.spliced),
    added: saveDispositionGroup(disposition.added),
    removed: saveDispositionGroup(disposition.removed),
    nameLimit: SAVE_DISPOSITION_NAME_LIMIT,
  }
}

/**
 * Sort and cap one disposition group without changing its exact count.
 * @param group - entries one writer reported for one status.
 * @returns the group's exact count and its sorted, capped name prefix.
 */
function saveDispositionGroup(group: SaveDispositionGroup): SaveDispositionGroup {
  return {
    count: group.count,
    names: [...group.names].sort(compareAscii).slice(0, SAVE_DISPOSITION_NAME_LIMIT),
  }
}

/** Compare two entry names by ASCII code unit order. */
function compareAscii(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0
}

/**
 * The disposition a regeneration writer reports. Every output entry is built
 * from the editable model, so the writer has no source-entry status to claim.
 * @returns An empty disposition that never claims verbatim or span identity.
 */
function regenerationDisposition(): SaveDisposition {
  return {
    schemaVersion: 1,
    verbatimCount: 0,
    spliced: { count: 0, names: [] },
    added: { count: 0, names: [] },
    removed: { count: 0, names: [] },
    nameLimit: SAVE_DISPOSITION_NAME_LIMIT,
  }
}

/**
 * Build the refusal for a macro-enabled save the regeneration writer would turn macro-free.
 * @param format - Macro-enabled target the caller requested without an anchored source.
 * @returns The `unsupported-feature` failure because the editable model cannot
 * produce the VBA project or the macro-enabled main content type.
 */
function macroRegenerationRefusal(format: 'docm' | 'xlsm'): OfficeExchangeError {
  return new OfficeExchangeError(
    'unsupported-feature',
    `Cannot save a .${format} snapshot without the source package it was read from: `
    + 'regenerating the file would drop its VBA project.',
  )
}

function requireFormat(path: string): OfficeFormat {
  const format = formatOf(path)
  if (format === undefined) {
    throw new OfficeExchangeError('unsupported-format', `Unsupported Office file extension: ${path}`)
  }
  return format
}

export default LocalOfficeExchange
