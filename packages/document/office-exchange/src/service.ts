/**
 * Office Exchange Service Definition (`ctx.officeExchange`). Providers map
 * Office files to and from unit snapshots and report features the editable
 * model cannot preserve; the editor and agent tools consume the seam without
 * importing a concrete conversion library.
 * @module @forwork/office-exchange/service
 */

import { Context, Service } from '@forwork/cordis'
import type { SourcePackage } from './package/types.ts'
import type { ExchangeReadResult, OfficeWriteResult, UnitSnapshot } from './snapshots.ts'
import type { OfficeFormat, UnsupportedFeature } from './types.ts'

declare module '@forwork/cordis' {
  interface Context {
    /** Office file to unit-snapshot conversion and legacy-format handling. */
    officeExchange: OfficeExchange
  }
}

/**
 * Service Definition for Office file interchange. A provider implements these
 * operations for the deployment's conversion libraries; consumers depend on the
 * seam only.
 */
export abstract class OfficeExchange extends Service {
  /** @param ctx - Owning context whose service key is `officeExchange`. */
  constructor(ctx: Context) {
    super(ctx, 'officeExchange')
  }

  /**
   * Read an Office file into an editable unit snapshot.
   * @param path - Absolute or workspace-relative file path.
   * @param signal - Optional cancellation.
   * @returns The detected format, snapshot, and unsupported features.
   */
  abstract read(path: string, signal?: AbortSignal): Promise<ExchangeReadResult>

  /**
   * Serialize a unit snapshot back to Office file bytes.
   *
   * `source` carries the bytes the caller read at its expected version so an
   * in-place writer can splice them. A snapshot without the source sidecar, or
   * a caller that supplies no source, uses the regeneration writer.
   * @param snapshot - Editable content to serialize.
   * @param format - Target Office format.
   * @param source - Source package bytes read at the caller's expected version;
   * omitted for a snapshot that has no source file.
   * @param signal - Optional cancellation.
   * @returns The encoded file bytes and what the save carried or changed.
   */
  abstract write(
    snapshot: UnitSnapshot,
    format: OfficeFormat,
    source?: SourcePackage,
    signal?: AbortSignal,
  ): Promise<OfficeWriteResult>

  /**
   * Report features a whole-file regeneration would drop.
   * @param path - File path to inspect.
   * @param signal - Optional cancellation.
   * @returns Detected unsupported features; empty when regeneration loses no mapped content.
   */
  abstract detectUnsupported(path: string, signal?: AbortSignal): Promise<readonly UnsupportedFeature[]>

  /**
   * Convert a legacy binary Office file to an editable format.
   * @param path - Legacy file path.
   * @param signal - Optional cancellation.
   * @param outputDir - Directory receiving the converted file; omitted uses a
   * fresh temporary directory outside any workspace.
   * @returns Path of the converted editable file.
   */
  abstract convertLegacy(path: string, signal?: AbortSignal, outputDir?: string): Promise<string>
}
