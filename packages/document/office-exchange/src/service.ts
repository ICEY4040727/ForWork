/**
 * Office Exchange Service Definition (`ctx.officeExchange`). Providers map
 * Office files to and from unit snapshots and report features the editable
 * model cannot preserve; the editor and agent tools consume the seam without
 * importing a concrete conversion library.
 * @module @forwork/office-exchange/service
 */

import { Context, Service } from '@forwork/cordis'
import type { ExchangeReadResult, UnitSnapshot } from './snapshots.ts'
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
   * @param snapshot - Editable content to serialize.
   * @param format - Target Office format.
   * @param signal - Optional cancellation.
   * @returns The encoded file bytes.
   */
  abstract write(snapshot: UnitSnapshot, format: OfficeFormat, signal?: AbortSignal): Promise<Uint8Array>

  /**
   * Report features that would be lost by an edit-and-save.
   * @param path - File path to inspect.
   * @param signal - Optional cancellation.
   * @returns Detected unsupported features; empty when the file round-trips cleanly.
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
