/**
 * Native editor snapshots produced and consumed by the Exchange seam. The
 * Univer interfaces live here, apart from the JSON-safe wire vocabulary in
 * `types.ts`, so a Remote Client imports only what it can carry.
 * @module @forwork/office-exchange/snapshots
 */

import type { IDocumentData, IWorkbookData } from '@univerjs/core'
import type { OfficeFormat, UnsupportedFeature } from './types.ts'

/**
 * A serializable unit snapshot. The discriminant selects the editor core's
 * native snapshot type (Option B) instead of a project-owned neutral model;
 * `data` is the Univer workbook or document data object, validated at the file
 * and model boundaries. The session log persists change operations, not these
 * snapshots.
 */
export type UnitSnapshot =
  | {
    /** A spreadsheet unit. */
    readonly kind: 'workbook'
    /** Univer workbook snapshot. */
    readonly data: IWorkbookData
  }
  | {
    /** A word-processing unit. */
    readonly kind: 'document'
    /** Univer document snapshot. */
    readonly data: IDocumentData
  }

/** Result of reading an Office file into an editable unit snapshot. */
export interface ExchangeReadResult {
  /** Detected source format. */
  readonly format: OfficeFormat
  /** Parsed editable content. */
  readonly snapshot: UnitSnapshot
  /**
   * Features that will not survive an edit-and-save. A non-empty list makes the
   * write path refuse by default rather than silently dropping content.
   */
  readonly unsupported: readonly UnsupportedFeature[]
}
