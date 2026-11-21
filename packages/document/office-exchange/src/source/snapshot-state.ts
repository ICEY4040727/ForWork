/**
 * Validation of a snapshot's source sidecar against the package being saved.
 *
 * Both in-place writers open the source package first, then parse the sidecar
 * with bounds taken from that package. This module owns that bind so a provider
 * only supplies the product name used in the failure report.
 * @module @forwork/office-exchange/source/snapshot-state
 */

import type { OfficeExchangeConfig } from '../config.ts'
import { OfficeExchangeError } from '../errors.ts'
import type { SourcePackage } from '../package/types.ts'
import { parseSourceState, SourceStateError } from './source-state.ts'
import type { OfficeSourceState } from './source-state.ts'

/**
 * Validate a snapshot sidecar against the source package it describes.
 * @param value - Untrusted `data.officeExchange.source` value.
 * @param source - Source package read for this save.
 * @param bounds - Validated anchor count and byte budgets.
 * @param subject - Product name used in the failure detail.
 * @returns The sidecar with every span bounded by its source part.
 * @throws {OfficeExchangeError} `conversion-failed` when the sidecar is malformed.
 */
export function parseSnapshotSourceState(
  value: unknown,
  source: SourcePackage,
  bounds: OfficeExchangeConfig,
  subject: 'Document' | 'Workbook',
): OfficeSourceState {
  const partByteLengths = new Map<string, number>()
  for (const entry of source.entries) {
    if (!entry.name.endsWith('/')) partByteLengths.set(entry.name, entry.uncompressedSize)
  }
  try {
    return parseSourceState(value, {
      partByteLengths,
      maxSourceAnchors: bounds.maxSourceAnchors,
      maxSourceAnchorBytes: bounds.maxSourceAnchorBytes,
    })
  } catch (error) {
    if (error instanceof SourceStateError) {
      throw new OfficeExchangeError(
        'conversion-failed',
        `${subject} snapshot source state is malformed: ${error.message}`,
      )
    }
    throw error
  }
}
