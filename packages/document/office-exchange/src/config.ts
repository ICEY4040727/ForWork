/** Bounded conversion sizes and timeouts for the Exchange seam. */

/**
 * Limits applied before an Office file is parsed or converted. Every value is a
 * positive integer; a provider fails rather than truncating work that exceeds a
 * bound.
 */
export interface OfficeExchangeConfig {
  /** Largest accepted source file, in bytes. */
  maxFileBytes: number
  /** Largest accepted spreadsheet cell count. */
  maxCells: number
  /** Longest allowed single conversion, in milliseconds. */
  maxConversionMs: number
  /** Longest wait for a live editor to answer one document command, in milliseconds. */
  commandTimeoutMs: number
  /** Largest number of source unit anchors recorded for one opened file. */
  maxSourceAnchors: number
  /** Largest byte length of the encoded source-anchor payload for one opened file. */
  maxSourceAnchorBytes: number
}

/** Default bounds shared by providers that do not configure their own. */
export const DEFAULT_OFFICE_EXCHANGE_CONFIG: OfficeExchangeConfig = {
  maxFileBytes: 25 * 1024 * 1024,
  maxCells: 2_000_000,
  maxConversionMs: 60_000,
  commandTimeoutMs: 5_000,
  maxSourceAnchors: 100_000,
  maxSourceAnchorBytes: 8 * 1024 * 1024,
}

/**
 * Merge overrides into the defaults and validate every resulting bound.
 * @param overrides - Bounds replacing the defaults; omitted fields keep theirs.
 * @returns A complete, validated configuration.
 * @throws RangeError When a bound is not a positive finite integer.
 */
export function resolveConfig(overrides: Partial<OfficeExchangeConfig>): OfficeExchangeConfig {
  const config: OfficeExchangeConfig = {
    maxFileBytes: overrides.maxFileBytes ?? DEFAULT_OFFICE_EXCHANGE_CONFIG.maxFileBytes,
    maxCells: overrides.maxCells ?? DEFAULT_OFFICE_EXCHANGE_CONFIG.maxCells,
    maxConversionMs: overrides.maxConversionMs ?? DEFAULT_OFFICE_EXCHANGE_CONFIG.maxConversionMs,
    commandTimeoutMs: overrides.commandTimeoutMs ?? DEFAULT_OFFICE_EXCHANGE_CONFIG.commandTimeoutMs,
    maxSourceAnchors: overrides.maxSourceAnchors ?? DEFAULT_OFFICE_EXCHANGE_CONFIG.maxSourceAnchors,
    maxSourceAnchorBytes: overrides.maxSourceAnchorBytes ?? DEFAULT_OFFICE_EXCHANGE_CONFIG.maxSourceAnchorBytes,
  }
  assertPositiveInteger('maxFileBytes', config.maxFileBytes)
  assertPositiveInteger('maxCells', config.maxCells)
  assertPositiveInteger('maxConversionMs', config.maxConversionMs)
  assertPositiveInteger('commandTimeoutMs', config.commandTimeoutMs)
  assertPositiveInteger('maxSourceAnchors', config.maxSourceAnchors)
  assertPositiveInteger('maxSourceAnchorBytes', config.maxSourceAnchorBytes)
  return config
}

function assertPositiveInteger(field: keyof OfficeExchangeConfig, value: number): void {
  if (!Number.isInteger(value) || value <= 0) {
    throw new RangeError(`${field} must be a positive integer, got ${String(value)}`)
  }
}
