/** Typed failures for the Office Exchange seam. */

/** Stable failure codes consumers branch on instead of message text. */
export type OfficeExchangeErrorCode =
  | 'unsupported-format'
  | 'unsupported-feature'
  | 'too-large'
  | 'conversion-failed'
  | 'not-implemented'
  | 'out-of-scope'

/** A classified Exchange failure. */
export class OfficeExchangeError extends Error {
  /**
   * @param code - Stable classification for branching consumers.
   * @param message - Localized, actionable explanation.
   */
  constructor(
    readonly code: OfficeExchangeErrorCode,
    message: string,
  ) {
    super(message)
    this.name = 'OfficeExchangeError'
  }
}
