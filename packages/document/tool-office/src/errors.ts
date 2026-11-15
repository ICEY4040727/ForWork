/**
 * Structured model-facing failures for the Office document tools.
 *
 * A refusal is an ordinary result value, not a thrown error, so the model can
 * branch on the stable `code` and retry with a fresh read or save. Unexpected
 * failures still throw and reach the executor as ordinary tool errors.
 * @module @deepseek-ai/dsh-tool-office/errors
 */

import { OfficeExchangeError } from '@deepseek-ai/dsh-office-exchange'
import type { UnsupportedFeature } from '@deepseek-ai/dsh-office-exchange'
import { RemoteError } from '@deepseek-ai/dsh-typert-protocol'

/** Stable refusal codes the model branches on. */
export type OfficeToolErrorCode =
  | 'conflict'
  | 'unsupported-feature'
  | 'out-of-scope'
  | 'document-missing'
  | 'document-not-open'
  | 'pending-edits'
  | 'bound-exceeded'
  | 'timeout'
  | 'invalid-request'
  | 'conversion-failed'

/** One refused document operation. */
export interface OfficeToolFailure {
  /** Marks a refused operation. */
  readonly status: 'error'
  /** Stable classification for the model. */
  readonly code: OfficeToolErrorCode
  /** Actionable explanation of the refusal. */
  readonly message: string
  /** Live document version, when the refusal reported one. */
  readonly currentVersion?: number
  /** Features an edit or save would drop, when the refusal named them. */
  readonly features?: readonly UnsupportedFeature[]
}

/** One accepted document operation. */
export type OfficeToolSuccess<T> = { readonly status: 'ok' } & T

/** The declared result of every Office document tool. */
export type OfficeToolResult<T> = OfficeToolSuccess<T> | OfficeToolFailure

/** A refusal raised inside a tool step and converted to an {@link OfficeToolFailure}. */
export class OfficeToolResultError extends Error {
  /**
   * @param code - stable refusal classification.
   * @param message - actionable explanation of the refusal.
   */
  constructor(
    readonly code: OfficeToolErrorCode,
    message: string,
  ) {
    super(message)
    this.name = 'OfficeToolResultError'
  }
}

/**
 * Build one refusal value.
 * @param code - stable refusal classification.
 * @param message - actionable explanation.
 * @param extras - optional live version or named features.
 * @returns the model-facing refusal.
 */
export function failure(
  code: OfficeToolErrorCode,
  message: string,
  extras: { readonly currentVersion?: number; readonly features?: readonly UnsupportedFeature[] } = {},
): OfficeToolFailure {
  return {
    status: 'error',
    code,
    message,
    ...extras.currentVersion === undefined ? {} : { currentVersion: extras.currentVersion },
    ...extras.features === undefined ? {} : { features: extras.features },
  }
}

/**
 * Convert a thrown Remote or Exchange failure into a model-facing refusal.
 *
 * Callers check cancellation first: an aborted call rejects with the caller's
 * reason so the executor owns cancellation semantics.
 * @param cause - the thrown failure.
 * @returns the matching refusal.
 * @throws the original cause when it carries no stable document classification.
 */
export function mapFailure(cause: unknown): OfficeToolFailure {
  if (cause instanceof OfficeToolResultError) return failure(cause.code, cause.message)
  if (cause instanceof RemoteError) {
    switch (cause.code) {
      case 'office-document/conflict':
        return failure('conflict', cause.message)
      case 'office-document/unsupported-feature':
      case 'office-document/unsupported-format':
        return failure('unsupported-feature', cause.message)
      case 'office-document/out-of-scope':
        return failure('out-of-scope', cause.message)
      case 'office-document/not-found':
        return failure('document-missing', cause.message)
      case 'office-document/bound-exceeded':
        return failure('bound-exceeded', cause.message)
      case 'office-document/conversion-failed':
        return failure('conversion-failed', cause.message)
      case 'gateway/cancelled':
        throw cause
      case 'gateway/internal':
      case 'gateway/bad-request':
        return failure('conversion-failed', cause.message)
    }
  }
  if (cause instanceof OfficeExchangeError) {
    switch (cause.code) {
      case 'unsupported-feature':
      case 'unsupported-format':
        return failure('unsupported-feature', cause.message)
      case 'too-large':
        return failure('bound-exceeded', cause.message)
      case 'out-of-scope':
        return failure('out-of-scope', cause.message)
      case 'conversion-failed':
      case 'not-implemented':
        return failure('conversion-failed', cause.message)
    }
  }
  throw cause
}

/**
 * Run one tool step and turn a classified failure into a refusal value.
 * @param run - the step to execute.
 * @returns the step's result or the refusal it threw.
 */
export async function guarded<T>(run: () => Promise<OfficeToolResult<T>>): Promise<OfficeToolResult<T>> {
  try {
    return await run()
  } catch (cause) {
    return mapFailure(cause)
  }
}
