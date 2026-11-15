/**
 * Pure types of the document domain: one open-document entry, the whole
 * open-document set, and the `document` Session-projection key declaration.
 *
 * This module carries no value import, so the Client program binds
 * `useProjection('document')` through the package's `./client` outlet while the
 * Host keeps its value modules (`./document-projection.ts` and the fact
 * writers) behind the package root.
 *
 * @module @deepseek-ai/dsh-session-document/types
 */

import type { Branded } from '@deepseek-ai/dsh-brand'

/**
 * One feature a source file holds that the editable model cannot preserve.
 *
 * This mirrors the Exchange wire value; the projection keeps the fields
 * structural so it depends on no document-format package.
 */
export interface DocumentUnsupportedFeature {
  /** Coarse feature class, such as `macro` or `comment`. */
  readonly kind: string
  /** Location or context of the feature, when known. */
  readonly detail?: string | undefined
}

/**
 * One document in the open set.
 *
 * `documentId` and `format` spell the Session event vocabulary's `DocumentId`
 * brand and `DocumentFormat` union structurally: that vocabulary's module is a
 * Host value module, and this outlet must stay usable from the Client program.
 */
export interface DocumentProjectionEntry {
  /** Stable Session-local identity. */
  readonly documentId: Branded<'DocumentId'>
  /** Current locator; the path may have been rebound. */
  readonly path: string
  /** Editable format detected at open. */
  readonly format: 'docx' | 'xlsx' | 'pptx'
  /** Latest content version established by open, save, or edit. */
  readonly contentVersion: number
  /** Whether the file still resolves. */
  readonly status: 'open' | 'missing'
  /** Features a lossy save would drop. */
  readonly unsupported: readonly DocumentUnsupportedFeature[]
}

/** Open documents in display order plus their current entries. */
export interface DocumentProjectionState {
  /** Document ids in open order. */
  readonly order: readonly string[]
  /** Current entry per document id. */
  readonly documents: Readonly<Record<string, DocumentProjectionEntry>>
}

/**
 * The client view of the open-document set. It is the state itself: the client
 * restores the order and entries it renders without recomputing a fold.
 */
export type DocumentProjectionView = DocumentProjectionState

declare module '@deepseek-ai/dsh-session-projection/types' {
  interface SessionProjectionMap {
    /** Open office documents in this Session, folded from `document/*` facts. */
    document: DocumentProjectionView
  }
}
