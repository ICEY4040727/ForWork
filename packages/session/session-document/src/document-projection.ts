/**
 * The `document` Session projection: the open-document set folded from durable
 * `document/*` facts.
 *
 * The state holds the display order and one entry per open document. `open`
 * inserts or updates an entry, `save` and `edit` advance its content version,
 * `rebind` replaces its path, `missing` marks the path unavailable, and `close`
 * removes the entry. An unrelated event returns the same state reference so the
 * registry performs no downstream work.
 *
 * @module @deepseek-ai/dsh-session-document/document-projection
 */

import { z as zod } from 'zod'
import type { ZodType } from 'zod'
import { DocumentId } from '@deepseek-ai/dsh-session'
import type { SessionEvent } from '@deepseek-ai/dsh-session'
import type { ProjectionDefinition } from '@deepseek-ai/dsh-session-projection'
import type {
  DocumentProjectionEntry,
  DocumentProjectionState,
  DocumentProjectionView,
} from './types.ts'

export type {
  DocumentProjectionEntry,
  DocumentProjectionState,
  DocumentProjectionView,
  DocumentUnsupportedFeature,
} from './types.ts'

const documentEntrySchema: ZodType<DocumentProjectionEntry> = zod.object({
  documentId: zod.string().min(1).transform(value => DocumentId(value)),
  path: zod.string().min(1),
  format: zod.union([zod.literal('docx'), zod.literal('xlsx'), zod.literal('pptx')]),
  contentVersion: zod.number().int().nonnegative(),
  status: zod.union([zod.literal('open'), zod.literal('missing')]),
  unsupported: zod.array(zod.object({
    kind: zod.string().min(1),
    detail: zod.string().optional(),
  })),
})

const documentProjectionStateSchema: ZodType<DocumentProjectionState> = zod.object({
  order: zod.array(zod.string().min(1)),
  documents: zod.record(zod.string().min(1), documentEntrySchema),
}).strict().superRefine((state, context) => {
  if (new Set(state.order).size !== state.order.length) {
    context.addIssue({ code: 'custom', message: 'document order entries must be unique' })
  }
  for (const documentId of state.order) {
    if (!(documentId in state.documents)) {
      context.addIssue({ code: 'custom', message: `document order names "${documentId}" without an entry` })
    }
  }
  for (const documentId of Object.keys(state.documents)) {
    if (!state.order.includes(documentId)) {
      context.addIssue({ code: 'custom', message: `document "${documentId}" is absent from the order` })
    }
  }
})

/**
 * Insert or update one entry, preserving its position when it is already open.
 * @param state - state covering the prior events.
 * @param documentId - identity to upsert.
 * @param entry - complete next entry.
 * @returns the next state, or the same reference when nothing changed.
 */
function upsert(
  state: DocumentProjectionState,
  documentId: string,
  entry: DocumentProjectionEntry,
): DocumentProjectionState {
  const current = state.documents[documentId]
  if (current !== undefined && sameEntry(current, entry)) return state
  const documents = { ...state.documents, [documentId]: entry }
  if (current !== undefined) return { order: state.order, documents }
  return { order: [...state.order, documentId], documents }
}

/**
 * Replace one open entry when the document is open.
 * @param state - state covering the prior events.
 * @param documentId - identity to update.
 * @param update - produces the next entry from the current one.
 * @returns the next state, or the same reference when the document is unknown or unchanged.
 */
function update(
  state: DocumentProjectionState,
  documentId: string,
  update: (entry: DocumentProjectionEntry) => DocumentProjectionEntry,
): DocumentProjectionState {
  const current = state.documents[documentId]
  if (current === undefined) return state
  const next = update(current)
  if (sameEntry(current, next)) return state
  return { order: state.order, documents: { ...state.documents, [documentId]: next } }
}

/** Compare two entries field by field, including their unsupported-feature lists. */
function sameEntry(left: DocumentProjectionEntry, right: DocumentProjectionEntry): boolean {
  return left.documentId === right.documentId
    && left.path === right.path
    && left.format === right.format
    && left.contentVersion === right.contentVersion
    && left.status === right.status
    && JSON.stringify(left.unsupported) === JSON.stringify(right.unsupported)
}

/**
 * Fold one committed Session event into the open-document set.
 * @param state - state covering all prior events.
 * @param event - the next committed Session event.
 * @returns the next state (the same reference for an unrelated event).
 */
export function foldDocumentProjection(
  state: DocumentProjectionState,
  event: SessionEvent,
): DocumentProjectionState {
  switch (event.type) {
    case 'document/open': {
      const data = event.data
      return upsert(state, data.documentId, {
        documentId: data.documentId,
        path: data.path,
        format: data.format,
        contentVersion: data.contentVersion,
        status: 'open',
        unsupported: [],
      })
    }
    case 'document/save':
      return update(state, event.data.documentId, entry => ({
        ...entry,
        contentVersion: Math.max(entry.contentVersion, event.data.contentVersion),
        status: 'open',
      }))
    case 'document/edit':
      return update(state, event.data.documentId, entry => ({
        ...entry,
        contentVersion: Math.max(entry.contentVersion, event.data.toVersion),
        status: 'open',
      }))
    case 'document/rebind':
      return update(state, event.data.documentId, entry => ({
        ...entry,
        path: event.data.toPath,
        status: 'open',
      }))
    case 'document/missing':
      return update(state, event.data.documentId, entry => ({ ...entry, status: 'missing' }))
    case 'document/close': {
      const documentId = event.data.documentId
      if (!(documentId in state.documents)) return state
      const documents = Object.fromEntries(
        Object.entries(state.documents).filter(([id]) => id !== documentId),
      )
      return { order: state.order.filter(id => id !== documentId), documents }
    }
    default:
      return state
  }
}

/** The open-document projection definition registered with `ctx.sessionProjections`. */
export const documentProjectionDefinition = {
  key: 'document',
  stateSchema: documentProjectionStateSchema,
  init: (): DocumentProjectionState => ({ order: [], documents: {} }),
  apply: foldDocumentProjection,
  wire: {
    viewSchema: documentProjectionStateSchema,
    view: (state: DocumentProjectionState): DocumentProjectionView => state,
  },
  stateVersion: 1,
} satisfies ProjectionDefinition<'document', DocumentProjectionState>
