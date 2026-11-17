/**
 * Host plugin that registers the `document` projection and contributes the
 * open-document set to the system prompt.
 *
 * The projection is the single source of the open set: the prompt section reads
 * the registering handle for the session of the Agent being assembled and
 * renders nothing when the assembly has no Agent or no document is open.
 *
 * @module @forwork/session-document/plugin
 */

import type { Context } from '@forwork/cordis'
import type {} from '@forwork/agent'
import type {} from '@forwork/session-projection'
import type {} from '@forwork/system-prompt'
import { documentProjectionDefinition } from './document-projection.ts'
import type { DocumentProjectionEntry, DocumentProjectionState } from './document-projection.ts'

/** Cordis plugin name used by loader diagnostics. */
export const name = 'session-document'

/** The projection registry and the prompt registry this plugin contributes to. */
export const inject = ['sessionProjections', 'systemPrompt']

/** Longest readable file name the section repeats before eliding its tail. */
const MAX_DISPLAY_NAME = 40

/** Fixed marker that replaces an elided tail, so the rendered name is deterministic. */
const ELLIPSIS = '...'

/** File name of one document path, independent of the host separator. */
function fileNameOf(path: string): string {
  const normalized = path.replace(/\\/g, '/')
  return normalized.slice(normalized.lastIndexOf('/') + 1)
}

/**
 * Shorten the readable file name; the addressable id keeps the full path.
 * @param name - file name of the logged path.
 * @returns the name, or its leading characters followed by a fixed ellipsis.
 */
function displayNameOf(name: string): string {
  if (name.length <= MAX_DISPLAY_NAME) return name
  return `${name.slice(0, MAX_DISPLAY_NAME - ELLIPSIS.length)}${ELLIPSIS}`
}

/**
 * Render one open entry as an addressable, compact model-facing line.
 *
 * The first field is the exact `documentId` from the projection, so the model
 * can pass it unchanged to the document tools; the parenthesized name, format,
 * and `missing` marker stay for readability.
 * @param position - one-based position in the open order.
 * @param entry - the open document entry.
 * @returns the rendered line.
 */
function renderEntry(position: number, entry: DocumentProjectionEntry): string {
  const markers: string[] = [entry.format]
  if (entry.status === 'missing') markers.push('missing')
  return `${position}. ${entry.documentId} (${displayNameOf(fileNameOf(entry.path))}, ${markers.join(', ')})`
}

/**
 * Render the open-document set for the model.
 *
 * Every line leads with the entry's `documentId` (`<sessionId>:<path>`), which
 * is the identity the document tools accept, so a later turn can address an
 * open document from this text alone. The text derives only from the logged
 * projection state.
 * @param state - folded open-document state for the requesting session.
 * @returns the section text, or `''` when no document is open.
 */
export function renderOpenDocuments(state: DocumentProjectionState): string {
  const lines: string[] = []
  state.order.forEach((documentId, index) => {
    const entry = state.documents[documentId]
    /* v8 ignore next 2 -- stateSchema rejects an order entry without a document. */
    if (entry === undefined) return
    lines.push(renderEntry(index + 1, entry))
  })
  if (lines.length === 0) return ''
  return ['Open documents in this session:', ...lines].join('\n')
}

/**
 * Register the projection and its prompt section on the calling fiber.
 * @param ctx - plugin context; both registrations are disposed with it.
 * @returns when the projection handle is ready and the section is registered.
 */
export async function apply(ctx: Context): Promise<void> {
  const projection = await ctx.sessionProjections.register(documentProjectionDefinition)
  ctx.fiber.assertActive()
  ctx.systemPrompt.section({
    name: 'document:open-set',
    order: ctx.systemPrompt.getSectionOrder('DOCUMENT_OPEN_SET'),
    text: (context) => {
      const session = context.agent?.session
      if (session === undefined) return ''
      return renderOpenDocuments(projection.stateOf(session))
    },
  })
}
