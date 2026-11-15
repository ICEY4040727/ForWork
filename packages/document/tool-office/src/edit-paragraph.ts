/**
 * `document_edit_paragraph`: replace text-document paragraphs in the live unit,
 * then persist the accepted result with one guarded save.
 * @module @deepseek-ai/dsh-tool-office/edit-paragraph
 */

import type { DocumentEditStep } from '@deepseek-ai/dsh-session'
import { priorParagraph, runEdit } from './edit.ts'
import type { EditOutput, EditPlan } from './edit.ts'
import { failure } from './errors.ts'
import type { OfficeToolFailure, OfficeToolResult } from './errors.ts'
import type { DocumentToolCall } from './runtime.ts'

/** One requested paragraph replacement. */
export interface ParagraphEditInput {
  /** Zero-based paragraph index inside the document. */
  readonly index: number
  /** Paragraph text after the edit. */
  readonly text: string
}

/** Input of one `document_edit_paragraph` call. */
export interface ParagraphEditRequest {
  /** Session-scoped document identity. */
  readonly documentId: string
  /** Content version the caller believes the live unit holds. */
  readonly fromVersion: number
  /** Paragraph replacements, applied in order. */
  readonly operations: readonly ParagraphEditInput[]
}

/**
 * Replace paragraphs in the live document.
 * @param call - runtime, Session, signal, and caps.
 * @param request - document identity, expected version, and paragraph replacements.
 * @returns the new content version and applied replacements, or a refusal.
 */
export async function editParagraphs(
  call: DocumentToolCall,
  request: ParagraphEditRequest,
): Promise<OfficeToolResult<EditOutput<ParagraphEditInput>>> {
  const steps: DocumentEditStep[] = []
  for (const operation of request.operations) {
    const step = stepOf(operation)
    if ('status' in step) return step
    steps.push(step)
  }
  const plan: EditPlan<ParagraphEditInput> = {
    steps,
    applied: request.operations,
    inverse(snapshot) {
      return request.operations.map((operation): DocumentEditStep => ({
        kind: 'paragraph',
        index: operation.index,
        text: priorParagraph(snapshot, operation.index),
      }))
    },
  }
  return await runEdit(call, request.documentId, request.fromVersion, plan)
}

/** Validate one requested replacement and project it to a command step. */
function stepOf(operation: ParagraphEditInput): DocumentEditStep | OfficeToolFailure {
  if (!Number.isSafeInteger(operation.index) || operation.index < 0) {
    return failure('invalid-request', `Paragraph index ${String(operation.index)} is not a non-negative integer.`)
  }
  return { kind: 'paragraph', index: operation.index, text: operation.text }
}
