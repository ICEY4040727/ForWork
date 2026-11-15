/**
 * `document_convert`: convert one legacy binary Office file to an editable
 * format inside the Session's workspace root. Conversion is Host-side and
 * never modifies the legacy source.
 * @module @deepseek-ai/dsh-tool-office/convert
 */

import {
  formatOf,
  isEditableFormat,
  OfficeExchangeError,
} from '@deepseek-ai/dsh-office-exchange'
import type { EditableOfficeFormat } from '@deepseek-ai/dsh-office-exchange'
import { confineToolPath } from './documents.ts'
import { failure, guarded } from './errors.ts'
import type { OfficeToolResult } from './errors.ts'
import type { DocumentToolCall } from './runtime.ts'

/** Accepted result of one legacy conversion. */
export interface ConvertOutput {
  /** Converted file path inside the workspace root. */
  readonly path: string
  /** Editable format of the converted file. */
  readonly format: EditableOfficeFormat
}

/**
 * Convert one legacy file to `.docx` or `.xlsx`.
 * @param call - runtime, Session, signal, and caps.
 * @param path - absolute or workspace-relative legacy path.
 * @returns the converted path and format, or a structured refusal.
 */
export async function convertDocument(
  call: DocumentToolCall,
  path: string,
): Promise<OfficeToolResult<ConvertOutput>> {
  return await guarded(async () => {
    if (path.trim().length === 0) return failure('invalid-request', 'Conversion requires a non-empty path.')
    const absolute = await confineToolPath(path, call.scope.workspaceRoot)
    const converted = await call.runtime.documents.convertLegacy(call.scope, absolute, call.signal)
    const format = formatOf(converted)
    if (format === undefined || !isEditableFormat(format)) {
      throw new OfficeExchangeError('conversion-failed', `Conversion produced "${converted}", which is not editable.`)
    }
    return { status: 'ok', path: converted, format }
  })
}
