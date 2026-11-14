/** Pure format classification for the Exchange seam. */
import type { EditableOfficeFormat, LegacyOfficeFormat, OfficeFormat } from './types.ts'

const EDITABLE: ReadonlySet<OfficeFormat> = new Set<EditableOfficeFormat>(['docx', 'xlsx'])

const LEGACY: ReadonlySet<OfficeFormat> = new Set<LegacyOfficeFormat>(['doc', 'xls', 'ppt'])

/**
 * Map a path's extension to a supported Office format, or `undefined`.
 * @param path - File path whose final extension is classified.
 * @returns The matching Office format, or `undefined` for a missing or unsupported extension.
 */
export function formatOf(path: string): OfficeFormat | undefined {
  const dot = path.lastIndexOf('.')
  if (dot < 0) return undefined
  const suffix = path.slice(dot + 1).toLowerCase()
  switch (suffix) {
    case 'docx':
    case 'xlsx':
    case 'pptx':
    case 'doc':
    case 'xls':
    case 'ppt':
      return suffix
    default:
      return undefined
  }
}

/**
 * Whether a format is directly editable without conversion.
 * @param format - Format to classify.
 * @returns `true` for `docx` and `xlsx`, narrowing the format.
 */
export function isEditableFormat(format: OfficeFormat): format is EditableOfficeFormat {
  return EDITABLE.has(format)
}

/**
 * Whether a format must be converted to an editable format before opening.
 * @param format - Format to classify.
 * @returns `true` for `doc`, `xls`, and `ppt`, narrowing the format.
 */
export function isLegacyFormat(format: OfficeFormat): format is LegacyOfficeFormat {
  return LEGACY.has(format)
}
