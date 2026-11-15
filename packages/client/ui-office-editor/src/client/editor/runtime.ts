/**
 * The one editor runtime chunk: both Univer presets and the wire conversions
 * they share live in this dynamic entry, so the package emits no sibling chunk
 * that its entries would have to require synchronously. Both presets reach the
 * browser only when a document mounts.
 */
import type { OfficeRuntime, OfficeRuntimeMount } from '../adapter.ts'
import { mount as mountDocument } from './docs.ts'
import { mount as mountWorkbook } from './sheets.ts'

/**
 * Mount the sheets or docs runtime for the opened kind.
 * @param input - the container, kind, and starting snapshot.
 * @returns the live runtime behind the adapter.
 */
export function mount(input: OfficeRuntimeMount): Promise<OfficeRuntime> {
  return input.kind === 'workbook' ? mountWorkbook(input) : mountDocument(input)
}
