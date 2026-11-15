/**
 * Lazy runtime selection. This module carries no Univer import: the editor
 * runtime chunk is requested only when a document mounts, so component tests
 * never load Univer. The chunk stays one dynamic entry because the loader only
 * materializes requested chunks; a sibling chunk shared by several entries
 * would arrive as a synchronous `require` the module table cannot answer.
 */
import type { OfficeRuntimeFactory } from '../adapter.ts'

/**
 * Mount the editor runtime for the opened kind.
 * @param mount - the container, kind, and starting snapshot.
 * @returns the live runtime behind the adapter.
 */
export const loadOfficeRuntime: OfficeRuntimeFactory = async (mount) => {
  const runtime = await import('./runtime.ts')
  return runtime.mount(mount)
}
