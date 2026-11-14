/**
 * Lazy runtime selection. This module carries no Univer import: each kind's
 * preset module is requested only when a document of that kind mounts, so the
 * browser fetches one editor chunk and component tests never load either.
 */
import type { OfficeRuntimeFactory } from '../adapter.ts'

/**
 * Mount the sheets or docs runtime for the opened kind.
 * @param mount - the container, kind, and starting snapshot.
 * @returns the live runtime behind the adapter.
 */
export const loadOfficeRuntime: OfficeRuntimeFactory = async (mount) => {
  const runtime = mount.kind === 'workbook'
    ? await import('./sheets.ts')
    : await import('./docs.ts')
  return runtime.mount(mount)
}
