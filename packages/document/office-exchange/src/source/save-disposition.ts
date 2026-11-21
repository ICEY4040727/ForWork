/**
 * Save-disposition reporting shared by the format writers.
 *
 * A writer reports what happened to each source entry and which byte ranges it
 * changed. These helpers turn its applied operations into that report without
 * knowing which format produced them; the caller adds entries it created or
 * removed outside the span plan.
 * @module @forwork/office-exchange/source/save-disposition
 */

import type { SourcePackage } from '../package/types.ts'
import type { SaveDisposition } from '../types.ts'
import type {
  ByteSpliceOperation,
  PartDisposition,
  SpliceOperation,
} from './splice.ts'

/** Cap the save-disposition contract applies to each disposed-name list. */
export const SAVE_DISPOSITION_NAME_LIMIT = 256

/**
 * Group byte-changing operations by the part they touch, preserving operation order.
 * @param operations - Applied operations; carried entries are ignored.
 * @returns One ordered operation list per decoded part name.
 */
export function byteOperationsByPart(
  operations: readonly SpliceOperation[],
): Map<string, ByteSpliceOperation[]> {
  const byPart = new Map<string, ByteSpliceOperation[]>()
  for (const operation of operations) {
    if (operation.kind === 'carry') continue
    const list = byPart.get(operation.part)
    if (list === undefined) byPart.set(operation.part, [operation])
    else list.push(operation)
  }
  return byPart
}

/**
 * Build one disposition per source entry from the operations the writer applied.
 * @param source - Source package whose entries are reported.
 * @param operations - Applied operations across every part.
 * @param removedNames - Entries the caller removed outside the span plan; omitted reports only splices.
 * @returns One disposition per file entry in ASCII name order.
 */
export function sourceDispositions(
  source: SourcePackage,
  operations: readonly SpliceOperation[],
  removedNames: ReadonlySet<string> = new Set(),
): readonly PartDisposition[] {
  const byPart = byteOperationsByPart(operations)
  const dispositions: PartDisposition[] = []
  for (const entry of source.entries) {
    if (entry.name.endsWith('/')) continue
    if (removedNames.has(entry.name)) continue
    dispositions.push({
      name: entry.name,
      status: byPart.has(entry.name) ? 'spliced' : 'verbatim',
    })
  }
  return dispositions.sort((left, right) => compareAscii(left.name, right.name))
}

/**
 * Assemble the contract's disposition payload from per-entry statuses.
 * @param entries - One status per file entry.
 * @returns Exact per-status counts and sorted, capped name lists.
 */
export function buildSaveDisposition(entries: readonly PartDisposition[]): SaveDisposition {
  const spliced = entries.filter(entry => entry.status === 'spliced')
  const added = entries.filter(entry => entry.status === 'added')
  const removed = entries.filter(entry => entry.status === 'removed')
  return {
    schemaVersion: 1,
    verbatimCount: entries.filter(entry => entry.status === 'verbatim').length,
    spliced: { count: spliced.length, names: dispositionNames(spliced) },
    added: { count: added.length, names: dispositionNames(added) },
    removed: { count: removed.length, names: dispositionNames(removed) },
    nameLimit: SAVE_DISPOSITION_NAME_LIMIT,
  }
}

function dispositionNames(entries: readonly PartDisposition[]): readonly string[] {
  return entries
    .map(entry => entry.name)
    .sort(compareAscii)
    .slice(0, SAVE_DISPOSITION_NAME_LIMIT)
}

function compareAscii(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0
}
