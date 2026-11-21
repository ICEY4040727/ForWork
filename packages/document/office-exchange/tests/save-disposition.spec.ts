/** Save-disposition grouping and summary behavior. */
import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { open } from '../src/package/ooxml-package.ts'
import {
  buildSaveDisposition,
  byteOperationsByPart,
  sourceDispositions,
} from '../src/source/save-disposition.ts'
import type { ByteSpliceOperation, PartDisposition, SpliceOperation } from '../src/source/splice.ts'

const LIMITS = { maxEntries: 100 }

function sourcePackage() {
  return open(zipSync({
    'a.xml': strToU8('<a/>'),
    'b.xml': strToU8('<b/>'),
    'word/': new Uint8Array(0),
  }), LIMITS).source
}

function replace(part: string, start: number, end: number): ByteSpliceOperation {
  return { kind: 'replace', part, start, end, bytes: strToU8('x') }
}

function insert(part: string, at: number): ByteSpliceOperation {
  return { kind: 'insert', part, at, bytes: strToU8('y') }
}

describe('byteOperationsByPart', () => {
  it('groups byte-changing operations by part and ignores carried entries', () => {
    const operations: readonly SpliceOperation[] = [
      insert('a.xml', 3),
      { kind: 'carry', part: 'a.xml', entry: 'a.xml' },
      replace('a.xml', 1, 2),
      replace('b.xml', 0, 1),
    ]

    const grouped = byteOperationsByPart(operations)

    expect([...grouped.keys()]).toEqual(['a.xml', 'b.xml'])
    expect(grouped.get('a.xml')?.map(operation => operation.kind)).toEqual(['insert', 'replace'])
  })
})

describe('sourceDispositions', () => {
  it('reports verbatim, spliced, and removed entries in ASCII name order', () => {
    const dispositions = sourceDispositions(sourcePackage(), [replace('b.xml', 0, 1)], new Set(['a.xml']))

    expect(dispositions).toEqual([{ name: 'b.xml', status: 'spliced' }])
    expect(sourceDispositions(sourcePackage(), [])).toEqual([
      { name: 'a.xml', status: 'verbatim' },
      { name: 'b.xml', status: 'verbatim' },
    ])
  })
})

describe('buildSaveDisposition', () => {
  it('counts statuses and lists names in ASCII order', () => {
    const entries: readonly PartDisposition[] = [
      { name: 'b.xml', status: 'spliced' },
      { name: 'a.xml', status: 'added' },
      { name: 'c.xml', status: 'removed' },
      { name: 'd.xml', status: 'verbatim' },
    ]

    expect(buildSaveDisposition(entries)).toEqual({
      schemaVersion: 1,
      verbatimCount: 1,
      spliced: { count: 1, names: ['b.xml'] },
      added: { count: 1, names: ['a.xml'] },
      removed: { count: 1, names: ['c.xml'] },
      nameLimit: 256,
    })
  })

  it('compares equal names as equal when a caller repeats one', () => {
    const disposition = buildSaveDisposition([
      { name: 'same.xml', status: 'removed' },
      { name: 'same.xml', status: 'removed' },
    ])

    expect(disposition.removed).toEqual({ count: 2, names: ['same.xml', 'same.xml'] })
  })

  it('orders names in either direction', () => {
    const disposition = buildSaveDisposition([
      { name: 'b.xml', status: 'added' },
      { name: 'a.xml', status: 'added' },
    ])

    expect(disposition.added.names).toEqual(['a.xml', 'b.xml'])
  })
})
