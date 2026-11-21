/** Validation of a snapshot sidecar against its source package. */
import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { resolveConfig } from '../src/config.ts'
import { OfficeExchangeError } from '../src/errors.ts'
import { open } from '../src/package/ooxml-package.ts'
import { SourceStateError } from '../src/source/source-state.ts'
import { parseSnapshotSourceState } from '../src/source/snapshot-state.ts'

const LIMITS = { maxEntries: 100 }
const DIGEST = 'a'.repeat(64)

function sourcePackage() {
  return open(zipSync({
    'word/document.xml': strToU8('<document>text</document>'),
    'word/': new Uint8Array(0),
  }), LIMITS).source
}

function sidecar(): Record<string, unknown> {
  return {
    schemaVersion: 1,
    format: 'docx',
    packageDigest: DIGEST,
    byteLength: 8,
    sourceParts: ['word/document.xml'],
    anchorMode: 'per-unit',
    anchors: [{
      kind: 'paragraph',
      part: 'word/document.xml',
      key: 'p1',
      start: 0,
      end: 8,
      digest: DIGEST,
    }],
  }
}

describe('parseSnapshotSourceState', () => {
  it('bounds a valid sidecar by the source package parts', () => {
    const state = parseSnapshotSourceState(sidecar(), sourcePackage(), resolveConfig({}), 'Document')

    expect(state.anchorMode).toBe('per-unit')
    expect(state.sourceParts).toEqual(['word/document.xml'])
  })

  it('reports a malformed sidecar as a conversion failure', () => {
    const fail = () => parseSnapshotSourceState({ schemaVersion: 2 }, sourcePackage(), resolveConfig({}), 'Workbook')

    expect(fail).toThrow(OfficeExchangeError)
    expect(fail).toThrow(/Workbook snapshot source state is malformed/u)
    try {
      fail()
    } catch (error) {
      expect(error).toBeInstanceOf(OfficeExchangeError)
      expect((error as OfficeExchangeError).code).toBe('conversion-failed')
    }
  })

  it('rethrows a parser failure that is not a source-state error', () => {
    expect(() => parseSnapshotSourceState(
      sidecar(),
      sourcePackage(),
      { ...resolveConfig({}), maxSourceAnchors: 0 },
      'Document',
    )).toThrow(RangeError)
  })

  it('exposes SourceStateError for direct callers', () => {
    expect(new SourceStateError('invalid-state', 'detail')).toMatchObject({
      name: 'SourceStateError',
      code: 'invalid-state',
      message: 'detail',
    })
  })
})
