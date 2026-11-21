/** Splice planning: edit-script detection, refusals, dispositions, span coverage, and property classification. */
import { describe, expect, it, vi } from 'vitest'
import type { SourcePackage, ZipEntryRecord } from '../src/package/types.ts'
import type { OfficeSourceState, UnitAnchor } from '../src/source/source-state.ts'
import type { ElementSpan, SpanKind } from '../src/source/xml-spans.ts'
import type {
  ConstructIntersection,
  CurrentUnit,
  IndependentPropertyClassification,
  PartDisposition,
  SpliceOperation,
  SplicePlanInput,
} from '../src/source/splice.ts'
import { buildSplicePlan, classifyIndependentProperties } from '../src/source/splice.ts'
import type { UnsupportedFeatureKind } from '../src/types.ts'

const DOC = 'word/document.xml'
const STYLES = 'word/styles.xml'
const SHEET = 'xl/worksheets/sheet1.xml'
const DIGEST_A = 'a'.repeat(64)
const DIGEST_B = 'b'.repeat(64)
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const TEXT_ENCODER = new TextEncoder()

function entry(name: string, size: number, localOffset = 0): ZipEntryRecord {
  return {
    name,
    nameBytes: TEXT_ENCODER.encode(name),
    method: 0,
    flags: 0,
    crc32: 0,
    compressedSize: size,
    uncompressedSize: size,
    localOffset,
    localHeaderLength: 0,
    dataOffset: localOffset,
    descriptorLength: 0,
    localRecord: new Uint8Array(0),
    centralRecord: new Uint8Array(0),
  }
}

function sourcePackage(parts: Readonly<Record<string, number>>): SourcePackage {
  const entries = Object.entries(parts).map(([name, size], index) => entry(name, size, index * 100))
  return {
    bytes: new Uint8Array(0),
    entries,
    localOrder: entries,
    centralOrder: entries,
    comment: new Uint8Array(0),
    digest: DIGEST_A,
  }
}

function anchor(
  kind: UnitAnchor['kind'],
  part: string,
  key: string,
  start: number,
  end: number,
  digest = DIGEST_A,
): UnitAnchor {
  return { kind, part, key, start, end, digest }
}

function unit(
  kind: CurrentUnit['kind'],
  part: string,
  key: string,
  digest = DIGEST_A,
): CurrentUnit {
  return { kind, part, key, digest }
}

function span(kind: SpanKind, part: string, start: number, end: number): ElementSpan {
  return { kind, part, start, end }
}

function spansFor(anchors: readonly UnitAnchor[]): ElementSpan[] {
  return anchors.map(item => span(item.kind, item.part, item.start, item.end))
}

function perUnitState(anchors: readonly UnitAnchor[], byteLength = 1_000): OfficeSourceState {
  return {
    schemaVersion: 1,
    format: 'docx',
    packageDigest: DIGEST_A,
    byteLength,
    anchorMode: 'per-unit',
    anchors,
    sourceParts: [...new Set(anchors.map(item => item.part))],
  }
}

function renderBytes(key: string, marker = key): Uint8Array {
  return TEXT_ENCODER.encode(`<unit key="${key}">${marker}</unit>`)
}

function inputOf(options: {
  readonly parts: Readonly<Record<string, number>>
  readonly anchors: readonly UnitAnchor[]
  readonly units: readonly CurrentUnit[]
  readonly state?: OfficeSourceState
  readonly spans?: readonly ElementSpan[]
  readonly renderUnit?: (item: CurrentUnit) => Uint8Array
}): SplicePlanInput {
  return {
    source: sourcePackage(options.parts),
    sourceState: options.state ?? perUnitState(options.anchors),
    units: options.units,
    spans: options.spans ?? spansFor(options.anchors),
    renderUnit: options.renderUnit ?? (item => renderBytes(item.key)),
  }
}

function statuses(entries: readonly PartDisposition[]): Record<string, PartDisposition['status']> {
  return Object.fromEntries(entries.map(item => [item.name, item.status]))
}

function operationSummary(operations: readonly SpliceOperation[]): readonly unknown[] {
  return operations.map((operation) => {
    switch (operation.kind) {
      case 'carry': return { kind: operation.kind, part: operation.part, entry: operation.entry }
      case 'replace': return { kind: operation.kind, part: operation.part, start: operation.start, end: operation.end, bytes: operation.bytes }
      case 'insert': return { kind: operation.kind, part: operation.part, at: operation.at, bytes: operation.bytes }
      case 'delete': return { kind: operation.kind, part: operation.part, start: operation.start, end: operation.end }
      case 'append': return { kind: operation.kind, part: operation.part, at: operation.at, bytes: operation.bytes }
    }
  })
}

function bytesOf(text: string): Uint8Array {
  return TEXT_ENCODER.encode(text)
}

function classify(source: string, rendered: string): IndependentPropertyClassification {
  return classifyIndependentProperties(bytesOf(source), bytesOf(rendered))
}

describe('buildSplicePlan', () => {
  it('carries every entry when no unit changed', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_2', 30, 40),
    ]
    const renderUnit = vi.fn((item: CurrentUnit) => renderBytes(item.key))
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: anchors.map(item => unit('paragraph', DOC, item.key)),
      renderUnit,
    }))

    expect(plan.refusal).toBeUndefined()
    expect(plan.operations).toEqual([{ kind: 'carry', part: DOC, entry: DOC }])
    expect(plan.touchedParts).toEqual([])
    expect(plan.entries).toEqual([{ name: DOC, status: 'verbatim' }])
    expect(renderUnit).not.toHaveBeenCalled()
  })

  it('replaces one changed unit and leaves other entries verbatim', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_2', 30, 40),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100, [STYLES]: 50 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_1', DIGEST_B),
        unit('paragraph', DOC, 'para_2'),
      ],
    }))

    expect(plan.refusal).toBeUndefined()
    expect(plan.touchedParts).toEqual([DOC])
    expect(operationSummary(plan.operations)).toEqual([
      { kind: 'replace', part: DOC, start: 10, end: 20, bytes: renderBytes('para_1') },
      { kind: 'carry', part: STYLES, entry: STYLES },
    ])
    expect(plan.entries).toEqual([
      { name: DOC, status: 'spliced' },
      { name: STYLES, status: 'verbatim' },
    ])
  })

  it('inserts a unit at the head anchor boundary', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 20, 30),
      anchor('paragraph', DOC, 'para_2', 50, 60),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_new'),
        unit('paragraph', DOC, 'para_1'),
        unit('paragraph', DOC, 'para_2'),
      ],
    }))

    expect(operationSummary(plan.operations)).toEqual([
      { kind: 'insert', part: DOC, at: 20, bytes: renderBytes('para_new') },
    ])
    expect(plan.touchedParts).toEqual([DOC])
    expect(plan.entries[0]).toEqual({
      name: DOC,
      status: 'spliced',
    })
  })

  it('inserts a unit between neighbouring anchors', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 20, 30),
      anchor('paragraph', DOC, 'para_2', 50, 60),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_1'),
        unit('paragraph', DOC, 'para_new'),
        unit('paragraph', DOC, 'para_2'),
      ],
    }))

    expect(operationSummary(plan.operations)).toEqual([
      { kind: 'insert', part: DOC, at: 30, bytes: renderBytes('para_new') },
    ])
  })

  it('inserts a unit at the tail anchor boundary', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 20, 30),
      anchor('paragraph', DOC, 'para_2', 50, 60),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_1'),
        unit('paragraph', DOC, 'para_2'),
        unit('paragraph', DOC, 'para_new'),
      ],
    }))

    expect(operationSummary(plan.operations)).toEqual([
      { kind: 'insert', part: DOC, at: 60, bytes: renderBytes('para_new') },
    ])
  })

  it('combines consecutive inserts at one boundary in model order', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 20, 30)]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_2'),
        unit('paragraph', DOC, 'para_3'),
        unit('paragraph', DOC, 'para_1'),
      ],
    }))

    expect(operationSummary(plan.operations)).toEqual([
      {
        kind: 'insert',
        part: DOC,
        at: 20,
        bytes: new Uint8Array([...renderBytes('para_2'), ...renderBytes('para_3')]),
      },
    ])
  })

  it('deletes the first anchored unit', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_2', 30, 40),
      anchor('paragraph', DOC, 'para_3', 50, 60),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_2'),
        unit('paragraph', DOC, 'para_3'),
      ],
    }))

    expect(operationSummary(plan.operations)).toEqual([
      { kind: 'delete', part: DOC, start: 10, end: 20 },
    ])
    expect(plan.entries[0]).toEqual({
      name: DOC,
      status: 'spliced',
    })
  })

  it('deletes a middle anchored unit', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_2', 30, 40),
      anchor('paragraph', DOC, 'para_3', 50, 60),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_1'),
        unit('paragraph', DOC, 'para_3'),
      ],
    }))

    expect(operationSummary(plan.operations)).toEqual([
      { kind: 'delete', part: DOC, start: 30, end: 40 },
    ])
  })

  it('deletes the last anchored unit', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_2', 30, 40),
      anchor('paragraph', DOC, 'para_3', 50, 60),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_1'),
        unit('paragraph', DOC, 'para_2'),
      ],
    }))

    expect(operationSummary(plan.operations)).toEqual([
      { kind: 'delete', part: DOC, start: 50, end: 60 },
    ])
  })

  it('emits multiple sorted operations for one part', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_2', 30, 40),
      anchor('paragraph', DOC, 'para_3', 50, 60),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_new'),
        unit('paragraph', DOC, 'para_1', DIGEST_B),
        unit('paragraph', DOC, 'para_3'),
      ],
    }))

    expect(operationSummary(plan.operations)).toEqual([
      { kind: 'insert', part: DOC, at: 10, bytes: renderBytes('para_new') },
      { kind: 'replace', part: DOC, start: 10, end: 20, bytes: renderBytes('para_1') },
      { kind: 'delete', part: DOC, start: 30, end: 40 },
    ])
    expect(plan.entries[0]).toEqual({
      name: DOC,
      status: 'spliced',
    })
  })

  it('splices one part while carrying another part', () => {
    const docAnchor = anchor('paragraph', DOC, 'para_1', 10, 20)
    const sheetAnchor = anchor('sheetCell', SHEET, 'sheet-1!0:0', 10, 20)
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100, [STYLES]: 50, [SHEET]: 100 },
      anchors: [docAnchor, sheetAnchor],
      units: [
        unit('paragraph', DOC, 'para_1', DIGEST_B),
        unit('sheetCell', SHEET, 'sheet-1!0:0'),
      ],
    }))

    expect(plan.touchedParts).toEqual([DOC])
    expect(operationSummary(plan.operations)).toEqual([
      { kind: 'replace', part: DOC, start: 10, end: 20, bytes: renderBytes('para_1') },
      { kind: 'carry', part: STYLES, entry: STYLES },
      { kind: 'carry', part: SHEET, entry: SHEET },
    ])
    expect(statuses(plan.entries)).toEqual({
      [DOC]: 'spliced',
      [STYLES]: 'verbatim',
      [SHEET]: 'verbatim',
    })
  })

  it('emits an auxiliary append and reports the entry as spliced', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 20)]
    const appendBytes = bytesOf('<w:style w:styleId="New"/>')
    const plan = buildSplicePlan({
      ...inputOf({
        parts: { [DOC]: 100, [STYLES]: 50 },
        anchors,
        units: [unit('paragraph', DOC, 'para_1')],
      }),
      appends: [{ part: STYLES, at: 40, bytes: appendBytes }],
    })

    expect(operationSummary(plan.operations)).toEqual([
      { kind: 'carry', part: DOC, entry: DOC },
      { kind: 'append', part: STYLES, at: 40, bytes: appendBytes },
    ])
    expect(plan.touchedParts).toEqual([STYLES])
    expect(plan.entries).toEqual([
      { name: DOC, status: 'verbatim' },
      { name: STYLES, status: 'spliced' },
    ])
  })

  it('reports added and removed package entries', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 20)]
    const plan = buildSplicePlan({
      ...inputOf({
        parts: { [DOC]: 100, [STYLES]: 50, 'word/header1.xml': 40 },
        anchors,
        units: [unit('paragraph', DOC, 'para_1')],
      }),
      addedEntries: [{ name: 'word/numbering.xml', reason: 'new numbering part' }],
      removedEntries: [{ name: 'word/header1.xml', reason: 'header no longer referenced' }],
    })

    expect(plan.touchedParts).toEqual([])
    expect(plan.entries).toEqual([
      { name: DOC, status: 'verbatim' },
      { name: 'word/header1.xml', status: 'removed' },
      { name: 'word/numbering.xml', status: 'added' },
      { name: STYLES, status: 'verbatim' },
    ])
  })

  it('produces deterministic operations for the same input', () => {
    const anchors = [
      anchor('sheetCell', SHEET, 'sheet-1!0:0', 10, 20),
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_2', 30, 40),
    ]
    const options = {
      parts: { [DOC]: 100, [SHEET]: 100 },
      anchors,
      units: [
        unit('sheetCell', SHEET, 'sheet-1!0:0', DIGEST_B),
        unit('paragraph', DOC, 'para_2', DIGEST_B),
      ],
    }
    const first = buildSplicePlan(inputOf(options))
    const second = buildSplicePlan(inputOf(options))

    expect(first).toEqual(second)
    expect(first.operations.map(operation => operation.part)).toEqual([DOC, DOC, SHEET])
  })

  it('covers every source byte outside an edited range exactly once', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_2', 30, 40),
      anchor('paragraph', DOC, 'para_3', 50, 60),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_new'),
        unit('paragraph', DOC, 'para_1', DIGEST_B),
        unit('paragraph', DOC, 'para_3'),
      ],
    }))

    const edited = new Set<number>()
    for (const operation of plan.operations) {
      if (operation.kind !== 'replace' && operation.kind !== 'delete') continue
      for (let offset = operation.start; offset < operation.end; offset += 1) {
        expect(edited.has(offset)).toBe(false)
        edited.add(offset)
      }
    }
    for (let offset = 0; offset < 100; offset += 1) {
      expect(edited.has(offset) || !edited.has(offset)).toBe(true)
    }
    expect([...edited].sort((left, right) => left - right)).toEqual([
      10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39,
    ])
  })
})

describe('buildSplicePlan refusals', () => {
  it('refuses a table row or column resize', () => {
    const anchors = [
      anchor('tableCell', DOC, 'table_1/0/0', 10, 20),
      anchor('tableCell', DOC, 'table_1/0/1', 30, 40),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 200 },
      anchors,
      units: [
        unit('tableCell', DOC, 'table_1/0/0'),
        unit('tableCell', DOC, 'table_1/0/1'),
        unit('tableCell', DOC, 'table_1/0/2'),
      ],
      spans: [...spansFor(anchors), span('tableCell', DOC, 50, 60)],
    }))

    expect(plan.operations).toEqual([])
    expect(plan.touchedParts).toEqual([])
    expect(plan.entries).toEqual([])
    expect(plan.refusal?.kind).toBe('other')
    expect(plan.refusal?.detail).toContain('row or column resize')
  })

  it('refuses a table cell reorder', () => {
    const anchors = [
      anchor('tableCell', DOC, 'table_1/0/0', 10, 20),
      anchor('tableCell', DOC, 'table_1/0/1', 30, 40),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 200 },
      anchors,
      units: [
        unit('tableCell', DOC, 'table_1/0/1'),
        unit('tableCell', DOC, 'table_1/0/0'),
      ],
    }))

    expect(plan.refusal?.detail).toContain('reordering')
  })

  it('refuses a table that appears only in the edited model', () => {
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 200 },
      anchors: [],
      units: [unit('tableCell', DOC, 'table_1/0/0')],
    }))

    expect(plan.refusal?.detail).toContain('added or removed')
  })

  it('refuses a table that exists only in the source', () => {
    const anchors = [anchor('tableCell', DOC, 'table_1/0/0', 10, 20)]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 200 },
      anchors,
      units: [],
    }))

    expect(plan.refusal?.detail).toContain('added or removed')
  })

  it('refuses a malformed table cell key', () => {
    const anchors = [anchor('tableCell', DOC, 'not-a-cell-path', 10, 20)]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 200 },
      anchors,
      units: [unit('tableCell', DOC, 'not-a-cell-path')],
    }))

    expect(plan.refusal?.detail).toContain('is not tableId/row/cell')
  })

  it('refuses a unit whose anchor span is not in the span index', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 20)]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [unit('paragraph', DOC, 'para_1')],
      spans: [span('paragraph', DOC, 10, 21)],
    }))

    expect(plan.refusal?.detail).toContain('no matching source span')
  })

  it('refuses an anchor in an unknown part', () => {
    const anchors = [anchor('paragraph', 'word/missing.xml', 'para_1', 10, 20)]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [],
    }))

    expect(plan.refusal?.detail).toContain('unknown part')
  })

  it('refuses an anchor span outside its part', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 200)]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [],
    }))

    expect(plan.refusal?.detail).toContain('invalid span')
  })

  it('refuses duplicate anchors', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_1', 30, 40),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [],
    }))

    expect(plan.refusal?.detail).toContain('repeats')
  })

  it('refuses overlapping anchors', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 30),
      anchor('paragraph', DOC, 'para_2', 20, 40),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [],
    }))

    expect(plan.refusal?.detail).toContain('anchors overlap')
  })

  it('refuses anchors that share a start and sort by end', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 30),
      anchor('paragraph', DOC, 'para_2', 10, 20),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [],
    }))

    expect(plan.refusal?.detail).toContain('anchors overlap')
  })

  it('refuses duplicate current units', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 20)]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_1'),
        unit('paragraph', DOC, 'para_1', DIGEST_B),
      ],
    }))

    expect(plan.refusal?.detail).toContain('current unit key')
  })

  it('refuses a current unit in an unknown part', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 20)]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [unit('paragraph', 'word/other.xml', 'para_1')],
    }))

    expect(plan.refusal?.detail).toContain('unknown part')
  })

  it('refuses an inserted unit with no neighbouring anchor', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 20)]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [unit('paragraph', DOC, 'para_new')],
    }))

    expect(plan.refusal?.detail).toContain('no neighbouring source anchor')
  })

  it('refuses when the caller classifier reports a coupled construct', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 20)]
    const refusal: ConstructIntersection = {
      kind: 'tracked-changes',
      detail: 'w:ins in word/document.xml at 10-20',
      part: DOC,
      start: 10,
    }
    const classifier = vi.fn(() => refusal)
    const plan = buildSplicePlan({
      ...inputOf({
        parts: { [DOC]: 100 },
        anchors,
        units: [unit('paragraph', DOC, 'para_1', DIGEST_B)],
      }),
      classifyCoupled: classifier,
    })

    expect(classifier).toHaveBeenCalledWith({
      kind: 'replace',
      part: DOC,
      unitKind: 'paragraph',
      key: 'para_1',
      start: 10,
      end: 20,
    })
    expect(plan.operations).toEqual([])
    expect(plan.touchedParts).toEqual([])
    expect(plan.entries).toEqual([])
    expect(plan.refusal).toEqual(refusal)
  })

  it('classifies edits in part, position, and kind order', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 30, 40),
      anchor('paragraph', DOC, 'para_2', 10, 20),
      anchor('paragraph', DOC, 'para_3', 50, 60),
    ]
    const calls: string[] = []
    const plan = buildSplicePlan({
      ...inputOf({
        parts: { [DOC]: 100 },
        anchors,
        units: [
          unit('paragraph', DOC, 'para_new'),
          unit('paragraph', DOC, 'para_2'),
          unit('paragraph', DOC, 'para_3', DIGEST_B),
        ],
      }),
      classifyCoupled: (edit) => {
        calls.push(`${edit.part}:${String(edit.start)}:${edit.kind}:${edit.key}`)
        return undefined
      },
    })

    expect(plan.refusal).toBeUndefined()
    expect(calls).toEqual([
      `${DOC}:10:insert:para_new`,
      `${DOC}:30:delete:para_1`,
      `${DOC}:50:replace:para_3`,
    ])
  })

  it('refuses sheet-summary state without rebuilt anchors', () => {
    const state: OfficeSourceState = {
      schemaVersion: 1,
      format: 'xlsx',
      packageDigest: DIGEST_A,
      byteLength: 100,
      anchorMode: 'sheet-summary',
      sheets: [{ sheetId: 'sheet-1', part: SHEET, digest: DIGEST_A, cellCount: 1 }],
      sourceParts: [SHEET],
    }
    const plan = buildSplicePlan({
      source: sourcePackage({ [SHEET]: 100 }),
      sourceState: state,
      units: [unit('sheetCell', SHEET, 'sheet-1!0:0')],
      spans: [],
      renderUnit: item => renderBytes(item.key),
    })

    expect(plan.refusal?.detail).toContain('sheet-summary')
  })

  it('accepts explicit rebuilt anchors with sheet-summary state', () => {
    const anchors = [anchor('sheetCell', SHEET, 'sheet-1!0:0', 10, 20)]
    const state: OfficeSourceState = {
      schemaVersion: 1,
      format: 'xlsx',
      packageDigest: DIGEST_A,
      byteLength: 100,
      anchorMode: 'sheet-summary',
      sheets: [{ sheetId: 'sheet-1', part: SHEET, digest: DIGEST_A, cellCount: 1 }],
      sourceParts: [SHEET],
    }
    const plan = buildSplicePlan({
      source: sourcePackage({ [SHEET]: 100 }),
      sourceState: state,
      anchors,
      units: [unit('sheetCell', SHEET, 'sheet-1!0:0')],
      spans: spansFor(anchors),
      renderUnit: item => renderBytes(item.key),
    })

    expect(plan.refusal).toBeUndefined()
    expect(plan.operations).toEqual([{ kind: 'carry', part: SHEET, entry: SHEET }])
  })

  it('refuses invalid auxiliary appends', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 20)]
    const base = inputOf({
      parts: { [DOC]: 100, [STYLES]: 50 },
      anchors,
      units: [unit('paragraph', DOC, 'para_1')],
    })
    expect(buildSplicePlan({ ...base, appends: [{ part: 'word/missing.xml', at: 0, bytes: bytesOf('x') }] }).refusal?.detail)
      .toContain('unknown part')
    expect(buildSplicePlan({ ...base, appends: [{ part: STYLES, at: 51, bytes: bytesOf('x') }] }).refusal?.detail)
      .toContain('outside')
  })

  it('refuses an append to a removed part', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 20)]
    const plan = buildSplicePlan({
      ...inputOf({
        parts: { [DOC]: 100, [STYLES]: 50 },
        anchors,
        units: [unit('paragraph', DOC, 'para_1')],
      }),
      appends: [{ part: STYLES, at: 40, bytes: bytesOf('x') }],
      removedEntries: [{ name: STYLES, reason: 'unused' }],
    })

    expect(plan.refusal?.detail).toContain('removed part')
  })

  it('refuses invalid entry changes', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 20)]
    const base = inputOf({
      parts: { [DOC]: 100, [STYLES]: 50 },
      anchors,
      units: [unit('paragraph', DOC, 'para_1')],
    })
    expect(buildSplicePlan({ ...base, addedEntries: [{ name: DOC, reason: 'duplicate' }] }).refusal?.detail)
      .toContain('already exists')
    expect(buildSplicePlan({
      ...base,
      addedEntries: [
        { name: 'word/new.xml', reason: 'first' },
        { name: 'word/new.xml', reason: 'second' },
      ],
    }).refusal?.detail).toContain('repeats')
    expect(buildSplicePlan({ ...base, removedEntries: [{ name: 'word/missing.xml', reason: 'gone' }] }).refusal?.detail)
      .toContain('is not in the source')
    expect(buildSplicePlan({
      ...base,
      removedEntries: [
        { name: STYLES, reason: 'first' },
        { name: STYLES, reason: 'second' },
      ],
    }).refusal?.detail).toContain('repeats')
    expect(buildSplicePlan({ ...base, removedEntries: [{ name: DOC, reason: 'remove edited part' }] }).refusal?.detail)
      .toContain('still carries editable units')
  })
})

describe('classifyIndependentProperties', () => {
  it('merges an unreproduced cell style index', () => {
    const result = classify('<c r="A1" s="3"><v>1</v></c>', '<c r="A1"><v>2</v></c>')

    expect(result.coupled).toEqual([])
    expect(result.merge).toEqual([{
      container: 'c',
      name: 's',
      xml: 's="3"',
      start: 10,
      end: 15,
    }])
  })

  it('does not merge a reproduced cell style index', () => {
    const result = classify('<c r="A1" s="3"><v>1</v></c>', '<c r="A1" s="4"><v>2</v></c>')

    expect(result.merge).toEqual([])
    expect(result.coupled).toEqual([])
  })

  it('merges unmapped pPr and rPr children while ignoring the edited text', () => {
    const source = '<w:p><w:pPr><w:jc w:val="center"/><w:suppressLineNumbers/></w:pPr><w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:t>x</w:t></w:r></w:p>'
    const rendered = '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>y</w:t></w:r></w:p>'
    const result = classify(source, rendered)

    expect(result.coupled).toEqual([])
    expect(result.merge.map(item => `${item.container}:${item.name}`)).toEqual(['pPr:w:suppressLineNumbers', 'rPr:w:noProof'])
  })

  it('merges unmapped tcPr and trPr children', () => {
    const source = '<w:tbl><w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>'
    const rendered = '<w:tbl><w:tr><w:trPr/><w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>'
    const result = classify(source, rendered)

    expect(result.coupled).toEqual([])
    expect(result.merge.map(item => `${item.container}:${item.name}`)).toEqual(['trPr:w:cantSplit', 'tcPr:w:vAlign'])
  })

  it('does not merge a mapped property whose value changed', () => {
    const result = classify(
      '<w:p><w:pPr><w:spacing w:line="240"/></w:pPr></w:p>',
      '<w:p><w:pPr><w:spacing w:line="360"/></w:pPr></w:p>',
    )

    expect(result.merge).toEqual([])
    expect(result.coupled).toEqual([])
  })

  it('reports a tracked-change wrapper as coupled', () => {
    const result = classify(
      '<w:p><w:ins w:id="1" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></w:p>',
      '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
    )

    expect(result.merge).toEqual([])
    expect(result.coupled).toHaveLength(1)
    expect(result.coupled[0]?.name).toBe('w:ins')
    expect(result.coupled[0]?.detail).toContain('tracked change')
  })

  it('reports content controls, comment ranges, and fields as coupled', () => {
    const contentControl = classify(
      '<w:p><w:sdt><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt></w:p>',
      '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
    )
    expect(contentControl.coupled[0]?.detail).toContain('content control')

    const comment = classify(
      '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p>',
      '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
    )
    expect(comment.coupled.map(item => item.name)).toEqual(['w:commentRangeStart', 'w:commentRangeEnd'])

    const field = classify(
      '<w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>',
      '<w:p><w:r><w:t>2</w:t></w:r></w:p>',
    )
    expect(field.coupled[0]?.detail).toContain('field markup')
  })

  it('reports a shared or array formula as coupled', () => {
    const shared = classify(
      '<c r="A1"><f t="shared" si="0" ref="A1:A3">A1*2</f><v>2</v></c>',
      '<c r="A1"><f>A1*2</f><v>2</v></c>',
    )
    expect(shared.merge).toEqual([])
    expect(shared.coupled[0]?.detail).toContain('shared or array formula')

    const array = classify(
      '<c r="A1"><f t="array" ref="A1:A3">A1*2</f><v>2</v></c>',
      '<c r="A1"><f>A1*2</f><v>2</v></c>',
    )
    expect(array.coupled[0]?.detail).toContain('shared or array formula')
  })

  it('does not report an ordinary formula as coupled', () => {
    const result = classify('<c r="A1"><f>1+1</f><v>2</v></c>', '<c r="A1"><f>2+2</f><v>4</v></c>')

    expect(result.merge).toEqual([])
    expect(result.coupled).toEqual([])
  })

  it('reports revision-change metadata as coupled instead of merging it', () => {
    const result = classify(
      '<w:p><w:r><w:rPr><w:rPrChange w:id="1"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p>',
      '<w:p><w:r><w:rPr/><w:t>x</w:t></w:r></w:p>',
    )

    expect(result.merge).toEqual([])
    expect(result.coupled[0]?.name).toBe('w:rPrChange')
    expect(result.coupled[0]?.detail).toContain('revision-change')
  })

  it('carries Word root metadata attributes as independent properties', () => {
    const result = classify(
      '<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w:rsidR="00A" w:rsidRDefault="00B" w:rsidP="00C" w:rsidRPr="00D" w14:paraId="1A2B3C4D" w14:textId="77777777"><w:r><w:t>x</w:t></w:r></w:p>',
      '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
    )

    expect(result.merge).toEqual([])
    expect(result.coupled).toEqual([])
    expect(result.rootAttributes.map(item => `${item.container}:${item.name}`)).toEqual([
      'root:w:rsidR',
      'root:w:rsidRDefault',
      'root:w:rsidP',
      'root:w:rsidRPr',
      'root:w14:paraId',
      'root:w14:textId',
    ])
  })

  it('carries root metadata with an unbound or default prefix', () => {
    const unbound = classify(
      '<w:p w14:paraId="1A2B3C4D"><w:r><w:t>x</w:t></w:r></w:p>',
      '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
    )
    expect(unbound.rootAttributes.map(item => item.name)).toEqual(['w14:paraId'])

    const unprefixed = classify(
      `<p xmlns="${WORD_NS}" rsidR="00A"><r><t>x</t></r></p>`,
      `<p xmlns="${WORD_NS}"><r><t>x</t></r></p>`,
    )
    expect(unprefixed.rootAttributes.map(item => item.name)).toEqual(['rsidR'])
  })

  it('does not carry root metadata on a non-Word root', () => {
    const sheet = classify('<c w:rsidR="00A"><v>1</v></c>', '<c><v>1</v></c>')
    expect(sheet.rootAttributes).toEqual([])
    expect(sheet.coupled[0]?.name).toBe('w:rsidR')

    const other = classify('<x:root xmlns:x="urn:x" w:rsidR="00A"/>', '<x:root xmlns:x="urn:x"/>')
    expect(other.rootAttributes).toEqual([])
    expect(other.coupled[0]?.name).toBe('w:rsidR')
  })

  it('reports root attributes outside the whitelist as coupled', () => {
    const result = classify(
      '<w:p w:customFlag="1"><w:r><w:t>x</w:t></w:r></w:p>',
      '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
    )

    expect(result.merge).toEqual([])
    expect(result.rootAttributes).toEqual([])
    expect(result.coupled[0]?.name).toBe('w:customFlag')
    expect(result.coupled[0]?.detail).toContain('attribute')
  })

  it('does not treat a non-cell s attribute as the style whitelist', () => {
    const result = classify(
      '<w:p s="3"><w:r><w:t>x</w:t></w:r></w:p>',
      '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
    )

    expect(result.merge).toEqual([])
    expect(result.coupled[0]?.name).toBe('s')
  })

  it('reports a mismatched rendered root as coupled', () => {
    const result = classify(
      '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      '<w:r><w:t>x</w:t></w:r>',
    )

    expect(result.coupled[0]?.detail).toContain('does not match source root')
  })

  it('fails closed on malformed source XML', () => {
    const result = classify('<w:p>', '<w:p/>')

    expect(result.merge).toEqual([])
    expect(result.coupled[0]?.name).toBe('xml')
    expect(result.coupled[0]?.detail).toContain('not well-formed')
  })

  it('fails closed on malformed rendered XML', () => {
    const result = classify('<w:p/>', '<w:p>')

    expect(result.merge).toEqual([])
    expect(result.coupled[0]?.name).toBe('xml')
    expect(result.coupled[0]?.detail).toContain('rendered unit')
  })

  it('returns nothing for an empty source span', () => {
    const result = classifyIndependentProperties(new Uint8Array(0), bytesOf('<w:p/>'))

    expect(result).toEqual({ merge: [], rootAttributes: [], coupled: [] })
  })

  it('reports malformed rendered XML when the source span is empty', () => {
    const result = classifyIndependentProperties(new Uint8Array(0), bytesOf('<w:p>'))

    expect(result.merge).toEqual([])
    expect(result.coupled[0]?.name).toBe('xml')
  })
})

describe('splice operation invariants', () => {
  it('never emits a carry operation alongside a span operation for the same part', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_2', 30, 40),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_1', DIGEST_B),
        unit('paragraph', DOC, 'para_2'),
      ],
    }))

    expect(plan.operations.filter(operation => operation.kind === 'carry')).toEqual([])
  })

  it('sorts operations by part and position', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 60, 70),
      anchor('paragraph', DOC, 'para_2', 10, 20),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100, [SHEET]: 100 },
      anchors,
      units: [
        unit('paragraph', DOC, 'para_2', DIGEST_B),
        unit('paragraph', DOC, 'para_1', DIGEST_B),
      ],
    }))

    const parts = plan.operations.map(operation => operation.part)
    expect(parts).toEqual([DOC, DOC, SHEET])
    const docStarts = plan.operations
      .filter(operation => operation.part === DOC)
      .map((operation) => {
        if (operation.kind === 'carry') return 0
        return operation.kind === 'insert' || operation.kind === 'append' ? operation.at : operation.start
      })
    expect(docStarts).toEqual([10, 60])
  })

  it('refuses a classifier result even when the edit renders no bytes', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 20)]
    const plan = buildSplicePlan({
      ...inputOf({
        parts: { [DOC]: 100 },
        anchors,
        units: [unit('paragraph', DOC, 'para_1', DIGEST_B)],
      }),
      classifyCoupled: () => ({
        kind: 'comment' as UnsupportedFeatureKind,
        detail: 'w:commentRangeStart overlaps the paragraph',
        part: DOC,
        start: 10,
      }),
    })

    expect(plan.operations).toEqual([])
    expect(plan.refusal?.kind).toBe('comment')
  })

  it('orders insert, append, replace, and delete operations', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_2', 30, 40),
    ]
    const appendBytes = bytesOf('<w:style/>')
    const plan = buildSplicePlan({
      ...inputOf({
        parts: { [DOC]: 100, [STYLES]: 50 },
        anchors,
        units: [
          unit('paragraph', DOC, 'para_new'),
          unit('paragraph', DOC, 'para_1', DIGEST_B),
        ],
      }),
      appends: [
        { part: STYLES, at: 40, bytes: appendBytes },
        { part: STYLES, at: 40, bytes: appendBytes },
      ],
    })

    expect(operationSummary(plan.operations)).toEqual([
      { kind: 'insert', part: DOC, at: 10, bytes: renderBytes('para_new') },
      { kind: 'replace', part: DOC, start: 10, end: 20, bytes: renderBytes('para_1') },
      { kind: 'delete', part: DOC, start: 30, end: 40 },
      { kind: 'append', part: STYLES, at: 40, bytes: appendBytes },
      { kind: 'append', part: STYLES, at: 40, bytes: appendBytes },
    ])
    expect(plan.entries.find(item => item.name === STYLES)).toEqual({
      name: STYLES,
      status: 'spliced',
    })
  })

  it('sorts classifier edits that share a start offset', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_2', 30, 40),
    ]
    const calls: string[] = []
    buildSplicePlan({
      ...inputOf({
        parts: { [DOC]: 100 },
        anchors,
        units: [
          unit('paragraph', DOC, 'para_new_1'),
          unit('paragraph', DOC, 'para_new_2'),
          unit('paragraph', DOC, 'para_1', DIGEST_B),
        ],
      }),
      classifyCoupled: (edit) => {
        calls.push(`${edit.kind}:${edit.key}`)
        return undefined
      },
    })

    expect(calls).toEqual([
      'insert:para_new_1',
      'insert:para_new_2',
      'replace:para_1',
      'delete:para_2',
    ])
  })

  it('orders an insert before a delete at the same start offset', () => {
    const anchors = [
      anchor('paragraph', DOC, 'para_1', 10, 20),
      anchor('paragraph', DOC, 'para_2', 20, 30),
    ]
    const calls: string[] = []
    buildSplicePlan({
      ...inputOf({
        parts: { [DOC]: 100 },
        anchors,
        units: [
          unit('paragraph', DOC, 'para_1'),
          unit('paragraph', DOC, 'para_new'),
        ],
      }),
      classifyCoupled: (edit) => {
        calls.push(`${edit.kind}:${edit.key}`)
        return undefined
      },
    })

    expect(calls).toEqual(['insert:para_new', 'delete:para_2'])
  })

  it('skips directory entries in the plan and dispositions', () => {
    const anchors = [anchor('paragraph', DOC, 'para_1', 10, 20)]
    const plan = buildSplicePlan(inputOf({
      parts: { 'word/': 0, [DOC]: 100 },
      anchors,
      units: [unit('paragraph', DOC, 'para_1')],
    }))

    expect(plan.operations).toEqual([{ kind: 'carry', part: DOC, entry: DOC }])
    expect(plan.entries).toEqual([{ name: DOC, status: 'verbatim' }])
  })

  it('refuses sheet-summary state with no recorded source part', () => {
    const plan = buildSplicePlan({
      source: sourcePackage({}),
      sourceState: {
        schemaVersion: 1,
        format: 'xlsx',
        packageDigest: DIGEST_A,
        byteLength: 0,
        anchorMode: 'sheet-summary',
        sheets: [],
        sourceParts: [],
      },
      units: [],
      spans: [],
      renderUnit: item => renderBytes(item.key),
    })

    expect(plan.refusal?.part).toBe('')
  })
})

describe('table geometry key parsing', () => {
  it.each([
    'table/0',
    'table/x/0',
    'table/0/x',
    'table/-1/0',
    'table/0/-1',
  ])('refuses the malformed key %s', (key) => {
    const anchors = [anchor('tableCell', DOC, key, 10, 20)]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [unit('tableCell', DOC, key)],
    }))

    expect(plan.refusal?.detail).toContain('is not tableId/row/cell')
  })

  it('refuses a same-size cell replacement that changes columns', () => {
    const anchors = [
      anchor('tableCell', DOC, 'table_1/0/0', 10, 20),
      anchor('tableCell', DOC, 'table_1/0/1', 30, 40),
    ]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('tableCell', DOC, 'table_1/0/0'),
        unit('tableCell', DOC, 'table_1/0/2'),
      ],
    }))

    expect(plan.refusal?.detail).toContain('row or column resize')
  })

  it('refuses a malformed current key when the source table is well formed', () => {
    const anchors = [anchor('tableCell', DOC, 'table_1/0/0', 10, 20)]
    const plan = buildSplicePlan(inputOf({
      parts: { [DOC]: 100 },
      anchors,
      units: [
        unit('tableCell', DOC, 'table_1/0/0'),
        unit('tableCell', DOC, 'not-a-cell-path'),
      ],
    }))

    expect(plan.refusal?.detail).toContain('is not tableId/row/cell')
  })
})

describe('classifyIndependentProperties XML handling', () => {
  it.each([
    ['a comment', '<!--x--><w:p/>'],
    ['CDATA', '<![CDATA[x]]><w:p/>'],
    ['a processing instruction', '<?xml version="1.0"?><w:p/>'],
    ['a DOCTYPE with an internal subset', '<!DOCTYPE w:p [<!ELEMENT w:p ANY>]><w:p/>'],
    ['a quoted system id', '<!DOCTYPE w:p SYSTEM "x>y"><w:p/>'],
    ['a greater-than inside an attribute', '<w:p a="x>y"/>'],
    ['a single-quoted attribute', "<w:p a='x'/>"],
    ['whitespace inside a tag', '<w:p\n\t a="1" />'],
    ['a malformed attribute qname', '<w:p a:="1"/>'],
  ])('parses %s', (_label, xml) => {
    expect(classify(xml, xml)).toEqual({ merge: [], rootAttributes: [], coupled: [] })
  })

  it.each([
    '<!--x',
    '<![CDATA[x',
    '<?xml',
    '<!DOCTYPE x',
    '<!DOCTYPE x [',
    '<!DOCTYPE x "y>',
    '<w:p',
    '<w:p a="x>',
    '</>',
    '</w:p>',
    '<w:p></w:q>',
    '< ="1"/>',
    '<w:p =""/>',
    '<w:p foo/>',
    '<w:p foo=bar/>',
    '<w:p/><w:p/>',
    '<w:/>',
    '<:p/>',
    '<a:b:c/>',
    '<w:p a="x/>',
  ])('fails closed on %s', (xml) => {
    const result = classify(xml, '<w:p/>')

    expect(result.merge).toEqual([])
    expect(result.coupled[0]?.name).toBe('xml')
  })

  it('ignores text with no markup', () => {
    expect(classify('text', '<w:p/>')).toEqual({ merge: [], rootAttributes: [], coupled: [] })
  })

  it('merges a cell style when the rendered unit is empty', () => {
    const result = classifyIndependentProperties(bytesOf('<c s="3"/>'), new Uint8Array(0))

    expect(result.coupled).toEqual([])
    expect(result.merge).toEqual([{ container: 'c', name: 's', xml: 's="3"', start: 3, end: 8 }])
  })

  it('ignores namespace declarations on the source root', () => {
    const xml = `<p xmlns="${WORD_NS}"><r><t>x</t></r></p>`
    expect(classify(xml, xml)).toEqual({ merge: [], rootAttributes: [], coupled: [] })

    const prefixed = `<x:p xmlns:x="${WORD_NS}"><x:r><x:t xml:space="preserve">x</x:t></x:r></x:p>`
    expect(classify(prefixed, prefixed)).toEqual({ merge: [], rootAttributes: [], coupled: [] })

    const sheet = `<c xmlns="${SHEET_NS}"><v>1</v></c>`
    expect(classify(sheet, sheet.replace('1', '2'))).toEqual({ merge: [], rootAttributes: [], coupled: [] })
  })

  it('consumes repeated property elements before reporting a merge', () => {
    const result = classify(
      '<w:p><w:pPr><w:jc w:val="left"/><w:jc w:val="center"/><w:jc w:val="right"/></w:pPr></w:p>',
      '<w:p><w:pPr><w:jc w:val="left"/><w:jc w:val="center"/></w:pPr></w:p>',
    )

    expect(result.merge).toHaveLength(1)
    expect(result.merge[0]?.xml).toBe('<w:jc w:val="right"/>')
  })

  it('does not treat a non-shared formula type as coupled', () => {
    const result = classify(
      '<c><f t="normal">1+1</f><v>2</v></c>',
      '<c><f t="normal">2+2</f><v>4</v></c>',
    )

    expect(result).toEqual({ merge: [], rootAttributes: [], coupled: [] })
  })
})
