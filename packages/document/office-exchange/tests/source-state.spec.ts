/** Sidecar validation: every structural, span, identity, digest, and budget rule. */
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import type { IDocumentData, IWorkbookData } from '@univerjs/core'
import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { resolveConfig } from '../src/config.ts'
import { open } from '../src/package/ooxml-package.ts'
import { readDocumentWithSource, writeDocumentFromSource } from '../src/provider/docx.ts'
import { readWorkbookWithSource } from '../src/provider/xlsx.ts'
import type { OfficeSourceState, SourceStateContext, SourceStateErrorCode } from '../src/source/source-state.ts'
import {
  DEFAULT_MAX_SOURCE_ANCHORS,
  DEFAULT_MAX_SOURCE_ANCHOR_BYTES,
  SourceStateError,
  parseSourceState,
} from '../src/source/source-state.ts'
import { unzipEntries } from './ooxml-package.ts'

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url))
const DOCX_FIXTURE = `${FIXTURES}/roundtrip.docx`
const XLSX_FIXTURE = `${FIXTURES}/roundtrip.xlsx`
const ONE_SHEET_FIXTURE = `${FIXTURES}/preserve-excel-table.xlsx`

type SidecarCarrier = {
  officeExchange?: {
    readonly source?: unknown
  }
}

function sourceStateOf(data: IDocumentData | IWorkbookData): OfficeSourceState | undefined {
  return (data as SidecarCarrier).officeExchange?.source as OfficeSourceState | undefined
}

async function contextOf(path: string): Promise<SourceStateContext> {
  const source = open(new Uint8Array(await readFile(path)), { maxEntries: 10_000 }).source
  return {
    partByteLengths: new Map(
      source.entries.filter(entry => !entry.name.endsWith('/')).map(entry => [entry.name, entry.uncompressedSize]),
    ),
  }
}

const DOCX_PART = 'word/document.xml'
const SHEET_PART = 'xl/worksheets/sheet1.xml'
const DIGEST = 'a'.repeat(64)
const PART_BYTE_LENGTHS: ReadonlyMap<string, number> = new Map([
  [DOCX_PART, 200],
  [SHEET_PART, 120],
])
const CONTEXT: SourceStateContext = { partByteLengths: PART_BYTE_LENGTHS }

function anchor(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    kind: 'paragraph',
    part: DOCX_PART,
    key: 'para_1',
    start: 0,
    end: 10,
    digest: DIGEST,
    ...overrides,
  }
}

function perUnitState(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    schemaVersion: 1,
    format: 'docx',
    packageDigest: DIGEST,
    byteLength: 500,
    anchorMode: 'per-unit',
    anchors: [anchor()],
    sourceParts: [DOCX_PART],
    ...overrides,
  }
}

function sheetSummary(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    sheetId: 'sheet-1',
    part: SHEET_PART,
    digest: DIGEST,
    cellCount: 12,
    ...overrides,
  }
}

function sheetState(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    schemaVersion: 1,
    format: 'xlsx',
    packageDigest: DIGEST,
    byteLength: 500,
    anchorMode: 'sheet-summary',
    sheets: [sheetSummary()],
    sourceParts: [SHEET_PART],
    ...overrides,
  }
}

function expectError(value: unknown, code: SourceStateErrorCode, context: SourceStateContext = CONTEXT): SourceStateError {
  try {
    parseSourceState(value, context)
  } catch (error) {
    if (error instanceof SourceStateError) {
      expect(error.code).toBe(code)
      expect(error.message.length).toBeGreaterThan(0)
      return error
    }
    throw error
  }
  throw new Error('expected parseSourceState to reject the value')
}

describe('parseSourceState acceptance', () => {
  it.each(['docx', 'docm', 'xlsx', 'xlsm'] as const)('accepts the %s format', (format) => {
    const state = perUnitState({ format })
    expect(parseSourceState(state, CONTEXT)).toEqual(state)
  })

  it('accepts every anchor kind, multiple parts, and adjacent or unsorted spans', () => {
    const anchors = [
      anchor({ kind: 'tableCell', key: 'table-1/0/0', start: 10, end: 20 }),
      anchor({ kind: 'paragraph', key: 'para_1', start: 0, end: 10 }),
      anchor({ kind: 'sheetCell', part: SHEET_PART, key: 'sheet-1!0:0', start: 0, end: 10 }),
    ]
    const state = perUnitState({ anchors, sourceParts: [DOCX_PART, SHEET_PART] })
    expect(parseSourceState(state, CONTEXT)).toEqual(state)
  })

  it('accepts a span that ends exactly at the part byte length', () => {
    const state = perUnitState({ anchors: [anchor({ start: 190, end: 200 })] })
    expect(parseSourceState(state, CONTEXT)).toEqual(state)
  })

  it('accepts an empty anchor list and empty source part list', () => {
    const state = perUnitState({ anchors: [], sourceParts: [] })
    expect(parseSourceState(state, CONTEXT)).toEqual(state)
  })

  it('accepts a sheet-summary sidecar', () => {
    const state = sheetState()
    expect(parseSourceState(state, CONTEXT)).toEqual(state)
  })

  it('accepts zero cells and an empty sheet list', () => {
    expect(parseSourceState(sheetState({ sheets: [sheetSummary({ cellCount: 0 })] }), CONTEXT)).toEqual(
      sheetState({ sheets: [sheetSummary({ cellCount: 0 })] }),
    )
    expect(parseSourceState(sheetState({ sheets: [] }), CONTEXT)).toEqual(sheetState({ sheets: [] }))
  })

  it('rebuilds the state without unknown fields', () => {
    const parsed = parseSourceState(perUnitState({ futureField: { nested: true } }), CONTEXT)
    expect(parsed).toEqual(perUnitState())
    expect(parsed).not.toHaveProperty('futureField')
  })

  it('accepts the default budgets when the caller does not supply them', () => {
    expect(parseSourceState(perUnitState(), { partByteLengths: PART_BYTE_LENGTHS })).toEqual(perUnitState())
    expect(DEFAULT_MAX_SOURCE_ANCHORS).toBe(100_000)
    expect(DEFAULT_MAX_SOURCE_ANCHOR_BYTES).toBe(8 * 1024 * 1024)
  })
})

describe('parseSourceState header validation', () => {
  it('rejects a value that is not an object', () => {
    expectError(null, 'invalid-state')
    expectError([], 'invalid-state')
    expectError('state', 'invalid-state')
  })

  it('rejects an unsupported schema version', () => {
    expect(expectError(perUnitState({ schemaVersion: 2 }), 'invalid-state').message).toContain('schemaVersion')
    expectError(perUnitState({ schemaVersion: '1' }), 'invalid-state')
  })

  it('rejects an unsupported format', () => {
    expectError(perUnitState({ format: 'pptx' }), 'invalid-state')
    expectError(perUnitState({ format: 42 }), 'invalid-state')
  })

  it('rejects a malformed package digest', () => {
    expectError(perUnitState({ packageDigest: 'A'.repeat(64) }), 'invalid-digest')
    expectError(perUnitState({ packageDigest: 'a'.repeat(63) }), 'invalid-digest')
    expectError(perUnitState({ packageDigest: 42 }), 'invalid-state')
  })

  it('rejects a malformed byte length', () => {
    expectError(perUnitState({ byteLength: -1 }), 'invalid-state')
    expectError(perUnitState({ byteLength: 1.5 }), 'invalid-state')
    expectError(perUnitState({ byteLength: '500' }), 'invalid-state')
  })

  it('rejects an unsupported anchor mode', () => {
    expectError(perUnitState({ anchorMode: 'unit' }), 'invalid-state')
    expectError(perUnitState({ anchorMode: 1 }), 'invalid-state')
  })

  it('rejects malformed source part lists', () => {
    expectError(perUnitState({ sourceParts: DOCX_PART }), 'invalid-state')
    expectError(perUnitState({ sourceParts: [1] }), 'invalid-state')
    expectError(perUnitState({ sourceParts: [''] }), 'invalid-state')
    expectError(perUnitState({ sourceParts: ['unknown.xml'] }), 'unknown-part')
    expectError(perUnitState({ sourceParts: [DOCX_PART, DOCX_PART] }), 'invalid-state')
  })
})

describe('parseSourceState anchor representation', () => {
  it('requires anchors in per-unit mode', () => {
    expectError(perUnitState({ anchors: undefined }), 'invalid-state')
    expectError(perUnitState({ anchors: 'none' }), 'invalid-state')
  })

  it('rejects sheets in per-unit mode', () => {
    expectError(perUnitState({ sheets: [] }), 'invalid-state')
  })

  it('requires sheets in sheet-summary mode', () => {
    expectError(sheetState({ sheets: undefined }), 'invalid-state')
    expectError(sheetState({ sheets: {} }), 'invalid-state')
  })

  it('rejects anchors in sheet-summary mode', () => {
    expectError(sheetState({ anchors: [] }), 'invalid-state')
  })

  it('rejects a non-object anchor entry', () => {
    expectError(perUnitState({ anchors: ['anchor'] }), 'invalid-state')
    expectError(perUnitState({ anchors: [null] }), 'invalid-state')
    expectError(perUnitState({ anchors: [[]] }), 'invalid-state')
  })

  it('rejects an unsupported anchor kind', () => {
    expectError(perUnitState({ anchors: [anchor({ kind: 'row' })] }), 'invalid-state')
  })

  it('rejects a missing, unknown, or non-string part', () => {
    expectError(perUnitState({ anchors: [anchor({ part: 42 })] }), 'invalid-state')
    expectError(perUnitState({ anchors: [anchor({ part: 'missing.xml' })] }), 'unknown-part')
  })

  it('rejects a missing, empty, or non-string key', () => {
    expectError(perUnitState({ anchors: [anchor({ key: 42 })] }), 'invalid-state')
    expectError(perUnitState({ anchors: [anchor({ key: '' })] }), 'invalid-state')
  })

  it('rejects non-integer and out-of-range spans', () => {
    expectError(perUnitState({ anchors: [anchor({ start: '0' })] }), 'invalid-state')
    expectError(perUnitState({ anchors: [anchor({ end: 1.5 })] }), 'invalid-state')
    expectError(perUnitState({ anchors: [anchor({ start: -1 })] }), 'invalid-span')
    expectError(perUnitState({ anchors: [anchor({ start: 10, end: 10 })] }), 'invalid-span')
    expectError(perUnitState({ anchors: [anchor({ start: 20, end: 10 })] }), 'invalid-span')
    expectError(perUnitState({ anchors: [anchor({ start: 190, end: 201 })] }), 'invalid-span')
  })

  it('rejects a malformed anchor digest', () => {
    expectError(perUnitState({ anchors: [anchor({ digest: 'A'.repeat(64) })] }), 'invalid-digest')
    expectError(perUnitState({ anchors: [anchor({ digest: 'a'.repeat(65) })] }), 'invalid-digest')
    expectError(perUnitState({ anchors: [anchor({ digest: 1 })] }), 'invalid-state')
  })

  it('rejects a duplicate (part, kind, key) identity', () => {
    expectError(perUnitState({
      anchors: [anchor({ start: 0, end: 10 }), anchor({ start: 20, end: 30 })],
    }), 'duplicate-key')
  })

  it('allows the same key for another kind or another part', () => {
    const otherKind = perUnitState({
      anchors: [
        anchor({ kind: 'paragraph', key: 'shared', start: 0, end: 10 }),
        anchor({ kind: 'tableCell', key: 'shared', start: 20, end: 30 }),
      ],
    })
    expect(parseSourceState(otherKind, CONTEXT)).toEqual(otherKind)
    const otherPart = perUnitState({
      anchors: [
        anchor({ kind: 'paragraph', part: DOCX_PART, key: 'shared', start: 0, end: 10 }),
        anchor({ kind: 'paragraph', part: SHEET_PART, key: 'shared', start: 0, end: 10 }),
      ],
      sourceParts: [DOCX_PART, SHEET_PART],
    })
    expect(parseSourceState(otherPart, CONTEXT)).toEqual(otherPart)
  })

  it('rejects overlapping spans in one part and allows overlap across parts', () => {
    expectError(perUnitState({
      anchors: [
        anchor({ key: 'one', start: 0, end: 20 }),
        anchor({ key: 'two', start: 10, end: 30 }),
      ],
    }), 'overlapping-spans')
    expectError(perUnitState({
      anchors: [
        anchor({ key: 'one', start: 10, end: 30 }),
        anchor({ key: 'two', start: 10, end: 20 }),
      ],
    }), 'overlapping-spans')
    const crossPart = perUnitState({
      anchors: [
        anchor({ key: 'one', part: DOCX_PART, start: 0, end: 20 }),
        anchor({ key: 'two', part: SHEET_PART, start: 10, end: 30 }),
      ],
      sourceParts: [DOCX_PART, SHEET_PART],
    })
    expect(parseSourceState(crossPart, CONTEXT)).toEqual(crossPart)
  })
})

describe('parseSourceState budgets', () => {
  it('enforces the anchor count budget at the exact boundary', () => {
    const anchors = [anchor({ key: 'one', start: 0, end: 10 }), anchor({ key: 'two', start: 20, end: 30 })]
    const state = perUnitState({ anchors })
    const exact = { partByteLengths: PART_BYTE_LENGTHS, maxSourceAnchors: 2 }
    expect(parseSourceState(state, exact)).toEqual(state)
    expectError(state, 'anchor-budget-exceeded', { partByteLengths: PART_BYTE_LENGTHS, maxSourceAnchors: 1 })
  })

  it('enforces the anchor payload byte budget at the exact boundary', () => {
    const anchors = [anchor()]
    const state = perUnitState({ anchors })
    const payloadBytes = Buffer.byteLength(JSON.stringify(anchors), 'utf8')
    const exact = { partByteLengths: PART_BYTE_LENGTHS, maxSourceAnchorBytes: payloadBytes }
    expect(parseSourceState(state, exact)).toEqual(state)
    expectError(state, 'anchor-budget-exceeded', { partByteLengths: PART_BYTE_LENGTHS, maxSourceAnchorBytes: payloadBytes - 1 })
  })

  it('applies the count budget to sheet summaries too', () => {
    const state = sheetState({
      sheets: [sheetSummary({ sheetId: 'one' }), sheetSummary({ sheetId: 'two' })],
    })
    expectError(state, 'anchor-budget-exceeded', { partByteLengths: PART_BYTE_LENGTHS, maxSourceAnchors: 1 })
  })

  it('rejects invalid caller budgets', () => {
    for (const value of [0, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(() => parseSourceState(perUnitState(), {
        partByteLengths: PART_BYTE_LENGTHS,
        maxSourceAnchors: value,
      })).toThrow(RangeError)
      expect(() => parseSourceState(perUnitState(), {
        partByteLengths: PART_BYTE_LENGTHS,
        maxSourceAnchorBytes: value,
      })).toThrow(RangeError)
    }
  })
})

describe('parseSourceState sheet summaries', () => {
  it('rejects a non-object sheet entry', () => {
    expectError(sheetState({ sheets: ['sheet'] }), 'invalid-state')
  })

  it('rejects a missing, empty, or non-string sheet id', () => {
    expectError(sheetState({ sheets: [sheetSummary({ sheetId: 42 })] }), 'invalid-state')
    expectError(sheetState({ sheets: [sheetSummary({ sheetId: '' })] }), 'invalid-state')
  })

  it('rejects a duplicate sheet id', () => {
    expectError(sheetState({
      sheets: [sheetSummary({ sheetId: 'same' }), sheetSummary({ sheetId: 'same' })],
    }), 'invalid-state')
  })

  it('rejects a missing, unknown, or non-string sheet part', () => {
    expectError(sheetState({ sheets: [sheetSummary({ part: 42 })] }), 'invalid-state')
    expectError(sheetState({ sheets: [sheetSummary({ part: 'missing.xml' })] }), 'unknown-part')
  })

  it('rejects a malformed sheet digest', () => {
    expectError(sheetState({ sheets: [sheetSummary({ digest: 'A'.repeat(64) })] }), 'invalid-digest')
  })

  it('rejects a non-integer, negative, or non-number cell count', () => {
    expectError(sheetState({ sheets: [sheetSummary({ cellCount: -1 })] }), 'invalid-state')
    expectError(sheetState({ sheets: [sheetSummary({ cellCount: 1.5 })] }), 'invalid-state')
    expectError(sheetState({ sheets: [sheetSummary({ cellCount: '12' })] }), 'invalid-state')
  })
})

describe('source attachment', () => {
  it('validates a document fixture sidecar against its source package', async () => {
    const bytes = new Uint8Array(await readFile(DOCX_FIXTURE))
    const data = await readDocumentWithSource(bytes, 'docx')
    const state = sourceStateOf(data)
    expect(state).toBeDefined()
    if (state === undefined) return
    expect(parseSourceState(state, await contextOf(DOCX_FIXTURE))).toEqual(state)
    expect(state.anchorMode).toBe('per-unit')
    if (state.anchorMode !== 'per-unit') return
    expect(state.anchors.filter(anchor => anchor.kind === 'paragraph')).toHaveLength(2)
    expect(state.anchors.filter(anchor => anchor.kind === 'tableCell')).toHaveLength(4)
    expect((data.body?.paragraphs ?? []).every(paragraph => paragraph.paragraphId?.startsWith('para_'))).toBe(true)
    expect(JSON.parse(JSON.stringify(state))).toEqual(state)
  })

  it('recovers the source sidecar from a persisted draft and saves the draft', async () => {
    const bytes = new Uint8Array(await readFile(DOCX_FIXTURE))
    const data = await readDocumentWithSource(bytes, 'docx')
    const state = sourceStateOf(data)
    expect(state).toBeDefined()
    const body = data.body
    if (body === undefined) throw new Error('fixture document has no body')
    expect(body.dataStream).toContain('Quarterly')
    body.dataStream = body.dataStream.replace('Quarterly', 'Financial')

    // The Remote persists the wire draft as JSON; a restart reads it back whole.
    const draft = JSON.stringify({
      documentId: 'draft-document',
      path: 'roundtrip.docx',
      contentVersion: 'version-1',
      snapshot: { kind: 'document', data },
      savedAt: new Date(0).toISOString(),
    })
    const recovered = JSON.parse(draft) as {
      readonly snapshot: { readonly kind: 'document'; readonly data: IDocumentData }
    }
    expect(sourceStateOf(recovered.snapshot.data)).toEqual(state)

    const written = await writeDocumentFromSource(recovered.snapshot.data, bytes, resolveConfig({}))
    const sourceEntries = unzipEntries(bytes)
    const writtenEntries = unzipEntries(written.bytes)
    expect(writtenEntries.get('word/styles.xml')).toEqual(sourceEntries.get('word/styles.xml'))
    expect(written.disposition.spliced.names).toEqual(['word/document.xml'])

    const reread = await readDocumentWithSource(written.bytes, 'docx')
    expect(reread.body?.dataStream).toContain('Financial')
    expect(reread.body?.dataStream).not.toContain('Quarterly')
  })

  it('refuses a stale sidecar whose source cannot be re-anchored', async () => {
    const bytes = new Uint8Array(await readFile(DOCX_FIXTURE))
    const data = await readDocumentWithSource(bytes, 'docx')
    const body = data.body
    if (body === undefined) throw new Error('fixture document has no body')
    body.dataStream = body.dataStream.replace('Quarterly', 'Financial')
    const unrelated = zipSync({
      '[Content_Types].xml': strToU8('<Types/>'),
      'notes/readme.txt': strToU8('not a document'),
    })

    await expect(writeDocumentFromSource(data, unrelated, resolveConfig({})))
      .rejects.toMatchObject({ code: 'conversion-failed' })
  })

  it('validates a workbook fixture sidecar with one anchor per cell', async () => {
    const bytes = new Uint8Array(await readFile(XLSX_FIXTURE))
    const data = await readWorkbookWithSource(bytes, 'xlsx')
    const state = sourceStateOf(data)
    expect(state).toBeDefined()
    if (state === undefined) return
    expect(parseSourceState(state, await contextOf(XLSX_FIXTURE))).toEqual(state)
    expect(state.anchorMode).toBe('per-unit')
    if (state.anchorMode !== 'per-unit') return
    expect(state.anchors.every(anchor => anchor.kind === 'sheetCell')).toBe(true)
    expect(state.anchors.length).toBeGreaterThan(0)
    expect(new Set(state.anchors.map(anchor => anchor.key)).size).toBe(state.anchors.length)
    expect(JSON.parse(JSON.stringify(state))).toEqual(state)
  })

  it('validates the sheet-summary fallback when the anchor budget cannot hold cells', async () => {
    const bytes = new Uint8Array(await readFile(ONE_SHEET_FIXTURE))
    const data = await readWorkbookWithSource(bytes, 'xlsx', resolveConfig({ maxSourceAnchors: 1 }))
    const state = sourceStateOf(data)
    expect(state?.anchorMode).toBe('sheet-summary')
    if (state === undefined) return
    expect(parseSourceState(state, await contextOf(ONE_SHEET_FIXTURE))).toEqual(state)
    if (state.anchorMode !== 'sheet-summary') return
    expect(state.sheets).toHaveLength(1)
    expect(state.sheets[0]?.cellCount).toBeGreaterThan(0)
    expect(JSON.parse(JSON.stringify(state))).toEqual(state)
  })
})

describe('SourceStateError', () => {
  it('reports its stable code and type name', () => {
    const error = expectError(perUnitState({ anchors: [anchor({ start: 10, end: 10 })] }), 'invalid-span')
    expect(error.name).toBe('SourceStateError')
    expect(error).toBeInstanceOf(Error)
  })
})
