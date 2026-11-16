/** Document privacy egress: reading, editing, and saving Office files must reach no network destination (SC-008). */
import http from 'node:http'
import https from 'node:https'
import type { OfficeDocumentSnapshot } from '@forwork/office-exchange'
import type { JsonValue } from '@forwork/util-values'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { assertNoEgress, recordEgress } from '../src/index.ts'
import { createHarness, disposeHarness, documentIdFor } from '../../../document/tool-office/tests/harness.ts'
import type { ToolTestHarness } from '../../../document/tool-office/tests/harness.ts'

const SIGNAL = new AbortController().signal
/** Destination of the control call that proves the recorder is active. */
const CONTROL_ENDPOINT = 'https://models.invalid/v1/chat'
const originalFetch = globalThis.fetch
const originalHttpRequest = http.request
const originalHttpsRequest = https.request

let h: ToolTestHarness

beforeEach(async () => {
  globalThis.fetch = async () => new Response('ok')
  h = await createHarness()
})

afterEach(async () => {
  await disposeHarness(h)
  globalThis.fetch = originalFetch
  http.request = originalHttpRequest
  https.request = originalHttpsRequest
})

/**
 * Run one document flow under a live egress recorder.
 * The control call confirms the recorder observed this test body, so the empty
 * destination list that follows is the flow's own behavior.
 * @param flow - document tool work to observe.
 * @returns the flow's result.
 */
async function recordDocumentFlow<T>(flow: () => Promise<T>): Promise<T> {
  const control = await recordEgress(async () => { await fetch(CONTROL_ENDPOINT) })
  expect(control.urls).toEqual([CONTROL_ENDPOINT])
  const capture = await recordEgress(flow)
  expect(capture.urls).toEqual([])
  assertNoEgress(capture.urls)
  return capture.result
}

/** Execute one tool call through the real pipeline and require a structured result. */
async function call(name: string, arguments_: Record<string, unknown>): Promise<JsonValue> {
  const result = await h.execute(name, arguments_, SIGNAL)
  expect(result.isError).toBe(false)
  return result.value ?? null
}

/** A JSON object view of one snapshot value. */
function objectOf(value: JsonValue | undefined): Record<string, JsonValue> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) throw new Error('expected a JSON object')
  return value
}

/** The sheets record of a workbook snapshot. */
function sheetsOf(snapshot: OfficeDocumentSnapshot): Record<string, JsonValue> {
  if (snapshot.kind !== 'workbook') throw new Error('expected a workbook')
  return objectOf(objectOf(snapshot.data)['sheets'])
}

/** The first sheet id of a workbook snapshot. */
function firstSheet(snapshot: OfficeDocumentSnapshot): string {
  if (snapshot.kind !== 'workbook') throw new Error('expected a workbook')
  const order = objectOf(snapshot.data)['sheetOrder']
  const id = Array.isArray(order) ? order[0] : undefined
  if (typeof id !== 'string') throw new Error('expected a sheet id')
  return id
}

/** Copy a read snapshot and set one cell value. */
function withCell(
  snapshot: OfficeDocumentSnapshot,
  sheetId: string,
  row: number,
  column: number,
  value: JsonValue,
): OfficeDocumentSnapshot {
  const copy = structuredClone(snapshot)
  const sheets = sheetsOf(copy)
  const sheet = objectOf(sheets[sheetId])
  const cells = objectOf(sheet['cellData'])
  const rowCells = cells[String(row)] === undefined ? {} : objectOf(cells[String(row)])
  rowCells[String(column)] = { v: value }
  cells[String(row)] = rowCells
  sheets[sheetId] = sheet
  return copy
}

/** Read one workbook cell value. */
function cellValue(snapshot: OfficeDocumentSnapshot, sheetId: string, row: number, column: number): JsonValue {
  const sheet = objectOf(sheetsOf(snapshot)[sheetId])
  const cells = objectOf(sheet['cellData'])
  const rowCells = cells[String(row)] === undefined ? {} : objectOf(cells[String(row)])
  return objectOf(rowCells[String(column)])['v'] ?? null
}

/** The `document/edit` facts the calling Session recorded. */
function editFacts(): number {
  return h.events.filter(event => event.type === 'document/edit').length
}

describe('document egress', () => {
  it('reads a workbook without reaching a network destination', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await h.server.read(h.scope, path, SIGNAL)
    vi.spyOn(h.server, 'dispatchCommand').mockResolvedValue({
      status: 'read',
      snapshot: read.snapshot,
      contentVersion: 4,
    })

    const value = await recordDocumentFlow(async () => await call('document_read', {
      documentId: documentIdFor(h, path),
    }))

    const result = objectOf(value)
    expect(result['status']).toBe('ok')
    expect(result['contentVersion']).toBe(4)
    expect(objectOf(result['summary'])['kind']).toBe('workbook')
    expect(result).not.toHaveProperty('snapshot')
  })

  it('applies and saves a cell edit without reaching a network destination', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await h.server.read(h.scope, path, SIGNAL)
    const sheetId = firstSheet(read.snapshot)
    vi.spyOn(h.server, 'dispatchCommand').mockResolvedValue({
      status: 'applied',
      snapshot: withCell(read.snapshot, sheetId, 1, 0, 'Sprocket'),
      contentVersion: 1,
    })

    const value = await recordDocumentFlow(async () => await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ sheetId, row: 1, column: 0, value: 'Sprocket' }],
    }))

    const result = objectOf(value)
    expect(result['status']).toBe('ok')
    expect(result['contentVersion']).toBe(1)
    expect(editFacts()).toBe(1)
    const saved = await h.server.read(h.scope, path, SIGNAL)
    expect(cellValue(saved.snapshot, sheetId, 1, 0)).toBe('Sprocket')
  })

  it('applies and saves a paragraph edit without reaching a network destination', async () => {
    const path = await h.copy('roundtrip.docx')
    const read = await h.server.read(h.scope, path, SIGNAL)
    vi.spyOn(h.server, 'dispatchCommand').mockResolvedValue({
      status: 'applied',
      snapshot: read.snapshot,
      contentVersion: 1,
    })
    const save = vi.spyOn(h.server, 'save')

    const value = await recordDocumentFlow(async () => await call('document_edit_paragraph', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ index: 0, text: 'Updated paragraph' }],
    }))

    const result = objectOf(value)
    expect(result['status']).toBe('ok')
    expect(result['contentVersion']).toBe(1)
    expect(save).toHaveBeenCalledOnce()
    expect(editFacts()).toBe(1)
  })
})
