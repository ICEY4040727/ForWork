/** Office document tools: bridge success paths, refusals, save counts, and edit facts. */
import { createHash } from 'node:crypto'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { Context } from '@forwork/cordis'
import { mountAgentLoopTestDependencies, mountAgentLoopTestHarness } from '@forwork/agent-loop-testkit'
import { LocalFileSystem } from '@forwork/fs-local'
import { ToolCallId } from '@forwork/llm'
import { OfficeDocumentRemote } from '@forwork/office-exchange'
import { open } from '@forwork/office-exchange/src/package/ooxml-package.ts'
import type {
  OfficeCommandResult,
  OfficeDocumentSnapshot,
  OfficeDocumentReadResult,
} from '@forwork/office-exchange'
import { SessionId } from '@forwork/session'
import type { DocumentCommandEventData } from '@forwork/session'
import { executeToolFixture } from '@forwork/tool-testkit'
import { RemoteError } from '@forwork/typert-protocol'
import type { JsonValue } from '@forwork/util-values'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  apply,
  executeDocumentTool,
  registerDocumentTools,
  resolveToolConfig,
} from '../src/index.ts'
import type { DocumentToolRuntime } from '../src/runtime.ts'
import * as ToolOffice from '../src/index.ts'
import { createHarness, disposeHarness, documentIdFor } from './harness.ts'
import type { ToolTestHarness } from './harness.ts'

const SIGNAL = new AbortController().signal
let h: ToolTestHarness

beforeEach(async () => {
  h = await createHarness()
})
afterEach(async () => {
  await disposeHarness(h)
})

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

/** One document body's data stream. */
function streamOf(body: Record<string, JsonValue>): string {
  const stream = body['dataStream']
  if (typeof stream !== 'string') throw new Error('expected a document data stream')
  return stream
}

/** Copy a read snapshot and set one cell value. */
function withCell(
  snapshot: OfficeDocumentSnapshot,
  sheetId: string,
  row: number,
  column: number,
  value: JsonValue,
  formula?: string,
): OfficeDocumentSnapshot {
  const copy = structuredClone(snapshot)
  const sheets = sheetsOf(copy)
  const sheet = objectOf(sheets[sheetId])
  const cells = objectOf(sheet['cellData'])
  const rowCells = cells[String(row)] === undefined ? {} : objectOf(cells[String(row)])
  rowCells[String(column)] = formula === undefined ? { v: value } : { v: value, f: formula }
  cells[String(row)] = rowCells
  sheets[sheetId] = sheet
  return copy
}

/** Read one workbook cell value. */
function cellValue(snapshot: OfficeDocumentSnapshot, sheetId: string, row: number, column: number): JsonValue {
  const sheet = objectOf(sheetsOf(snapshot)[sheetId])
  const cells = objectOf(sheet['cellData'])
  const rowCells = cells[String(row)] === undefined ? {} : objectOf(cells[String(row)])
  const cell = rowCells[String(column)]
  return cell === undefined ? null : objectOf(cell)['v'] ?? null
}

/** Read one workbook cell style. */
function cellStyle(snapshot: OfficeDocumentSnapshot, sheetId: string, row: number, column: number): JsonValue {
  const sheet = objectOf(sheetsOf(snapshot)[sheetId])
  const cells = objectOf(sheet['cellData'])
  const rowCells = cells[String(row)] === undefined ? {} : objectOf(cells[String(row)])
  const cell = rowCells[String(column)]
  return cell === undefined ? null : objectOf(cell)['s'] ?? null
}

/** One sheet's row and column capacity. */
function sheetSize(
  snapshot: OfficeDocumentSnapshot,
  sheetId: string,
): { readonly rows: number; readonly columns: number } {
  const sheet = objectOf(sheetsOf(snapshot)[sheetId])
  const rows = sheet['rowCount']
  const columns = sheet['columnCount']
  if (typeof rows !== 'number' || typeof columns !== 'number') throw new Error('expected sheet dimensions')
  return { rows, columns }
}

/** Copy a snapshot and shift its cells for one row insertion or deletion. */
function withRows(
  snapshot: OfficeDocumentSnapshot,
  sheetId: string,
  row: number,
  count: number,
  action: 'insert' | 'delete',
): OfficeDocumentSnapshot {
  const copy = structuredClone(snapshot)
  const size = sheetSize(copy, sheetId)
  const sheet = objectOf(sheetsOf(copy)[sheetId])
  const cells = objectOf(sheet['cellData'])
  const moved: Record<string, JsonValue> = {}
  for (const [rowKey, rowValue] of Object.entries(cells)) {
    const index = Number(rowKey)
    if (action === 'delete' && index >= row && index < row + count) continue
    const target = action === 'insert'
      ? index >= row ? index + count : index
      : index >= row + count ? index - count : index
    moved[String(target)] = rowValue
  }
  sheet['cellData'] = moved
  sheet['rowCount'] = action === 'insert' ? size.rows + count : size.rows - count
  return copy
}

/** Copy a snapshot and shift its cells for one column insertion or deletion. */
function withColumns(
  snapshot: OfficeDocumentSnapshot,
  sheetId: string,
  column: number,
  count: number,
  action: 'insert' | 'delete',
): OfficeDocumentSnapshot {
  const copy = structuredClone(snapshot)
  const size = sheetSize(copy, sheetId)
  const sheet = objectOf(sheetsOf(copy)[sheetId])
  const cells = objectOf(sheet['cellData'])
  const moved: Record<string, JsonValue> = {}
  for (const [rowKey, rowValue] of Object.entries(cells)) {
    const rowCells = objectOf(rowValue)
    if (rowCells === undefined) continue
    const targetCells: Record<string, JsonValue> = {}
    for (const [columnKey, cellValue_] of Object.entries(rowCells)) {
      const index = Number(columnKey)
      if (action === 'delete' && index >= column && index < column + count) continue
      const target = action === 'insert'
        ? index >= column ? index + count : index
        : index >= column + count ? index - count : index
      targetCells[String(target)] = cellValue_
    }
    moved[rowKey] = targetCells
  }
  sheet['cellData'] = moved
  sheet['columnCount'] = action === 'insert' ? size.columns + count : size.columns - count
  return copy
}

/** Copy a snapshot and assign one style to every cell in a range. */
function withStyle(
  snapshot: OfficeDocumentSnapshot,
  sheetId: string,
  range: { readonly startRow: number; readonly startColumn: number; readonly endRow: number; readonly endColumn: number },
  style: JsonValue,
): OfficeDocumentSnapshot {
  const copy = structuredClone(snapshot)
  const sheet = objectOf(sheetsOf(copy)[sheetId])
  const cells = objectOf(sheet['cellData'])
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const rowCells = cells[String(row)] === undefined ? {} : objectOf(cells[String(row)])
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const cell = rowCells[String(column)] === undefined ? {} : objectOf(rowCells[String(column)])
      cell['s'] = structuredClone(style)
      rowCells[String(column)] = cell
    }
    cells[String(row)] = rowCells
  }
  return copy
}

/** The first sheet id of a workbook. */
function firstSheet(snapshot: OfficeDocumentSnapshot): string {
  if (snapshot.kind !== 'workbook') throw new Error('expected a workbook')
  const order = objectOf(snapshot.data)['sheetOrder']
  const id = Array.isArray(order) ? order[0] : undefined
  if (typeof id !== 'string') throw new Error('expected a sheet id')
  return id
}

/** Read the workbook fixture through the real Remote. */
async function readWorkbook(path: string): Promise<OfficeDocumentReadResult> {
  return await h.server.read(h.scope, path, SIGNAL)
}

/**
 * Copy a workbook fixture into the workspace without any merged range.
 *
 * The round-trip fixture merges `A7:C7`, and a structural row or column shift
 * refuses to move a merged range. Removing the merge keeps the plain-sheet
 * fixtures exercising insertion and deletion on their own.
 * @param fixture - workbook fixture file name copied from the shared fixtures.
 * @returns the workspace-relative path of the merge-free workbook.
 */
async function copyMergeFreeWorkbook(fixture: string): Promise<string> {
  const path = await h.copy(fixture)
  const absolute = join(h.root, path)
  const packageFile = open(new Uint8Array(await readFile(absolute)), { maxEntries: 10_000 })
  const part = 'xl/worksheets/sheet1.xml'
  const source = new TextDecoder().decode(packageFile.readPart(part))
  const stripped = source.replace(/<mergeCells\b[^>]*>[\s\S]*?<\/mergeCells>/u, '')
  if (stripped === source) throw new Error(`${fixture} has no merged range to remove`)
  packageFile.replacePart(part, new TextEncoder().encode(stripped))
  await writeFile(absolute, packageFile.save())
  return path
}

/** Count the `document/edit` facts in the calling Session. */
function editFacts(): number {
  return h.events.filter(event => event.type === 'document/edit').length
}

/** SHA-256 hex digest of one workspace file. */
async function hashOfFile(path: string): Promise<string> {
  return createHash('sha256').update(await readFile(path)).digest('hex')
}

/** The first `document/edit` fact in the calling Session. */
function firstEdit(): Extract<import('@forwork/session').SessionEvent, { type: 'document/edit' }> {
  const edit = h.events.find(event => event.type === 'document/edit')
  if (edit?.type !== 'document/edit') throw new Error('missing edit fact')
  return edit
}

/** Execute a tool call that returns a structured value. */
async function call(name: string, arguments_: Record<string, unknown>, signal: AbortSignal = SIGNAL): Promise<JsonValue> {
  const result = await h.execute(name, arguments_, signal)
  expect(result.isError, result.isError ? result.error.message : undefined).toBe(false)
  return result.value ?? null
}

/**
 * Register one minimal live target for a document and answer its appended
 * commands. It implements the same registration and correlated-response
 * interface the client bridge uses, without a browser or Univer runtime.
 * @param h - mounted composition.
 * @param path - workspace-relative document path.
 * @param respond - result for one command, or `undefined` to stay silent.
 * @returns release of the registration and its event listener.
 */
async function attachLiveTarget(
  h: ToolTestHarness,
  path: string,
  respond: (request: DocumentCommandEventData) => Promise<OfficeCommandResult | undefined>,
): Promise<() => Promise<void>> {
  const documentId = documentIdFor(h, path)
  const registrationId = `live-${documentId}`
  await h.server.registerCommandTarget(h.scope, documentId, registrationId, SIGNAL)
  const stop = h.ctx.on('session/event', (session, event) => {
    if (session !== h.agent.session || event.type !== 'document/command') return
    if (event.data.documentId !== documentId) return
    void respond(event.data).then(async (result) => {
      if (result === undefined) return
      await h.server.completeCommand(h.scope, event.data.requestId, result, SIGNAL)
    })
  })
  return async () => {
    stop()
    await h.server.unregisterCommandTarget(h.scope, documentId, registrationId, SIGNAL)
  }
}

describe('document_read', () => {
  it('reads the live structure and never saves', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const dispatch = vi.spyOn(h.server, 'dispatchCommand').mockResolvedValue({
      status: 'read',
      snapshot: read.snapshot,
      contentVersion: 4,
    })
    const save = vi.spyOn(h.server, 'save')

    const result = objectOf(await call('document_read', { documentId: documentIdFor(h, path) }))

    expect(result['status']).toBe('ok')
    expect(result['contentVersion']).toBe(4)
    expect(objectOf(result['summary'])['kind']).toBe('workbook')
    expect(save).not.toHaveBeenCalled()
    expect(dispatch).toHaveBeenCalledOnce()
  })

  it.each(['document-not-open', 'timeout', 'pending-edits'] as const)('maps a %s answer', async (status) => {
    const path = await h.copy('roundtrip.xlsx')
    const outcomes = {
      'document-not-open': { status: 'document-not-open' },
      timeout: { status: 'timeout' },
      'pending-edits': { status: 'pending-edits' },
    } as const
    vi.spyOn(h.server, 'dispatchCommand').mockResolvedValue(outcomes[status])
    const result = objectOf(await call('document_read', { documentId: documentIdFor(h, path) }))
    expect(result['status']).toBe('error')
    expect(result['code']).toBe(status)
  })

  it('maps conflict and unsupported answers', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const dispatch = vi.spyOn(h.server, 'dispatchCommand')
    dispatch.mockResolvedValueOnce({ status: 'conflict', currentVersion: 9 })
    let result = objectOf(await call('document_read', { documentId: documentIdFor(h, path) }))
    expect(result['code']).toBe('conflict')
    expect(result['currentVersion']).toBe(9)

    dispatch.mockResolvedValueOnce({ status: 'unsupported-feature', features: [{ kind: 'macro' }] })
    result = objectOf(await call('document_read', { documentId: documentIdFor(h, path) }))
    expect(result['code']).toBe('unsupported-feature')
    expect(result['features']).toEqual([{ kind: 'macro' }])

    dispatch.mockResolvedValueOnce({
      status: 'unsupported-feature',
      features: [{ kind: 'comment', detail: 'Sales' }],
    })
    result = objectOf(await call('document_read', { documentId: documentIdFor(h, path) }))
    expect(result['code']).toBe('unsupported-feature')
    expect(result['features']).toEqual([{ kind: 'comment', detail: 'Sales' }])
  })

  it('refuses an applied answer for a read command', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(path)
    vi.spyOn(h.server, 'dispatchCommand').mockResolvedValue({
      status: 'applied',
      snapshot: read.snapshot,
      contentVersion: 1,
    })
    const result = objectOf(await call('document_read', { documentId: documentIdFor(h, path) }))
    expect(result['code']).toBe('conversion-failed')
  })

  it('refuses an identity from another Session and an empty path', async () => {
    let result = objectOf(await call('document_read', { documentId: 'other-session:report.xlsx' }))
    expect(result['code']).toBe('document-missing')
    result = objectOf(await call('document_read', { documentId: `${h.agent.session.id}:` }))
    expect(result['code']).toBe('document-missing')
  })

  it('refuses a missing file and a path outside the workspace', async () => {
    const missing = new RemoteError('office-document/not-found', 'No regular file.', { path: 'missing.xlsx' })
    vi.spyOn(h.server, 'dispatchCommand').mockRejectedValue(missing)
    let result = objectOf(await call('document_read', { documentId: documentIdFor(h, 'missing.xlsx') }))
    expect(result['code']).toBe('document-missing')
    result = objectOf(await call('document_read', { documentId: documentIdFor(h, '../outside.xlsx') }))
    expect(result['code']).toBe('out-of-scope')
  })
})

describe('document_edit_cells', () => {
  it('applies one command, saves once, and appends one edit fact with its inverse', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    const dispatch = vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async (request) => {
      const current = await readWorkbook(path)
      expect(request.operation).toEqual({
        kind: 'edit',
        steps: [{ kind: 'cell', sheetId, row: 1, column: 0, value: 'Sprocket' }],
      })
      return {
        status: 'applied',
        snapshot: withCell(current.snapshot, sheetId, 1, 0, 'Sprocket'),
        contentVersion: request.fromVersion + 1,
      }
    })
    const save = vi.spyOn(h.server, 'save')

    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ sheetId, row: 1, column: 0, value: 'Sprocket' }],
    }))

    expect(result['status']).toBe('ok')
    expect(result['contentVersion']).toBe(1)
    expect(save).toHaveBeenCalledOnce()
    expect(dispatch).toHaveBeenCalledOnce()
    expect(editFacts()).toBe(1)
    const edit = firstEdit()
    expect(edit.data.actor).toBe('agent')
    expect(edit.data.fromVersion).toBe(0)
    expect(edit.data.toVersion).toBe(1)
    expect(edit.data.inverse).toEqual([{ kind: 'cell', sheetId, row: 1, column: 0, value: 'Widget' }])

    const reread = await readWorkbook(path)
    expect(cellValue(reread.snapshot, sheetId, 1, 0)).toBe('Sprocket')
  })

  it('inserts worksheet rows and records a structural inverse', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    const dispatch = vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async (request) => {
      const current = await readWorkbook(path)
      expect(request.operation).toEqual({
        kind: 'edit',
        steps: [{ kind: 'row', sheetId, row: 1, count: 1, action: 'insert' }],
      })
      return {
        status: 'applied',
        snapshot: withRows(current.snapshot, sheetId, 1, 1, 'insert'),
        contentVersion: request.fromVersion + 1,
      }
    })
    const save = vi.spyOn(h.server, 'save')

    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ kind: 'row', sheetId, row: 1, count: 1, action: 'insert' }],
    }))

    expect(result['status']).toBe('ok')
    expect(save).toHaveBeenCalledOnce()
    expect(dispatch).toHaveBeenCalledOnce()
    const edit = firstEdit()
    expect(edit.data.operation).toEqual([{ kind: 'row', sheetId, row: 1, count: 1, action: 'insert' }])
    expect(edit.data.inverse).toEqual([{ kind: 'row', sheetId, row: 1, count: 1, action: 'delete' }])
    const reread = await readWorkbook(path)
    expect(sheetSize(reread.snapshot, sheetId).rows).toBe(8)
    expect(cellValue(reread.snapshot, sheetId, 2, 0)).toBe('Widget')
    expect(cellValue(reread.snapshot, sheetId, 4, 0)).toBe('Cog')
  })

  it('deletes worksheet rows and restores their cells in the inverse', async () => {
    const path = await copyMergeFreeWorkbook('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async (request) => {
      const current = await readWorkbook(path)
      expect(request.operation).toEqual({
        kind: 'edit',
        steps: [{ kind: 'row', sheetId, row: 1, count: 1, action: 'delete' }],
      })
      return {
        status: 'applied',
        snapshot: withRows(current.snapshot, sheetId, 1, 1, 'delete'),
        contentVersion: request.fromVersion + 1,
      }
    })

    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ kind: 'row', sheetId, row: 1, count: 1, action: 'delete' }],
    }))

    expect(result['status']).toBe('ok')
    const edit = firstEdit()
    expect(edit.data.operation).toEqual([{ kind: 'row', sheetId, row: 1, count: 1, action: 'delete' }])
    expect(edit.data.inverse[0]).toEqual({ kind: 'row', sheetId, row: 1, count: 1, action: 'insert' })
    expect(edit.data.inverse).toContainEqual({ kind: 'cell', sheetId, row: 1, column: 0, value: 'Widget' })
    const reread = await readWorkbook(path)
    expect(cellValue(reread.snapshot, sheetId, 1, 0)).toBe('Gadget')
    expect(cellValue(reread.snapshot, sheetId, 2, 0)).toBe('Cog')
  })

  it('inserts worksheet columns and records a structural inverse', async () => {
    const path = await copyMergeFreeWorkbook('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async (request) => {
      const current = await readWorkbook(path)
      expect(request.operation).toEqual({
        kind: 'edit',
        steps: [{ kind: 'column', sheetId, column: 1, count: 1, action: 'insert' }],
      })
      return {
        status: 'applied',
        snapshot: withColumns(current.snapshot, sheetId, 1, 1, 'insert'),
        contentVersion: request.fromVersion + 1,
      }
    })

    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ kind: 'column', sheetId, column: 1, count: 1, action: 'insert' }],
    }))

    expect(result['status']).toBe('ok')
    const edit = firstEdit()
    expect(edit.data.operation).toEqual([{ kind: 'column', sheetId, column: 1, count: 1, action: 'insert' }])
    expect(edit.data.inverse).toEqual([{ kind: 'column', sheetId, column: 1, count: 1, action: 'delete' }])
    const reread = await readWorkbook(path)
    expect(sheetSize(reread.snapshot, sheetId).columns).toBe(4)
    expect(cellValue(reread.snapshot, sheetId, 0, 2)).toBe('Units')
    expect(cellValue(reread.snapshot, sheetId, 0, 3)).toBe('Revenue')
  })

  it('deletes worksheet columns and restores their cells in the inverse', async () => {
    const path = await copyMergeFreeWorkbook('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async (request) => {
      const current = await readWorkbook(path)
      expect(request.operation).toEqual({
        kind: 'edit',
        steps: [{ kind: 'column', sheetId, column: 1, count: 1, action: 'delete' }],
      })
      return {
        status: 'applied',
        snapshot: withColumns(current.snapshot, sheetId, 1, 1, 'delete'),
        contentVersion: request.fromVersion + 1,
      }
    })

    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ kind: 'column', sheetId, column: 1, count: 1, action: 'delete' }],
    }))

    expect(result['status']).toBe('ok')
    const edit = firstEdit()
    expect(edit.data.operation).toEqual([{ kind: 'column', sheetId, column: 1, count: 1, action: 'delete' }])
    expect(edit.data.inverse[0]).toEqual({ kind: 'column', sheetId, column: 1, count: 1, action: 'insert' })
    expect(edit.data.inverse).toContainEqual({ kind: 'cell', sheetId, row: 1, column: 1, value: 4 })
    const reread = await readWorkbook(path)
    expect(cellValue(reread.snapshot, sheetId, 0, 1)).toBe('Revenue')
    expect(cellValue(reread.snapshot, sheetId, 0, 2)).toBe(null)
  })

  it('refuses a structural shift that would leave a merged range behind', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async (request) => {
      const current = await readWorkbook(path)
      return {
        status: 'applied',
        snapshot: withRows(current.snapshot, sheetId, 1, 1, 'delete'),
        contentVersion: request.fromVersion + 1,
      }
    })
    const before = await hashOfFile(join(h.root, path))

    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ kind: 'row', sheetId, row: 1, count: 1, action: 'delete' }],
    }))

    expect(result['status']).toBe('error')
    expect(result['code']).toBe('unsupported-feature')
    expect(result['message']).toContain('merged range A7:C7')
    expect(await hashOfFile(join(h.root, path))).toBe(before)
    expect(editFacts()).toBe(0)
    expect(h.events.filter(event => event.type === 'document/save')).toHaveLength(0)
  })

  it('styles a worksheet range and restores prior styles in the inverse', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    const range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async (request) => {
      const current = await readWorkbook(path)
      expect(request.operation).toEqual({
        kind: 'edit',
        steps: [{ kind: 'style', sheetId, ...range, style: { it: 1 } }],
      })
      return {
        status: 'applied',
        snapshot: withStyle(current.snapshot, sheetId, range, { it: 1 }),
        contentVersion: request.fromVersion + 1,
      }
    })

    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ kind: 'style', sheetId, ...range, style: { it: 1 } }],
    }))

    expect(result['status']).toBe('ok')
    const edit = firstEdit()
    expect(edit.data.operation).toEqual([{ kind: 'style', sheetId, ...range, style: { it: 1 } }])
    expect(edit.data.inverse[0]).toEqual({ kind: 'style', sheetId, ...range, style: null })
    expect(edit.data.inverse.length).toBeGreaterThan(1)
    const reread = await readWorkbook(path)
    const style = objectOf(cellStyle(reread.snapshot, sheetId, 0, 0))
    expect(style?.['it']).toBe(1)
  })

  it('applies a mixed batch through one command', async () => {
    const path = await copyMergeFreeWorkbook('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    const steps = [
      { sheetId, row: 5, column: 0, value: 'Batch' },
      { kind: 'row', sheetId, row: 6, count: 1, action: 'insert' },
      { kind: 'column', sheetId, column: 2, count: 1, action: 'insert' },
      { kind: 'style', sheetId, startRow: 5, startColumn: 0, endRow: 5, endColumn: 0, style: { it: 1 } },
    ] as const
    const dispatch = vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async (request) => {
      const current = await readWorkbook(path)
      expect(request.operation).toEqual({
        kind: 'edit',
        steps: [{ kind: 'cell', sheetId, row: 5, column: 0, value: 'Batch' }, steps[1], steps[2], steps[3]],
      })
      let snapshot = withCell(current.snapshot, sheetId, 5, 0, 'Batch')
      snapshot = withRows(snapshot, sheetId, 6, 1, 'insert')
      snapshot = withColumns(snapshot, sheetId, 2, 1, 'insert')
      snapshot = withStyle(snapshot, sheetId, {
        startRow: 5,
        startColumn: 0,
        endRow: 5,
        endColumn: 0,
      }, { it: 1 })
      return { status: 'applied', snapshot, contentVersion: request.fromVersion + 1 }
    })

    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [steps[0], steps[1], steps[2], steps[3]],
    }))

    expect(result['status']).toBe('ok')
    expect(dispatch).toHaveBeenCalledOnce()
    const reread = await readWorkbook(path)
    expect(sheetSize(reread.snapshot, sheetId)).toEqual({ rows: 8, columns: 4 })
    expect(cellValue(reread.snapshot, sheetId, 5, 0)).toBe('Batch')
    expect(cellValue(reread.snapshot, sheetId, 0, 3)).toBe('Revenue')
    expect(objectOf(cellStyle(reread.snapshot, sheetId, 5, 0))?.['it']).toBe(1)
  })

  it('attributes its guarded save to the agent and the open turn without opening twice', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    h.agent.session.append('turn/start', { turn: 3 })
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async request => ({
      status: 'applied',
      snapshot: withCell(read.snapshot, sheetId, 1, 0, 'Sprocket'),
      contentVersion: request.fromVersion + 1,
    }))

    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ sheetId, row: 1, column: 0, value: 'Sprocket' }],
    }))

    expect(result['status']).toBe('ok')
    expect(h.events.filter(event => event.type === 'document/open')).toHaveLength(1)
    const saves = h.events.filter(event => event.type === 'document/save')
    expect(saves).toHaveLength(1)
    expect(saves[0]?.data).toEqual({
      documentId: documentIdFor(h, path),
      actor: 'agent',
      contentVersion: 1,
      turn: 3,
      disposition: {
        schemaVersion: 1,
        verbatimCount: 10,
        spliced: { count: 1, names: ['xl/worksheets/sheet1.xml'] },
        added: { count: 0, names: [] },
        removed: { count: 0, names: [] },
      },
    })
  })

  it('restores a prior formula in the inverse', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const base = await readWorkbook(path)
    const sheetId = firstSheet(base.snapshot)
    const snapshot = withCell(base.snapshot, sheetId, 1, 0, 'Widget', '=A1')
    vi.spyOn(h.server, 'read').mockResolvedValue({ ...base, snapshot })
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async request => ({
      status: 'applied',
      snapshot,
      contentVersion: request.fromVersion + 1,
    }))
    await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ sheetId, row: 1, column: 0, value: 'Sprocket' }],
    })
    const edit = firstEdit()
    expect(edit.data.operation).toEqual([{ kind: 'cell', sheetId, row: 1, column: 0, value: 'Sprocket' }])
    expect(edit.data.inverse).toEqual([
      { kind: 'cell', sheetId, row: 1, column: 0, value: 'Widget' },
      { kind: 'formula', sheetId, row: 1, column: 0, formula: '=A1' },
    ])
  })

  it.each([
    [{ status: 'pending-edits' }, 'pending-edits'],
    [{ status: 'conflict', currentVersion: 5 }, 'conflict'],
    [{ status: 'unsupported-feature', features: [{ kind: 'macro' }] }, 'unsupported-feature'],
    [{ status: 'document-not-open' }, 'document-not-open'],
    [{ status: 'timeout' }, 'timeout'],
    [{ status: 'read', snapshot: { kind: 'workbook', data: {} }, contentVersion: 0 }, 'conversion-failed'],
  ] as const)('writes nothing for %o', async (outcome, code) => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    vi.spyOn(h.server, 'dispatchCommand').mockResolvedValue(outcome)
    const save = vi.spyOn(h.server, 'save')

    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ sheetId, row: 1, column: 0, value: 'Sprocket' }],
    }))

    expect(result['code']).toBe(code)
    expect(save).not.toHaveBeenCalled()
    expect(editFacts()).toBe(0)
  })

  it('records a null value and a null inverse for an empty cell', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async request => ({
      status: 'applied',
      snapshot: read.snapshot,
      contentVersion: request.fromVersion + 1,
    }))
    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ sheetId, row: 10, column: 0, value: null }],
    }))
    expect(result['status']).toBe('ok')
    const edit = firstEdit()
    expect(edit.data.operation).toEqual([{ kind: 'cell', sheetId, row: 10, column: 0, value: null }])
    expect(edit.data.inverse).toEqual([{ kind: 'cell', sheetId, row: 10, column: 0, value: null }])
  })

  it('applies a formula edit and records a clear-formula inverse', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async request => ({
      status: 'applied',
      snapshot: read.snapshot,
      contentVersion: request.fromVersion + 1,
    }))
    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ sheetId, row: 1, column: 0, formula: '=SUM(B1:B2)' }],
    }))
    expect(result['status']).toBe('ok')
    const edit = firstEdit()
    expect(edit.data.operation).toEqual([{ kind: 'formula', sheetId, row: 1, column: 0, formula: '=SUM(B1:B2)' }])
    expect(edit.data.inverse).toEqual([{ kind: 'formula', sheetId, row: 1, column: 0, formula: null }])
  })

  it('refuses an empty operation list and a cell edit against a text document', async () => {
    const workbook = await h.copy('roundtrip.xlsx')
    let result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, workbook),
      fromVersion: 0,
      operations: [],
    }))
    expect(result['code']).toBe('invalid-request')

    const document = await h.copy('roundtrip.docx')
    result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, document),
      fromVersion: 0,
      operations: [{ sheetId: 's', row: 0, column: 0, value: 1 }],
    }))
    expect(result['code']).toBe('invalid-request')
    expect(result['message']).toContain('.xlsx')
  })

  it('refuses a paragraph edit against a document without paragraph structure', async () => {
    const path = await h.copy('roundtrip.docx')
    vi.spyOn(h.server, 'read').mockResolvedValue({
      format: 'docx',
      snapshot: { kind: 'document', data: { body: { dataStream: 'x' } } },
      unsupported: [],
      version: 'v1',
    })
    const result = objectOf(await call('document_edit_paragraph', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ index: 0, text: 'x' }],
    }))
    expect(result['code']).toBe('invalid-request')
  })

  it('refuses when the editor applied a different version', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    vi.spyOn(h.server, 'dispatchCommand').mockResolvedValue({
      status: 'applied',
      snapshot: read.snapshot,
      contentVersion: 9,
    })
    const save = vi.spyOn(h.server, 'save')
    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ sheetId, row: 1, column: 0, value: 'Sprocket' }],
    }))
    expect(result['code']).toBe('conversion-failed')
    expect(save).not.toHaveBeenCalled()
    expect(editFacts()).toBe(0)
  })

  it('refuses when the file changed between the read and the save', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(path)
    const sheetId = firstSheet(read.snapshot)
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async (request) => {
      await writeFile(join(h.root, path), 'externally replaced')
      return { status: 'applied', snapshot: read.snapshot, contentVersion: request.fromVersion + 1 }
    })
    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ sheetId, row: 1, column: 0, value: 'Sprocket' }],
    }))
    expect(result['code']).toBe('conflict')
    expect(editFacts()).toBe(0)
  })

  it('saves a non-intersecting edit and preserves the construct a read reported (FR-021)', async () => {
    const path = await h.copy('unsupported-tracked.docx')
    const base = await h.server.read(h.scope, path, SIGNAL)
    expect(base.unsupported.map(feature => feature.kind)).toContain('tracked-changes')
    const snapshot = structuredClone(base.snapshot)
    const body = objectOf(objectOf(snapshot.data)['body'])
    body['dataStream'] = streamOf(body).replace('Quarterly', 'Financial')
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async request => ({
      status: 'applied',
      snapshot,
      contentVersion: request.fromVersion + 1,
    }))

    const result = objectOf(await call('document_edit_paragraph', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ index: 0, text: 'Financial Report' }],
    }))

    expect(result['status']).toBe('ok')
    const reread = await h.server.read(h.scope, path, SIGNAL)
    expect(objectOf(objectOf(reread.snapshot.data)['body'])['dataStream']).toContain('Financial Report')
    expect(reread.unsupported.map(feature => feature.kind)).toContain('tracked-changes')
    expect(h.events.filter(event => event.type === 'document/save')).toHaveLength(1)
    expect(editFacts()).toBe(1)
  })

  it('refuses a coupled edit through the same agent save path', async () => {
    const path = await h.copy('preserve-comment.docx')
    const base = await h.server.read(h.scope, path, SIGNAL)
    const snapshot = structuredClone(base.snapshot)
    const body = objectOf(objectOf(snapshot.data)['body'])
    body['dataStream'] = streamOf(body).replace('Commented', 'Reviewed!')
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async request => ({
      status: 'applied',
      snapshot,
      contentVersion: request.fromVersion + 1,
    }))
    const before = await hashOfFile(join(h.root, path))

    const result = objectOf(await call('document_edit_paragraph', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ index: 0, text: 'Reviewed text' }],
    }))

    expect(result['code']).toBe('unsupported-feature')
    expect(result['message']).toContain('commentRangeStart')
    expect(await hashOfFile(join(h.root, path))).toBe(before)
    expect(editFacts()).toBe(0)
    expect(h.events.filter(event => event.type === 'document/save')).toHaveLength(0)
  })

  it('honours the configured operation bound', async () => {
    await disposeHarness(h)
    h = await createHarness({ maxEditSteps: 1 })
    const path = await h.copy('roundtrip.xlsx')
    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [
        { sheetId: 's', row: 0, column: 0, value: 1 },
        { sheetId: 's', row: 0, column: 1, value: 2 },
      ],
    }))
    expect(result['code']).toBe('bound-exceeded')
  })

  it('refuses invalid operations without dispatching', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const documentId = documentIdFor(h, path)
    const dispatch = vi.spyOn(h.server, 'dispatchCommand')
    const cases: readonly Record<string, unknown>[] = [
      { sheetId: '', row: 0, column: 0, value: 1 },
      { sheetId: 's', row: -1, column: 0, value: 1 },
      { sheetId: 's', row: 0, column: 0 },
      { sheetId: 's', row: 0, column: 0, value: 1, formula: '=A1' },
      { sheetId: 's', row: 0, column: 0, formula: '' },
      { kind: 'row', sheetId: 's', row: -1, count: 1, action: 'insert' },
      { kind: 'row', sheetId: 's', row: 0, count: 0, action: 'delete' },
      { kind: 'column', sheetId: '', column: 0, count: 1, action: 'insert' },
      { kind: 'column', sheetId: 's', column: 0, count: -1, action: 'delete' },
      { kind: 'style', sheetId: 's', startRow: 1, startColumn: 0, endRow: 0, endColumn: 0, style: {} },
      { kind: 'style', sheetId: 's', startRow: 0, startColumn: 0, endRow: 0, endColumn: 0, style: 5 },
    ]
    for (const operation of cases) {
      const result = objectOf(await call('document_edit_cells', {
        documentId,
        fromVersion: 0,
        operations: [operation],
      }))
      expect(result['code']).toBe('invalid-request')
    }
    expect(dispatch).not.toHaveBeenCalled()
  })

  it('rejects an unsupported operation shape before dispatch', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const documentId = documentIdFor(h, path)
    const dispatch = vi.spyOn(h.server, 'dispatchCommand')
    const cases: readonly Record<string, unknown>[] = [
      { kind: 'row', sheetId: 'Sales', row: 0, count: 1 },
      { kind: 'merge', sheetId: 'Sales', startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 },
      { kind: 'style', sheetId: 'Sales', row: 0, column: 0, value: 1 },
    ]
    for (const operation of cases) {
      const result = await h.execute('document_edit_cells', {
        documentId,
        fromVersion: 0,
        operations: [operation],
      }, SIGNAL)
      expect(result.isError).toBe(true)
      if (result.isError) expect(result.error.message).toContain('invalid arguments')
    }
    expect(dispatch).not.toHaveBeenCalled()
  })

  it('refuses an edit whose negative version cannot describe a document', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: -1,
      operations: [{ sheetId: 's', row: 0, column: 0, value: 1 }],
    }))
    expect(result['code']).toBe('invalid-request')
  })

  it('refuses an unknown sheet and an unknown paragraph', async () => {
    const workbook = await h.copy('roundtrip.xlsx')
    const read = await readWorkbook(workbook)
    const sheetId = firstSheet(read.snapshot)
    vi.spyOn(h.server, 'dispatchCommand').mockResolvedValue({
      status: 'applied',
      snapshot: read.snapshot,
      contentVersion: 1,
    })
    let result = objectOf(await call('document_edit_cells', {
      documentId: documentIdFor(h, workbook),
      fromVersion: 0,
      operations: [{ sheetId: 'missing', row: 0, column: 0, value: 1 }],
    }))
    expect(result['code']).toBe('invalid-request')

    const document = await h.copy('roundtrip.docx')
    result = objectOf(await call('document_edit_paragraph', {
      documentId: documentIdFor(h, document),
      fromVersion: 0,
      operations: [{ index: 999, text: 'x' }],
    }))
    expect(result['code']).toBe('invalid-request')
    expect(result['message']).toContain('paragraph 999')
    void sheetId
  })
})

describe('document_edit_paragraph', () => {
  it('saves once and appends one edit fact with the prior text as its inverse', async () => {
    const path = await h.copy('roundtrip.docx')
    const read = await readWorkbook(path)
    vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async request => ({
      status: 'applied',
      snapshot: read.snapshot,
      contentVersion: request.fromVersion + 1,
    }))
    const save = vi.spyOn(h.server, 'save')

    const result = objectOf(await call('document_edit_paragraph', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ index: 0, text: 'Updated paragraph' }],
    }))

    expect(result['status']).toBe('ok')
    expect(save).toHaveBeenCalledOnce()
    expect(editFacts()).toBe(1)
    const edit = firstEdit()
    expect(edit.data.operation).toEqual([{ kind: 'paragraph', index: 0, text: 'Updated paragraph' }])
    expect(edit.data.inverse).toHaveLength(1)
    const inverse = edit.data.inverse[0]
    if (inverse?.kind !== 'paragraph') throw new Error('expected a paragraph inverse')
    expect(inverse.text).not.toBe('Updated paragraph')
  })

  it('refuses a negative paragraph index', async () => {
    const path = await h.copy('roundtrip.docx')
    const result = objectOf(await call('document_edit_paragraph', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ index: -1, text: 'x' }],
    }))
    expect(result['code']).toBe('invalid-request')
  })

  it('refuses a paragraph edit against a workbook', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const result = objectOf(await call('document_edit_paragraph', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ index: 0, text: 'x' }],
    }))
    expect(result['code']).toBe('invalid-request')
  })
})

describe('live renderer target', () => {
  it('refuses promptly with document-not-open when no renderer target is registered', async () => {
    await disposeHarness(h)
    h = await createHarness({ commandTimeoutMs: 1_000 })
    const path = await h.copy('roundtrip.xlsx')
    const started = performance.now()

    const result = objectOf(await call('document_read', { documentId: documentIdFor(h, path) }))

    expect(result['code']).toBe('document-not-open')
    expect(performance.now() - started).toBeLessThan(500)
    expect(h.events.filter(event => event.type === 'document/command')).toHaveLength(0)
  })

  it('applies an edit through a registered target and saves it once', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const initial = await readWorkbook(path)
    const sheetId = firstSheet(initial.snapshot)
    const detach = await attachLiveTarget(h, path, async (request) => {
      const current = await readWorkbook(path)
      if (request.operation.kind !== 'edit') {
        return { status: 'read', snapshot: current.snapshot, contentVersion: request.fromVersion }
      }
      return {
        status: 'applied',
        snapshot: withCell(current.snapshot, sheetId, 1, 0, 'Sprocket'),
        contentVersion: request.fromVersion + 1,
      }
    })
    try {
      const save = vi.spyOn(h.server, 'save')
      const result = objectOf(await call('document_edit_cells', {
        documentId: documentIdFor(h, path),
        fromVersion: 0,
        operations: [{ sheetId, row: 1, column: 0, value: 'Sprocket' }],
      }))

      expect(result['status']).toBe('ok')
      expect(result['contentVersion']).toBe(1)
      expect(save).toHaveBeenCalledOnce()
      expect(editFacts()).toBe(1)
      const reread = await readWorkbook(path)
      expect(cellValue(reread.snapshot, sheetId, 1, 0)).toBe('Sprocket')
    } finally {
      await detach()
    }
  })

  it('times out when a registered target never answers', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const detach = await attachLiveTarget(h, path, async () => undefined)
    try {
      const save = vi.spyOn(h.server, 'save')
      const result = objectOf(await call('document_read', { documentId: documentIdFor(h, path) }))

      expect(result['code']).toBe('timeout')
      expect(save).not.toHaveBeenCalled()
      expect(editFacts()).toBe(0)
    } finally {
      await detach()
    }
  })
})

describe('document_convert', () => {
  it('returns the converted path and format', async () => {
    const path = await h.copy('legacy.xls')
    vi.spyOn(h.server, 'convertLegacy').mockResolvedValue(join(h.root, 'converted.docx'))
    const result = objectOf(await call('document_convert', { path }))
    expect(result).toEqual({ status: 'ok', path: join(h.root, 'converted.docx'), format: 'docx' })
  })

  it('refuses an empty path, an out-of-root path, and a non-editable output', async () => {
    let result = objectOf(await call('document_convert', { path: '   ' }))
    expect(result['code']).toBe('invalid-request')
    result = objectOf(await call('document_convert', { path: '../outside.xls' }))
    expect(result['code']).toBe('out-of-scope')
    await h.copy('legacy.xls')
    vi.spyOn(h.server, 'convertLegacy').mockResolvedValue(join(h.root, 'converted.ppt'))
    result = objectOf(await call('document_convert', { path: 'legacy.xls' }))
    expect(result['code']).toBe('conversion-failed')
  })
})

describe('execution context', () => {
  it('refuses a Session with no workspace root', async () => {
    const agent = await h.createAgent(SessionId('no-cwd'))
    const result = await h.execute('document_read', { documentId: 'x:y' }, SIGNAL, agent)
    expect(result.isError).toBe(false)
    expect(result.value).toMatchObject({ status: 'error', code: 'out-of-scope' })
  })

  it('requires a calling agent', async () => {
    const runtime: DocumentToolRuntime = {
      documents: h.server,
      commands: {
        execute: vi.fn(),
        currentVersion: () => 0,
        remember: () => {},
      },
      turnOf: () => undefined,
    }
    await expect(executeDocumentTool(
      { name: 'document_read', documentId: 'x' },
      runtime,
      { summaryCells: 10, maxEditSteps: 10 },
      { signal: SIGNAL },
    )).rejects.toThrow(/require a calling agent/)
  })

  it('registers all four tools with their schemas', () => {
    const registered: { name: string; parameters: Record<string, unknown> }[] = []
    const runtime: DocumentToolRuntime = {
      documents: h.server,
      commands: { execute: vi.fn(), currentVersion: () => 0, remember: () => {} },
      turnOf: () => undefined,
    }
    registerDocumentTools({ register: (definition) => { registered.push(definition) } }, runtime, {
      summaryCells: 10,
      maxEditSteps: 10,
    })
    expect(registered.map(definition => definition.name)).toEqual([
      'document_read',
      'document_edit_cells',
      'document_edit_paragraph',
      'document_convert',
    ])
    expect(registered[0]?.parameters).toHaveProperty('properties.documentId')
    expect(registered[1]?.parameters).toHaveProperty('properties.operations')
    expect(registered[1]?.parameters).toHaveProperty('properties.operations.items.oneOf')
  })

  it('reports a classified failure when the Remote is mounted without the command broker', async () => {
    const bare = new Context()
    try {
      await bare.plugin(LocalFileSystem, { cwd: h.root })
      await mountAgentLoopTestDependencies(bare)
      await bare.plugin(OfficeDocumentRemote, { maxBytes: 1024 * 1024 })
      await bare.plugin(ToolOffice)
      const loop = await mountAgentLoopTestHarness(bare)
      const agent = await loop.create(SessionId('no-broker'), {}, { cwd: h.root })
      await h.copy('roundtrip.xlsx')
      const result = await executeToolFixture(bare, {
        callId: ToolCallId('no-broker-read'),
        name: 'document_read',
        arguments: { documentId: `${agent.session.id}:roundtrip.xlsx` },
        signal: SIGNAL,
        agent,
      })
      expect(result.isError).toBe(false)
      expect(result.value).toMatchObject({ status: 'error', code: 'conversion-failed' })
    } finally {
      await bare.fiber.dispose()
    }
  })

  it('fails loudly when the Office Remote is not mounted and for invalid caps', () => {
    expect(() => { apply(new Context(), {}) }).toThrow(/officeDocument is absent/)
    expect(() => resolveToolConfig({ summaryCells: 0 })).toThrow(/summaryCells/)
    expect(() => resolveToolConfig({ maxEditSteps: 0 })).toThrow(/maxEditSteps/)
    expect(resolveToolConfig({})).toEqual({ summaryCells: 200, maxEditSteps: 2_000 })
  })
})
