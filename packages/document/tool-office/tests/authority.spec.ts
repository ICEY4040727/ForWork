/** Agent write authority: out-of-root writes are refused and in-root writes are accepted (SC-011). */
import { mkdir, mkdtemp, realpath, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, relative } from 'node:path'
import { assertWithinRoots, OfficeExchangeError } from '@forwork/office-exchange'
import type { OfficeDocumentSnapshot } from '@forwork/office-exchange'
import type { JsonValue } from '@forwork/util-values'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createHarness, disposeHarness, documentIdFor } from './harness.ts'
import type { ToolTestHarness } from './harness.ts'

const SIGNAL = new AbortController().signal
let h: ToolTestHarness
let outside: string
let sibling: string

beforeEach(async () => {
  h = await createHarness()
  outside = await mkdtemp(join(tmpdir(), 'dsh-tool-office-outside-'))
  sibling = `${h.root}-sibling`
  await mkdir(sibling, { recursive: true })
})
afterEach(async () => {
  await disposeHarness(h)
  await rm(outside, { recursive: true, force: true })
  await rm(sibling, { recursive: true, force: true })
})

/** A JSON object view of one tool result value. */
function objectOf(value: JsonValue | undefined): Record<string, JsonValue> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) throw new Error('expected a JSON object')
  return value
}

/** Execute one tool call and require a structured value, never a thrown tool error. */
async function call(name: string, arguments_: Record<string, unknown>): Promise<Record<string, JsonValue>> {
  const result = await h.execute(name, arguments_, SIGNAL)
  expect(result.isError).toBe(false)
  return objectOf(result.value)
}

/** Require the documented out-of-scope refusal and its named reason. */
function expectOutOfScope(result: Record<string, JsonValue>): void {
  expect(result['status']).toBe('error')
  expect(result['code']).toBe('out-of-scope')
  const message = result['message']
  if (typeof message !== 'string') throw new Error('expected a refusal message')
  expect(message).toContain('outside the permitted write roots')
}

/** The first sheet id of a workbook snapshot. */
function firstSheet(snapshot: OfficeDocumentSnapshot): string {
  if (snapshot.kind !== 'workbook') throw new Error('expected a workbook')
  const order = objectOf(snapshot.data)['sheetOrder']
  const id = Array.isArray(order) ? order[0] : undefined
  if (typeof id !== 'string') throw new Error('expected a sheet id')
  return id
}

describe('assertWithinRoots write confinement', () => {
  it('accepts an in-root target that exists, is missing, or is the root itself', async () => {
    const existing = join(h.root, 'report.xlsx')
    await writeFile(existing, 'x')
    const root = await realpath(h.root)

    await expect(assertWithinRoots(existing, [h.root], { allowMissing: false })).resolves.toBe(await realpath(existing))
    await expect(assertWithinRoots(join(h.root, 'nested', 'new.xlsx'), [h.root])).resolves.toBe(
      join(root, 'nested', 'new.xlsx'),
    )
    await expect(assertWithinRoots(h.root, [h.root])).resolves.toBe(root)
  })

  it('refuses an existing target outside every root', async () => {
    const target = join(outside, 'report.xlsx')
    await writeFile(target, 'x')

    const error = await assertWithinRoots(target, [h.root]).then(() => undefined, (cause: unknown) => cause)
    expect(error).toBeInstanceOf(OfficeExchangeError)
    if (!(error instanceof OfficeExchangeError)) throw new Error('expected an OfficeExchangeError')
    expect(error.code).toBe('out-of-scope')
    expect(error.message).toContain('outside the permitted write roots')
  })

  it('refuses a missing target outside every root', async () => {
    await expect(assertWithinRoots(join(outside, 'missing.xlsx'), [h.root])).rejects.toMatchObject({
      code: 'out-of-scope',
    })
  })

  it('refuses a sibling whose name shares the root prefix', async () => {
    const target = join(sibling, 'report.xlsx')
    await writeFile(target, 'x')

    await expect(assertWithinRoots(target, [h.root])).rejects.toMatchObject({ code: 'out-of-scope' })
  })
})

describe('agent write tools', () => {
  it('refuses cell and paragraph edits outside the workspace root', async () => {
    const workbook = join(outside, 'report.xlsx')
    const document = join(outside, 'report.docx')
    await writeFile(workbook, 'x')
    await writeFile(document, 'x')
    const read = vi.spyOn(h.server, 'read')
    const dispatch = vi.spyOn(h.server, 'dispatchCommand')
    const save = vi.spyOn(h.server, 'save')

    expectOutOfScope(await call('document_edit_cells', {
      documentId: documentIdFor(h, workbook),
      fromVersion: 0,
      operations: [{ sheetId: 's', row: 0, column: 0, value: 1 }],
    }))
    expectOutOfScope(await call('document_edit_paragraph', {
      documentId: documentIdFor(h, document),
      fromVersion: 0,
      operations: [{ index: 0, text: 'x' }],
    }))

    expect(read).not.toHaveBeenCalled()
    expect(dispatch).not.toHaveBeenCalled()
    expect(save).not.toHaveBeenCalled()
    expect(h.events.filter(event => event.type === 'document/edit')).toHaveLength(0)
  })

  it('refuses a relative path that escapes the workspace root', async () => {
    const target = join(outside, 'report.xlsx')
    await writeFile(target, 'x')

    expectOutOfScope(await call('document_edit_cells', {
      documentId: documentIdFor(h, relative(h.root, target)),
      fromVersion: 0,
      operations: [{ sheetId: 's', row: 0, column: 0, value: 1 }],
    }))
  })

  it('refuses conversion of a legacy file outside the workspace root', async () => {
    const legacy = join(outside, 'legacy.xls')
    await writeFile(legacy, 'x')
    const convertLegacy = vi.spyOn(h.server, 'convertLegacy')

    expectOutOfScope(await call('document_convert', { path: legacy }))
    expect(convertLegacy).not.toHaveBeenCalled()
  })

  it('accepts an in-root cell edit and saves once', async () => {
    const path = await h.copy('roundtrip.xlsx')
    const read = await h.server.read(h.scope, path, SIGNAL)
    const sheetId = firstSheet(read.snapshot)
    const dispatch = vi.spyOn(h.server, 'dispatchCommand').mockImplementation(async request => ({
      status: 'applied',
      snapshot: read.snapshot,
      contentVersion: request.fromVersion + 1,
    }))
    const save = vi.spyOn(h.server, 'save').mockResolvedValue({ version: 'v2' })

    const result = await call('document_edit_cells', {
      documentId: documentIdFor(h, path),
      fromVersion: 0,
      operations: [{ sheetId, row: 1, column: 0, value: 'Sprocket' }],
    })

    expect(result['status']).toBe('ok')
    expect(dispatch).toHaveBeenCalledOnce()
    expect(save).toHaveBeenCalledOnce()
    expect(h.events.filter(event => event.type === 'document/edit')).toHaveLength(1)
  })

  it('accepts conversion of an in-root legacy file', async () => {
    const path = await h.copy('legacy.xls')
    const converted = join(h.root, 'legacy.docx')
    vi.spyOn(h.server, 'convertLegacy').mockResolvedValue(converted)

    expect(await call('document_convert', { path })).toEqual({
      status: 'ok',
      path: converted,
      format: 'docx',
    })
  })

  it('refuses an out-of-root write attempted through the Remote save seam', async () => {
    const target = join(outside, 'report.xlsx')

    await expect(h.server.save(
      h.scope,
      target,
      { kind: 'workbook', data: {} },
      'xlsx',
      'v1',
      { actor: 'user' },
      SIGNAL,
    )).rejects.toMatchObject({ code: 'office-document/out-of-scope' })
  })
})
