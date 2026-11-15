// @vitest-environment jsdom
/**
 * Office editor timing guarantees.
 *
 * SC-001: a 10,000-cell workbook opens and saves well inside 30 s. The Remote
 * and the editor runtime are test doubles, so the measured work is the hook's
 * read/save path plus a fake mount that walks every cell; no network and no
 * Univer runtime participate.
 *
 * SC-003: a snapshot that changes outside the editor reaches the mounted
 * surface inside the 2 s budget.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { Mock } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import type { JsonValue } from '@deepseek-ai/dsh-util-values'
import type { RemoteResult } from '@deepseek-ai/dsh-typert-protocol'
import type {
  OfficeDocumentReadResult,
  OfficeDocumentSaveResult,
  OfficeDocumentSnapshot,
} from '@deepseek-ai/dsh-office-exchange/types'
import type { OfficeAdapter, OfficeSnapshot } from '../src/client/adapter.ts'
import { OfficePanel } from '../src/client/OfficePanel.tsx'
import type { OfficeDocumentRemote } from '../src/client/useOfficeDocument.ts'
import { en } from '../src/client/locales.ts'
import { panelFixture } from './panel.tsx'

const CELL_COUNT = 10_000
const ABSOLUTE_PATH = '/work/budget.xlsx'
const BUDGET_MS = 30_000
const VISIBILITY_BUDGET_MS = 2000

function success<T>(value: T): RemoteResult<T> {
  return { ok: true, value }
}

function successVoid(): RemoteResult<void> {
  return { ok: true, value: undefined }
}

/** A synthetic workbook carrying one cell per requested cell count. */
function workbook(cells: number): OfficeDocumentSnapshot {
  const cellData: Record<string, JsonValue> = {}
  for (let index = 0; index < cells; index += 1) {
    cellData[`${Math.floor(index / 100)}:${index % 100}`] = { v: index }
  }
  return {
    kind: 'workbook',
    data: { id: 'book', sheetOrder: ['sheet-1'], sheets: { 'sheet-1': { id: 'sheet-1', cellData } } },
  }
}

/** Visit every JSON node of a snapshot, standing in for the runtime's load work. */
function visit(value: JsonValue, onNode: () => void): void {
  onNode()
  if (Array.isArray(value)) {
    for (const item of value) visit(item, onNode)
    return
  }
  if (typeof value === 'object' && value !== null) {
    for (const item of Object.values(value)) visit(item, onNode)
  }
}

function readResult(version: string, content: OfficeDocumentSnapshot): OfficeDocumentReadResult {
  return { format: 'xlsx', snapshot: content, unsupported: [], version }
}

type ReadSignature = OfficeDocumentRemote['read']
type SaveSignature = OfficeDocumentRemote['save']

function makeRemote(): {
  read: Mock<ReadSignature>
  save: Mock<SaveSignature>
  face: OfficeDocumentRemote
} {
  const read = vi.fn<ReadSignature>()
  const save = vi.fn<SaveSignature>().mockResolvedValue(success<OfficeDocumentSaveResult>({ version: 'v2' }))
  return {
    read,
    save,
    face: {
      read,
      save,
      convertLegacy: vi.fn(),
      writeDraft: vi.fn(async () => success({ savedAt: 't' })),
      readDraft: vi.fn(async () => success(null)),
      clearDraft: vi.fn(async () => successVoid()),
    },
  }
}

/** An adapter whose mount walks the snapshot and records what it mounted. */
function recordingAdapter(mounted: OfficeSnapshot[]): OfficeAdapter {
  return {
    async mount(container, input) {
      mounted.push(input.snapshot)
      let nodes = 0
      visit(input.snapshot.data, () => { nodes += 1 })
      container.dataset.officeMount = `${input.kind}:${nodes}`
      return {
        dispose() { delete container.dataset.officeMount },
        getSnapshot: () => input.snapshot,
        apply() { return { status: 'applied' } },
        undo() {},
        redo() {},
        isDirty: () => false,
        clearDirty() {},
      }
    },
  }
}

afterEach(() => {
  cleanup()
  vi.useRealTimers()
})

describe('SC-001 open and save timing', () => {
  it('opens and saves a 10,000-cell workbook inside the budget', async () => {
    const content = workbook(CELL_COUNT)
    const remote = makeRemote()
    remote.read.mockResolvedValueOnce(success(readResult('v1', content)))
    const mounted: OfficeSnapshot[] = []
    const fixture = panelFixture({
      meta: { absolutePath: ABSOLUTE_PATH, version: 'v1' },
      remote: remote.face,
      mount: (container, input) => recordingAdapter(mounted).mount(container, input),
    })
    const started = performance.now()
    const view = render(<OfficePanel {...fixture.props} />)
    await waitFor(() => { expect(mounted).toHaveLength(1) })
    const opened = performance.now() - started
    expect(mounted[0]).toBe(content)
    fireEvent.click(view.getByText(en.save))
    await waitFor(() => { expect(remote.save).toHaveBeenCalledOnce() })
    const saved = performance.now() - started
    console.info(
      `[office performance] cells=${CELL_COUNT} open=${opened.toFixed(1)}ms save=${saved.toFixed(1)}ms`,
    )
    expect(remote.save.mock.calls[0]?.[4]).toBe('v1')
    expect(opened).toBeLessThan(BUDGET_MS)
    expect(saved).toBeLessThan(BUDGET_MS)
  })
})

describe('SC-003 changed-snapshot visibility', () => {
  it('reflects a changed snapshot on the mounted surface inside the budget', async () => {
    vi.useFakeTimers()
    const first: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'first' } }
    const second: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'second' } }
    const remote = makeRemote()
    remote.read
      .mockResolvedValueOnce(success(readResult('v1', first)))
      .mockResolvedValueOnce(success(readResult('v2', second)))
    const mounted: OfficeSnapshot[] = []
    const fixture = panelFixture({
      meta: { absolutePath: ABSOLUTE_PATH, version: 'v1' },
      remote: remote.face,
      mount: (container, input) => recordingAdapter(mounted).mount(container, input),
    })
    const view = render(<OfficePanel {...fixture.props} />)
    await act(async () => { await vi.advanceTimersByTimeAsync(0) })
    expect(mounted).toHaveLength(1)
    expect(mounted[0]).toBe(first)
    const next = panelFixture({
      meta: { absolutePath: ABSOLUTE_PATH, version: 'v2' },
      remote: remote.face,
      mount: (container, input) => recordingAdapter(mounted).mount(container, input),
    })
    view.rerender(<OfficePanel {...next.props} />)
    await act(async () => { await vi.advanceTimersByTimeAsync(VISIBILITY_BUDGET_MS) })
    expect(mounted.at(-1)).toBe(second)
  })
})
