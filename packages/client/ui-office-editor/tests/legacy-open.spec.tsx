// @vitest-environment jsdom
/**
 * Opening a legacy `.doc`, `.xls`, or `.ppt` source converts it and opens the
 * copy in a new tab.
 *
 * The converting body never reads the binary; it asks the `officeDocument`
 * Remote for a converted copy, addresses that copy through the workspace-path
 * helper relative to the workspace root the opened file implies, and hands the
 * new address to the tab's own `openResource` replacement while the original
 * path survives in the plugin's notice store.
 */
import { afterEach, describe, expect, it, vi, type Mock } from 'vitest'
import { act, cleanup, render, waitFor } from '@testing-library/react'
import { RemoteError } from '@forwork/typert-protocol'
import type { RemoteResult } from '@forwork/typert-protocol'
import { sessionFileAddress } from '@forwork/util-workspace-path'
import type { OfficeDocumentReadResult, OfficeDocumentSaveResult } from '@forwork/office-exchange/types'
import type { OfficeDocumentRemote } from '../src/client/useOfficeDocument.ts'
import { convertedAddress, legacyFormatOf, workspaceRootOf } from '../src/client/target.ts'
import { en } from '../src/client/locales.ts'
import { OfficePanel } from '../src/client/OfficePanel.tsx'
import { SESSION_ID, TAB_ID, panelFixture, renderPanel } from './panel.tsx'

afterEach(cleanup)

const SOURCE_PATH = 'work/report.doc'
const SOURCE_ADDRESS = sessionFileAddress('s-1', SOURCE_PATH)
const ABSOLUTE_SOURCE = '/ws/work/report.doc'
const CONVERTED = '/ws/.dsh/office-tmp/convert-abc/report.docx'
const CONVERTED_ADDRESS = sessionFileAddress('s-1', '.dsh/office-tmp/convert-abc/report.docx')

type ReadSignature = OfficeDocumentRemote['read']
type ConvertSignature = OfficeDocumentRemote['convertLegacy']

function success<T>(value: T): RemoteResult<T> {
  return { ok: true, value }
}

function remoteWith(conversion: () => Promise<RemoteResult<string>>): {
  read: Mock<ReadSignature>
  convert: Mock<ConvertSignature>
  face: OfficeDocumentRemote
} {
  const read = vi.fn<ReadSignature>(async (): Promise<RemoteResult<OfficeDocumentReadResult>> => {
    throw new Error('a legacy source must never be read')
  })
  const convert = vi.fn<ConvertSignature>(conversion)
  const save = vi.fn<OfficeDocumentRemote['save']>(async (): Promise<RemoteResult<OfficeDocumentSaveResult>> => success({ version: 'v2' }))
  return {
    read,
    convert,
    face: {
      read,
      convertLegacy: convert,
      save,
      writeDraft: vi.fn(),
      readDraft: vi.fn(async (): Promise<RemoteResult<null>> => success(null)),
      clearDraft: vi.fn(async (): Promise<RemoteResult<void>> => success(undefined)),
    },
  }
}

function meta(absolutePath = ABSOLUTE_SOURCE) {
  return { absolutePath, version: 'v1' }
}

describe('legacy path helpers', () => {
  it('names only the legacy formats a conversion targets', () => {
    expect(legacyFormatOf('report.doc')).toBe('doc')
    expect(legacyFormatOf('budget.XLS')).toBe('xls')
    expect(legacyFormatOf('deck.Ppt')).toBe('ppt')
    expect(legacyFormatOf('report.docx')).toBeUndefined()
    expect(legacyFormatOf('report')).toBeUndefined()
  })

  it('derives the workspace root only from a relative address and its absolute path', () => {
    expect(workspaceRootOf(SOURCE_PATH, undefined)).toBeUndefined()
    expect(workspaceRootOf('/ws/work/report.doc', ABSOLUTE_SOURCE)).toBeUndefined()
    expect(workspaceRootOf(SOURCE_PATH, '/ws/other/report.doc')).toBeUndefined()
    expect(workspaceRootOf(SOURCE_PATH, ABSOLUTE_SOURCE)).toBe('/ws')
    expect(workspaceRootOf('work/report.doc', 'C:\\ws\\work\\report.doc')).toBe('C:/ws')
    expect(workspaceRootOf('a.doc', '/a.doc')).toBe('/')
  })

  it('addresses a copy inside the workspace and refuses one outside it', () => {
    expect(convertedAddress(SESSION_ID, CONVERTED, '/ws')).toBe(CONVERTED_ADDRESS)
    expect(convertedAddress(SESSION_ID, CONVERTED, '/ws/')).toBe(CONVERTED_ADDRESS)
    expect(convertedAddress(SESSION_ID, CONVERTED, undefined))
      .toBe(sessionFileAddress('s-1', CONVERTED))
    expect(convertedAddress(SESSION_ID, '/elsewhere/report.docx', '/ws')).toBeUndefined()
  })
})

describe('legacy open', () => {
  it('waits to convert a dirty unit so a rebind never drops live edits', async () => {
    const remote = remoteWith(async () => success(CONVERTED))
    const fixture = panelFixture({ address: SOURCE_ADDRESS, meta: meta(), remote: remote.face })
    fixture.changes.markDirty(TAB_ID, true)

    render(<OfficePanel {...fixture.props} />)
    await act(async () => {})
    expect(remote.convert).not.toHaveBeenCalled()

    act(() => { fixture.changes.markDirty(TAB_ID, false) })
    await waitFor(() => { expect(remote.convert).toHaveBeenCalledOnce() })
  })

  it('converts the source and opens the copy in a new tab with the source recorded', async () => {
    const remote = remoteWith(async () => success(CONVERTED))
    const view = renderPanel({ address: SOURCE_ADDRESS, meta: meta(), remote: remote.face })
    await waitFor(() => { expect(remote.convert).toHaveBeenCalledOnce() })
    expect(remote.convert.mock.calls[0]?.[0]).toBe(SESSION_ID)
    expect(remote.convert.mock.calls[0]?.[1]).toBe(SOURCE_PATH)
    expect(remote.read).not.toHaveBeenCalled()
    await waitFor(() => { expect(view.fixture.openResource).toHaveBeenCalledOnce() })
    expect(view.fixture.openResource.mock.calls[0]?.[0]).toBe(CONVERTED_ADDRESS)
    expect(view.fixture.openResource.mock.calls[0]?.[1]).toEqual({ replaceTab: true })
    expect(view.fixture.notices.source.getSnapshot()[CONVERTED_ADDRESS]).toEqual({
      source: SOURCE_PATH,
      dismissed: false,
    })
    expect(view.getByText(en.converting)).toBeDefined()
  })

  it('opens the copy without a Session address path change when the root is unknown', async () => {
    const remote = remoteWith(async () => success(CONVERTED))
    const view = renderPanel({ address: sessionFileAddress('s-1', ABSOLUTE_SOURCE), remote: remote.face })
    await waitFor(() => { expect(view.fixture.openResource).toHaveBeenCalledOnce() })
    expect(view.fixture.openResource.mock.calls[0]?.[0]).toBe(sessionFileAddress('s-1', CONVERTED))
  })

  it('reports a refused conversion and opens nothing', async () => {
    const remote = remoteWith(async () => ({
      ok: false,
      error: new RemoteError('office-document/conversion-failed', 'no engine', { path: SOURCE_PATH }),
    }))
    const view = renderPanel({ address: SOURCE_ADDRESS, meta: meta(), remote: remote.face })
    expect(await view.findByText(en['error.conversionFailed'])).toBeDefined()
    expect(view.fixture.openResource).not.toHaveBeenCalled()
  })

  it('reports a copy outside the Session workspace as a conversion failure', async () => {
    const remote = remoteWith(async () => success('/elsewhere/report.docx'))
    const view = renderPanel({ address: SOURCE_ADDRESS, meta: meta(), remote: remote.face })
    expect(await view.findByText(en['error.conversionFailed'])).toBeDefined()
    expect(view.fixture.openResource).not.toHaveBeenCalled()
    expect(view.fixture.notices.source.getSnapshot()).toEqual({})
  })

  it('reports a rejected conversion call', async () => {
    const remote = remoteWith(async () => { throw new Error('offline') })
    const view = renderPanel({ address: SOURCE_ADDRESS, meta: meta(), remote: remote.face })
    expect(await view.findByText(en['error.unknown'])).toBeDefined()
  })

  it('stringifies a conversion rejected without an Error', async () => {
    const remote = remoteWith(async () => { throw 'offline' })
    const view = renderPanel({ address: SOURCE_ADDRESS, meta: meta(), remote: remote.face })
    expect(await view.findByText(en['error.unknown'])).toBeDefined()
  })

  it('does not convert an editable file', async () => {
    const remote = remoteWith(async () => success(CONVERTED))
    const view = renderPanel({ address: sessionFileAddress('s-1', 'work/budget.xlsx'), meta: meta('/ws/work/budget.xlsx'), remote: remote.face })
    await waitFor(() => { expect(remote.read).toHaveBeenCalledOnce() })
    expect(remote.convert).not.toHaveBeenCalled()
    expect(view.fixture.openResource).not.toHaveBeenCalled()
  })

  it('ignores a conversion that settles after the body unmounts', async () => {
    let settle: (value: RemoteResult<string>) => void = () => {}
    const convert = vi.fn<ConvertSignature>(() => new Promise((done) => { settle = done }))
    const remote: OfficeDocumentRemote = {
      read: vi.fn(),
      convertLegacy: convert,
      save: vi.fn(),
      writeDraft: vi.fn(),
      readDraft: vi.fn(async (): Promise<RemoteResult<null>> => success(null)),
      clearDraft: vi.fn(),
    }
    const view = renderPanel({ address: SOURCE_ADDRESS, meta: meta(), remote })
    await waitFor(() => { expect(convert).toHaveBeenCalledOnce() })
    view.unmount()
    settle(success(CONVERTED))
    await Promise.resolve()
    await Promise.resolve()
    expect(view.fixture.openResource).not.toHaveBeenCalled()
  })

  it('ignores a rejection that arrives after the body unmounts', async () => {
    let fail: (reason: unknown) => void = () => {}
    const convert = vi.fn<ConvertSignature>(() => new Promise((_done, reject) => { fail = reject }))
    const remote: OfficeDocumentRemote = {
      read: vi.fn(),
      convertLegacy: convert,
      save: vi.fn(),
      writeDraft: vi.fn(),
      readDraft: vi.fn(async (): Promise<RemoteResult<null>> => success(null)),
      clearDraft: vi.fn(),
    }
    const view = renderPanel({ address: SOURCE_ADDRESS, meta: meta(), remote })
    await waitFor(() => { expect(convert).toHaveBeenCalledOnce() })
    view.unmount()
    fail(new Error('late'))
    await Promise.resolve()
    await Promise.resolve()
    expect(view.fixture.openResource).not.toHaveBeenCalled()
  })
})
