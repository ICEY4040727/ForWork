// @vitest-environment jsdom
/**
 * The converted-copy notice.
 *
 * A tab opened from a legacy source shows one localized banner naming that
 * source and stating that the original is unchanged; dismissing it hides the
 * banner for that converted copy, and a tab opened from an editable file shows
 * no banner at all.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent } from '@testing-library/react'
import { RemoteError } from '@forwork/typert-protocol'
import type { RemoteResult } from '@forwork/typert-protocol'
import type { OfficeDocumentReadResult, OfficeDocumentSaveResult, OfficeDocumentSnapshot } from '@forwork/office-exchange/types'
import { sessionFileAddress } from '@forwork/util-workspace-path'
import type { OfficeDocumentRemote } from '../src/client/useOfficeDocument.ts'
import { en } from '../src/client/locales.ts'
import { renderPanel } from './panel.tsx'

afterEach(cleanup)

const SNAPSHOT: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book' } }
const CONVERTED_ADDRESS = sessionFileAddress('s-1', '.dsh/office-tmp/convert-abc/report.docx')
const SOURCE = 'work/report.doc'

function noticeText(): string {
  return en.legacyNotice.replace('{source}', SOURCE)
}

function editableRemote(): OfficeDocumentRemote {
  return {
    read: async (): Promise<RemoteResult<OfficeDocumentReadResult>> => ({
      ok: true,
      value: { format: 'docx', snapshot: SNAPSHOT, unsupported: [], version: 'v1' },
    }),
    convertLegacy: vi.fn(),
    save: async (): Promise<RemoteResult<OfficeDocumentSaveResult>> => ({ ok: true, value: { version: 'v2' } }),
    writeDraft: vi.fn(),
    readDraft: async () => ({ ok: true, value: null }),
    clearDraft: async () => ({ ok: true, value: undefined }),
  }
}

describe('converted-copy notice', () => {
  it('renders the notice for the converted copy and hides it on dismiss', () => {
    const view = renderPanel({ address: CONVERTED_ADDRESS, remote: editableRemote() })
    act(() => { view.fixture.notices.remember(CONVERTED_ADDRESS, SOURCE) })
    expect(view.getByText(noticeText())).toBeDefined()
    fireEvent.click(view.getByText(en.dismiss))
    expect(view.fixture.notices.source.getSnapshot()[CONVERTED_ADDRESS]).toEqual({
      source: SOURCE,
      dismissed: true,
    })
    expect(view.queryByText(noticeText())).toBeNull()
  })

  it('keeps a dismissed notice hidden when the body mounts again', () => {
    const view = renderPanel({ address: CONVERTED_ADDRESS, remote: editableRemote() })
    act(() => {
      view.fixture.notices.remember(CONVERTED_ADDRESS, SOURCE)
      view.fixture.notices.dismiss(CONVERTED_ADDRESS)
    })
    expect(view.queryByText(noticeText())).toBeNull()
    expect(view.queryByText(en.dismiss)).toBeNull()
  })

  it('shows no notice for a file that was not converted', () => {
    const view = renderPanel({
      address: sessionFileAddress('s-1', 'work/budget.docx'),
      remote: editableRemote(),
    })
    expect(view.queryByText(en.dismiss)).toBeNull()
    expect(view.queryByText(noticeText())).toBeNull()
  })
})

describe('notice store', () => {
  it('ignores an unknown address and drops one record at a time', () => {
    const view = renderPanel({ address: CONVERTED_ADDRESS, remote: editableRemote() })
    const { notices } = view.fixture
    notices.dismiss('dsh-resource://file/session/s-1/nowhere.docx')
    notices.remember(CONVERTED_ADDRESS, SOURCE)
    notices.remember('dsh-resource://file/session/s-1/other.docx', 'work/other.xls')
    notices.forget('dsh-resource://file/session/s-1/nowhere.docx')
    notices.forget(CONVERTED_ADDRESS)
    expect(notices.source.getSnapshot()).toEqual({
      'dsh-resource://file/session/s-1/other.docx': { source: 'work/other.xls', dismissed: false },
    })
  })
})

describe('notice with a failed read', () => {
  it('keeps the tab usable without a notice when the file cannot be edited', async () => {
    const remote: OfficeDocumentRemote = {
      read: async () => ({
        ok: false,
        error: new RemoteError('office-document/unsupported-format', 'binary', { path: 'a.docx' }),
      }),
      convertLegacy: vi.fn(),
      save: vi.fn(),
      writeDraft: vi.fn(),
      readDraft: async () => ({ ok: true, value: null }),
      clearDraft: vi.fn(),
    }
    const view = renderPanel({ address: CONVERTED_ADDRESS, remote })
    expect(await view.findByText(en['error.unsupportedFormat'])).toBeDefined()
    expect(view.queryByText(en.dismiss)).toBeNull()
  })
})
