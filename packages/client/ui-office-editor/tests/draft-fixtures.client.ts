/** Shared fake Remote and draft values for the recovery-draft specs. */
import { vi } from 'vitest'
import type { Mock } from 'vitest'
import { brandString } from '@forwork/brand'
import type { SessionId } from '@forwork/session/types'
import type { RemoteResult } from '@forwork/typert-protocol'
import type {
  OfficeDocumentDraft,
  OfficeDocumentReadResult,
  OfficeDocumentSaveResult,
  OfficeDocumentSnapshot,
} from '@forwork/office-exchange/types'
import type { OfficeDocumentRemote } from '../src/client/useOfficeDocument.ts'

/** The Session every fixture addresses. */
export const SESSION_ID = brandString<SessionId>('s-1')
/** The workspace path every fixture opens. */
export const PATH = '/work/budget.xlsx'
/** The document identity the hook derives for the fixture path. */
export const DOCUMENT_ID = `${SESSION_ID}:${PATH}`
/** A lifetime that never aborts. */
export const SIGNAL = new AbortController().signal
/** The snapshot the file holds. */
export const snapshot: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book' } }
/** The snapshot the draft holds. */
export const edited: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book', name: 'edited' } }

/** Wrap a value as a successful Remote result. */
export function success<T>(value: T): RemoteResult<T> {
  return { ok: true, value }
}

/** The successful result of a Remote call that returns nothing. */
export function successVoid(): RemoteResult<void> {
  return { ok: true, value: undefined }
}

/** One stored draft captured from the given file version. */
export function draftAt(contentVersion: string, captured: OfficeDocumentSnapshot = edited): OfficeDocumentDraft {
  return { documentId: DOCUMENT_ID, path: PATH, contentVersion, snapshot: captured, savedAt: '2026-11-14T00:00:00.000Z' }
}

/** The read result for the fixture file. */
export function readResult(version = 'v1'): OfficeDocumentReadResult {
  return { format: 'xlsx', snapshot, unsupported: [], version }
}

type ReadSignature = OfficeDocumentRemote['read']
type WriteDraftSignature = OfficeDocumentRemote['writeDraft']
type ReadDraftSignature = OfficeDocumentRemote['readDraft']
type ClearDraftSignature = OfficeDocumentRemote['clearDraft']
type SaveSignature = OfficeDocumentRemote['save']

/** A fully mocked `officeDocument` Remote face. */
export interface DraftRemote {
  /** The read call. */
  readonly read: Mock<ReadSignature>
  /** The draft write call. */
  readonly writeDraft: Mock<WriteDraftSignature>
  /** The draft read call. */
  readonly readDraft: Mock<ReadDraftSignature>
  /** The draft removal call. */
  readonly clearDraft: Mock<ClearDraftSignature>
  /** The save call. */
  readonly save: Mock<SaveSignature>
  /** The Remote face the hook receives. */
  readonly face: OfficeDocumentRemote
}

/** Build a Remote face that reads the fixture file and has no stored draft. */
export function makeDraftRemote(): DraftRemote {
  const read = vi.fn<ReadSignature>().mockResolvedValue(success(readResult()))
  const writeDraft = vi.fn<WriteDraftSignature>().mockResolvedValue(success({ savedAt: 't' }))
  const readDraft = vi.fn<ReadDraftSignature>().mockResolvedValue(success(null))
  const clearDraft = vi.fn<ClearDraftSignature>().mockResolvedValue(successVoid())
  const save = vi.fn<SaveSignature>().mockResolvedValue(success<OfficeDocumentSaveResult>({ version: 'v2' }))
  return {
    read,
    writeDraft,
    readDraft,
    clearDraft,
    save,
    face: { read, convertLegacy: vi.fn(), save, writeDraft, readDraft, clearDraft },
  }
}
