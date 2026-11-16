import { describe, expect, it } from 'vitest'
import { Context } from '@forwork/cordis'
import { DocumentId, SessionId } from '@forwork/session'
import SessionStore from '@forwork/session'
import {
  documentProjectionDefinition,
  foldDocumentProjection,
  recordDocumentClose,
  recordDocumentMissing,
  recordDocumentOpen,
  restoreSet,
} from '../src/index.ts'
import type { DocumentProjectionEntry, DocumentProjectionState } from '../src/index.ts'

const FIRST = DocumentId('restore:first.xlsx')
const SECOND = DocumentId('restore:second.docx')
const THIRD = DocumentId('restore:third.pptx')

function openEntry(
  documentId: DocumentProjectionEntry['documentId'],
  path: string,
  format: DocumentProjectionEntry['format'],
  status: DocumentProjectionEntry['status'] = 'open',
): DocumentProjectionEntry {
  return { documentId, path, format, contentVersion: 2, status, unsupported: [] }
}

function stateOf(entries: readonly DocumentProjectionEntry[]): DocumentProjectionState {
  const documents: Record<string, DocumentProjectionEntry> = {}
  for (const entry of entries) documents[entry.documentId] = entry
  return { order: entries.map(entry => entry.documentId), documents }
}

describe('restoreSet', () => {
  it('returns every open document in the projection order', () => {
    const state = stateOf([
      openEntry(SECOND, 'second.docx', 'docx'),
      openEntry(FIRST, 'first.xlsx', 'xlsx'),
      openEntry(THIRD, 'third.pptx', 'pptx'),
    ])

    expect(restoreSet(state).map(entry => entry.path)).toEqual(['second.docx', 'first.xlsx', 'third.pptx'])
  })

  it('keeps the identity, path, and format of each entry', () => {
    const state = stateOf([openEntry(FIRST, 'renamed.xlsx', 'xlsx')])

    expect(restoreSet(state)).toEqual([{
      documentId: FIRST,
      path: 'renamed.xlsx',
      format: 'xlsx',
      status: 'open',
    }])
  })

  it('retains a missing document with its status', () => {
    const state = stateOf([
      openEntry(FIRST, 'first.xlsx', 'xlsx', 'missing'),
      openEntry(SECOND, 'second.docx', 'docx'),
    ])

    expect(restoreSet(state)).toEqual([
      { documentId: FIRST, path: 'first.xlsx', format: 'xlsx', status: 'missing' },
      { documentId: SECOND, path: 'second.docx', format: 'docx', status: 'open' },
    ])
  })

  it('returns nothing for an empty open set', () => {
    expect(restoreSet(documentProjectionDefinition.init())).toEqual([])
  })

  it('skips an order entry with no document record', () => {
    const state = stateOf([openEntry(FIRST, 'first.xlsx', 'xlsx')])

    expect(restoreSet({ ...state, order: [SECOND, FIRST] })).toEqual([{
      documentId: FIRST,
      path: 'first.xlsx',
      format: 'xlsx',
      status: 'open',
    }])
  })

  it('omits a closed document and reports a missing one that stayed open', async () => {
    const ctx = new Context()
    await ctx.plugin(SessionStore)
    try {
      const session = ctx.sessions.create(SessionId('restore-facts'))
      const openFirst = recordDocumentOpen(session, {
        documentId: FIRST,
        path: 'first.xlsx',
        format: 'xlsx',
        contentVersion: 0,
      })
      const openSecond = recordDocumentOpen(session, {
        documentId: SECOND,
        path: 'second.docx',
        format: 'docx',
        contentVersion: 0,
      })
      const missingFirst = recordDocumentMissing(session, { documentId: FIRST, path: 'first.xlsx' })
      const closedSecond = recordDocumentClose(session, { documentId: SECOND, reason: 'closed' })

      const state = [openFirst, openSecond, missingFirst, closedSecond]
        .reduce(foldDocumentProjection, documentProjectionDefinition.init())

      expect(restoreSet(state)).toEqual([
        { documentId: FIRST, path: 'first.xlsx', format: 'xlsx', status: 'missing' },
      ])
    } finally {
      await ctx.fiber.dispose()
    }
  })
})
