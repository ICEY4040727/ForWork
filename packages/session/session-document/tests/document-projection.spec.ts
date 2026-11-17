import { describe, expect, it } from 'vitest'
import { Context } from '@forwork/cordis'
import { DocumentId, SessionId, SessionSeq } from '@forwork/session'
import type { DocumentEditOperation, SessionEvent } from '@forwork/session'
import SessionStore from '@forwork/session'
import {
  documentProjectionDefinition,
  foldDocumentProjection,
  openDocumentAtPath,
  recordDocumentClose,
  recordDocumentMissing,
  recordDocumentOpen,
  recordDocumentRebind,
  recordDocumentSave,
} from '../src/index.ts'
import type { DocumentProjectionState } from '../src/index.ts'

const FIRST = DocumentId('projection:first.xlsx')
const SECOND = DocumentId('projection:second.docx')

const OPERATION: DocumentEditOperation = [
  { kind: 'cell', sheetId: 'sheet-1', row: 0, column: 0, value: 1 },
]
const INVERSE: DocumentEditOperation = [
  { kind: 'cell', sheetId: 'sheet-1', row: 0, column: 0, value: null },
]

function empty(): DocumentProjectionState {
  return documentProjectionDefinition.init()
}

function fold(events: readonly SessionEvent[]): DocumentProjectionState {
  return events.reduce(foldDocumentProjection, empty())
}

describe('foldDocumentProjection', () => {
  it('inserts documents in open order and keeps that order across later facts', async () => {
    const ctx = new Context()
    await ctx.plugin(SessionStore)
    try {
      const session = ctx.sessions.create(SessionId('projection-order'))
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
      const saveFirst = recordDocumentSave(session, { documentId: FIRST, actor: 'user', contentVersion: 1 })
      const editFirst = session.append('document/edit', {
        documentId: FIRST,
        actor: 'agent',
        fromVersion: 1,
        toVersion: 2,
        operation: OPERATION,
        inverse: INVERSE,
      })
      const rebindSecond = recordDocumentRebind(session, {
        documentId: SECOND,
        fromPath: 'second.docx',
        toPath: 'renamed.docx',
      })

      const state = fold([openFirst, openSecond, saveFirst, editFirst, rebindSecond])

      expect(state.order).toEqual([FIRST, SECOND])
      expect(state.documents[FIRST]).toMatchObject({
        documentId: FIRST,
        path: 'first.xlsx',
        format: 'xlsx',
        contentVersion: 2,
        status: 'open',
      })
      expect(state.documents[SECOND]).toMatchObject({
        documentId: SECOND,
        path: 'renamed.docx',
        format: 'docx',
        status: 'open',
      })
    } finally {
      await ctx.fiber.dispose()
    }
  })

  it('marks a document missing without closing it and closes only the named document', async () => {
    const ctx = new Context()
    await ctx.plugin(SessionStore)
    try {
      const session = ctx.sessions.create(SessionId('projection-status'))
      const openFirst = recordDocumentOpen(session, {
        documentId: FIRST,
        path: 'first.xlsx',
        format: 'xlsx',
        contentVersion: 3,
      })
      const openSecond = recordDocumentOpen(session, {
        documentId: SECOND,
        path: 'second.docx',
        format: 'docx',
        contentVersion: 0,
      })
      const missingFirst = recordDocumentMissing(session, { documentId: FIRST, path: 'first.xlsx' })
      const closedFirst = recordDocumentClose(session, { documentId: FIRST, reason: 'closed' })

      const missingState = fold([openFirst, openSecond, missingFirst])
      expect(missingState.order).toEqual([FIRST, SECOND])
      expect(missingState.documents[FIRST]?.status).toBe('missing')
      expect(missingState.documents[FIRST]?.contentVersion).toBe(3)

      const closedState = fold([openFirst, openSecond, missingFirst, closedFirst])
      expect(closedState.order).toEqual([SECOND])
      expect(closedState.documents[FIRST]).toBeUndefined()
      expect(closedState.documents[SECOND]?.status).toBe('open')
    } finally {
      await ctx.fiber.dispose()
    }
  })

  it('keeps a reopened document in place and re-establishes its version epoch', async () => {
    const ctx = new Context()
    await ctx.plugin(SessionStore)
    try {
      const session = ctx.sessions.create(SessionId('projection-reopen'))
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
      const saveFirst = recordDocumentSave(session, { documentId: FIRST, actor: 'user', contentVersion: 1 })
      const reopenFirst = recordDocumentOpen(session, {
        documentId: FIRST,
        path: 'first.xlsx',
        format: 'xlsx',
        contentVersion: 0,
      })

      const state = fold([openFirst, openSecond, saveFirst, reopenFirst])

      expect(state.order).toEqual([FIRST, SECOND])
      expect(state.documents[FIRST]?.contentVersion).toBe(0)
    } finally {
      await ctx.fiber.dispose()
    }
  })

  it('ignores a fact for a document the open set does not hold', async () => {
    const ctx = new Context()
    await ctx.plugin(SessionStore)
    try {
      const session = ctx.sessions.create(SessionId('projection-unknown'))
      const openFirst = recordDocumentOpen(session, {
        documentId: FIRST,
        path: 'first.xlsx',
        format: 'xlsx',
        contentVersion: 0,
      })
      const state = fold([openFirst])
      const saveOrphan = recordDocumentSave(session, { documentId: SECOND, actor: 'user', contentVersion: 1 })
      const editOrphan = session.append('document/edit', {
        documentId: SECOND,
        actor: 'agent',
        fromVersion: 0,
        toVersion: 1,
        operation: OPERATION,
        inverse: INVERSE,
      })
      const rebindOrphan = recordDocumentRebind(session, {
        documentId: SECOND,
        fromPath: 'second.docx',
        toPath: 'renamed.docx',
      })

      const missingOrphan = recordDocumentMissing(session, { documentId: SECOND, path: 'second.docx' })

      expect(foldDocumentProjection(state, saveOrphan)).toBe(state)
      expect(foldDocumentProjection(state, editOrphan)).toBe(state)
      expect(foldDocumentProjection(state, rebindOrphan)).toBe(state)
      expect(foldDocumentProjection(state, missingOrphan)).toBe(state)
    } finally {
      await ctx.fiber.dispose()
    }
  })

  it('returns the same state reference for an unrelated event or an unchanged fact', async () => {
    const ctx = new Context()
    await ctx.plugin(SessionStore)
    try {
      const session = ctx.sessions.create(SessionId('projection-identity'))
      const open = recordDocumentOpen(session, {
        documentId: FIRST,
        path: 'first.xlsx',
        format: 'xlsx',
        contentVersion: 0,
      })
      const state = fold([open])
      const unrelated = session.append('turn/start', { turn: 1 })
      const repeated = recordDocumentMissing(session, { documentId: FIRST, path: 'first.xlsx' })

      expect(foldDocumentProjection(state, open)).toBe(state)
      expect(foldDocumentProjection(state, unrelated)).toBe(state)
      const afterMissing = foldDocumentProjection(state, repeated)
      expect(foldDocumentProjection(afterMissing, repeated)).toBe(afterMissing)
      const unknownClose = foldDocumentProjection(afterMissing, closedEvent())
      expect(unknownClose).toBe(afterMissing)
    } finally {
      await ctx.fiber.dispose()
    }
  })

  it('serializes and restores through the wire view schema', async () => {
    const ctx = new Context()
    await ctx.plugin(SessionStore)
    try {
      const session = ctx.sessions.create(SessionId('projection-wire'))
      const open = recordDocumentOpen(session, {
        documentId: FIRST,
        path: 'first.xlsx',
        format: 'xlsx',
        contentVersion: 0,
      })
      const state = fold([open])
      const view = documentProjectionDefinition.wire.view(state)
      const restored = documentProjectionDefinition.wire.viewSchema.parse(JSON.parse(JSON.stringify(view)))

      expect(restored).toEqual(state)
      expect(restored.documents[FIRST]?.documentId).toBe(FIRST)
    } finally {
      await ctx.fiber.dispose()
    }
  })

  it('rejects a state whose order and entries disagree', () => {
    const entry = {
      documentId: FIRST,
      path: 'first.xlsx',
      format: 'xlsx',
      contentVersion: 0,
      status: 'open',
      unsupported: [],
    } as const
    expect(() => documentProjectionDefinition.stateSchema.parse({
      order: [FIRST],
      documents: { [FIRST]: entry, [SECOND]: { ...entry, documentId: SECOND } },
    })).toThrow()
    expect(() => documentProjectionDefinition.stateSchema.parse({
      order: [FIRST, SECOND],
      documents: { [FIRST]: entry },
    })).toThrow()
    expect(() => documentProjectionDefinition.stateSchema.parse({
      order: [FIRST, FIRST],
      documents: { [FIRST]: entry },
    })).toThrow()
  })
})

/** A close fact for a document the state never opened. */
function closedEvent(): SessionEvent<'document/close'> {
  return {
    type: 'document/close',
    seq: SessionSeq(0),
    time: 0,
    data: { documentId: SECOND, reason: 'closed' },
  }
}

describe('openDocumentAtPath', () => {
  it('resolves a rebound document at its new locator under the original identity', () => {
    const state = fold([
      {
        type: 'document/open',
        seq: SessionSeq(0),
        time: 0,
        data: { documentId: FIRST, path: 'first.xlsx', format: 'xlsx', contentVersion: 2 },
      },
      {
        type: 'document/rebind',
        seq: SessionSeq(1),
        time: 1,
        data: { documentId: FIRST, fromPath: 'first.xlsx', toPath: 'renamed.xlsx' },
      },
    ])

    expect(openDocumentAtPath(state, 'renamed.xlsx')).toMatchObject({
      documentId: FIRST,
      path: 'renamed.xlsx',
      contentVersion: 2,
    })
    expect(openDocumentAtPath(state, 'first.xlsx')).toBeUndefined()
  })

  it('reports no entry for a path no open document carries', () => {
    expect(openDocumentAtPath(empty(), 'first.xlsx')).toBeUndefined()
  })
})
