import { describe, expect, it } from 'vitest'
import { Context } from '@forwork/cordis'
import { DocumentId, SessionId } from '@forwork/session'
import type { Session, SessionEvent } from '@forwork/session'
import SessionStore from '@forwork/session'
import {
  documentIdFor,
  recordDocumentClose,
  recordDocumentMissing,
  recordDocumentOpen,
  recordDocumentRebind,
  recordDocumentSave,
} from '../src/index.ts'

const DOCUMENT = DocumentId('facts:report.docx')

async function mount(): Promise<{ ctx: Context; session: Session; events: SessionEvent[] }> {
  const ctx = new Context()
  await ctx.plugin(SessionStore)
  const session = ctx.sessions.create(SessionId('facts'))
  const events: SessionEvent[] = []
  ctx.on('session/event', (subject, event) => {
    if (subject === session) events.push(event)
  })
  return { ctx, session, events }
}

describe('document facts', () => {
  it('appends each fact once, in order, with its payload', async () => {
    const { ctx, session, events } = await mount()
    try {
      recordDocumentOpen(session, {
        documentId: DOCUMENT,
        path: 'report.docx',
        format: 'docx',
        contentVersion: 0,
      })
      recordDocumentSave(session, { documentId: DOCUMENT, actor: 'user', contentVersion: 1 })
      recordDocumentRebind(session, { documentId: DOCUMENT, fromPath: 'report.docx', toPath: 'final.docx' })
      recordDocumentMissing(session, { documentId: DOCUMENT, path: 'final.docx' })
      recordDocumentClose(session, { documentId: DOCUMENT, reason: 'closed' })

      expect(events.map(event => event.type)).toEqual([
        'document/open',
        'document/save',
        'document/rebind',
        'document/missing',
        'document/close',
      ])
      expect(events.map(event => event.seq)).toEqual([0, 1, 2, 3, 4])
      expect(events[0]?.data).toEqual({
        documentId: DOCUMENT,
        path: 'report.docx',
        format: 'docx',
        contentVersion: 0,
      })
      expect(events[1]?.data).toEqual({ documentId: DOCUMENT, actor: 'user', contentVersion: 1 })
      expect(events[2]?.data).toEqual({
        documentId: DOCUMENT,
        fromPath: 'report.docx',
        toPath: 'final.docx',
      })
      expect(events[3]?.data).toEqual({ documentId: DOCUMENT, path: 'final.docx' })
      expect(events[4]?.data).toEqual({ documentId: DOCUMENT, reason: 'closed' })
    } finally {
      await ctx.fiber.dispose()
    }
  })

  it('is append-only: a later fact leaves every earlier event untouched', async () => {
    const { ctx, session, events } = await mount()
    try {
      const open = recordDocumentOpen(session, {
        documentId: DOCUMENT,
        path: 'report.docx',
        format: 'docx',
        contentVersion: 0,
      })
      const snapshot = JSON.stringify(open)
      recordDocumentRebind(session, { documentId: DOCUMENT, fromPath: 'report.docx', toPath: 'final.docx' })

      expect(session.seq).toBe(2)
      expect(JSON.stringify(events[0])).toBe(snapshot)
      expect(events[0]?.type).toBe('document/open')
    } finally {
      await ctx.fiber.dispose()
    }
  })

  it('records the causing turn only when the save has one', async () => {
    const { ctx, session, events } = await mount()
    try {
      recordDocumentSave(session, { documentId: DOCUMENT, actor: 'agent', contentVersion: 1, turn: 2 })
      recordDocumentSave(session, { documentId: DOCUMENT, actor: 'user', contentVersion: 2 })

      expect(events[0]?.data).toEqual({ documentId: DOCUMENT, actor: 'agent', contentVersion: 1, turn: 2 })
      expect(events[1]?.data).toEqual({ documentId: DOCUMENT, actor: 'user', contentVersion: 2 })
      expect(events[1]?.data).not.toHaveProperty('turn')
    } finally {
      await ctx.fiber.dispose()
    }
  })

  it('builds the identity the client and tools share', () => {
    expect(documentIdFor(SessionId('facts'), 'report.docx')).toBe(DOCUMENT)
  })
})
