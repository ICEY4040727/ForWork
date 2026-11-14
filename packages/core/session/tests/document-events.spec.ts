import { describe, expect, expectTypeOf, it } from 'vitest'
import { snapshotJsonValue } from '@deepseek-ai/dsh-util-values'
import {
  DocumentId,
  KNOWN_SESSION_EVENT_TYPES,
  Session,
  SessionId,
  SessionSeq,
  snapshotSessionEvent,
} from '@deepseek-ai/dsh-session'
import type {
  SessionEvent,
  SessionEventMap,
} from '@deepseek-ai/dsh-session'

/** The five document facts emitted as one Session's open-document history. */
function documentHistory(): readonly [
  SessionEvent<'document/open'>,
  SessionEvent<'document/edit'>,
  SessionEvent<'document/rebind'>,
  SessionEvent<'document/missing'>,
  SessionEvent<'document/close'>,
] {
  const session = Session.create(SessionId('document-events'))
  const documentId = DocumentId('document-1')
  const open = session.append('document/open', {
    documentId,
    path: 'C:/documents/report.xlsx',
    format: 'xlsx',
    contentVersion: 0,
    turn: 1,
  })
  const edit = session.append('document/edit', {
    documentId,
    actor: 'agent',
    fromVersion: 0,
    toVersion: 1,
    operation: [{ kind: 'cell', sheetId: 'sheet-1', row: 1, column: 2, value: 42 }],
    inverse: [{ kind: 'cell', sheetId: 'sheet-1', row: 1, column: 2, value: null }],
    turn: 1,
  })
  const rebind = session.append('document/rebind', {
    documentId,
    fromPath: 'C:/documents/report.xlsx',
    toPath: 'C:/documents/archive/report.xlsx',
  })
  const missing = session.append('document/missing', {
    documentId,
    path: 'C:/documents/archive/report.xlsx',
  })
  const close = session.append('document/close', {
    documentId,
    reason: 'user',
  })
  return [open, edit, rebind, missing, close]
}

describe('document Session events', () => {
  it('brands a document id without changing the value', () => {
    expect(DocumentId('document-1')).toBe('document-1')
  })

  it('keeps every document payload lossless JSON', () => {
    const [open, edit, rebind, missing, close] = documentHistory()

    expect(snapshotJsonValue(open.data)).toEqual(open.data)
    expect(snapshotJsonValue(edit.data)).toEqual(edit.data)
    expect(snapshotJsonValue(rebind.data)).toEqual(rebind.data)
    expect(snapshotJsonValue(missing.data)).toEqual(missing.data)
    expect(snapshotJsonValue(close.data)).toEqual(close.data)
    expect(edit.data.operation).toEqual([{ kind: 'cell', sheetId: 'sheet-1', row: 1, column: 2, value: 42 }])
    expect(edit.data.inverse).toEqual([{ kind: 'cell', sheetId: 'sheet-1', row: 1, column: 2, value: null }])
    expect(edit.data.toVersion).toBe(edit.data.fromVersion + 1)
  })

  it('restores the document history through the package event snapshot', () => {
    const history = documentHistory()
    const restored = Session.create(
      SessionId('document-events-restored'),
      history.map(event => snapshotSessionEvent(event)),
    )

    // Restoring a seed appends the resume marker after the five document facts.
    expect(restored.seq).toBe(history.length + 1)
    expect(restored.eventAt(SessionSeq(0))).toEqual(history[0])
    expect(restored.eventAt(SessionSeq(history.length - 1))).toEqual(history[history.length - 1])
  })

  it('declares every document fact required and inside the generated vocabulary', () => {
    const history = documentHistory()
    for (const event of history) expect(event.ignorable).toBeUndefined()
    for (const type of ['document/open', 'document/edit', 'document/rebind', 'document/missing', 'document/close']) {
      expect(KNOWN_SESSION_EVENT_TYPES.has(type)).toBe(true)
    }
  })

  it('document/edit has no raw unit snapshot field', () => {
    expectTypeOf<SessionEventMap['document/edit']>().toHaveProperty('operation')
    expectTypeOf<SessionEventMap['document/edit']>().toHaveProperty('inverse')
    expectTypeOf<SessionEventMap['document/edit']>().not.toHaveProperty('snapshot')
  })

  it('retains an unknown ignorable document event without interpreting it', () => {
    const future = {
      type: 'document/future',
      seq: SessionSeq(0),
      time: 1,
      data: { documentId: 'document-1' },
      ignorable: true,
    }
    const restored = Session.create(SessionId('document-events-forward'), [future] as readonly object[] as readonly SessionEvent[])

    expect(restored.eventAt(SessionSeq(0))).toEqual(future)
    expect(restored.surface.nodes).toEqual([])
    expect(restored.deriveMessages()).toEqual([])
  })
})
