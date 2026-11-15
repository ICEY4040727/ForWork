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
  DocumentEditStep,
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
    for (const type of ['document/open', 'document/edit', 'document/rebind', 'document/missing', 'document/close', 'document/command']) {
      expect(KNOWN_SESSION_EVENT_TYPES.has(type)).toBe(true)
    }
  })

  it('document/edit has no raw unit snapshot field', () => {
    expectTypeOf<SessionEventMap['document/edit']>().toHaveProperty('operation')
    expectTypeOf<SessionEventMap['document/edit']>().toHaveProperty('inverse')
    expectTypeOf<SessionEventMap['document/edit']>().not.toHaveProperty('snapshot')
  })

  it('accepts every edit-step kind in one operation and its inverse', () => {
    const session = Session.create(SessionId('document-steps'))
    const documentId = DocumentId('document-1')
    const operation: readonly DocumentEditStep[] = [
      { kind: 'cell', sheetId: 'sheet-1', row: 0, column: 1, value: 7 },
      { kind: 'paragraph', index: 0, text: 'Updated paragraph' },
      { kind: 'row', sheetId: 'sheet-1', row: 2, count: 1, action: 'insert' },
      { kind: 'column', sheetId: 'sheet-1', column: 3, count: 2, action: 'delete' },
      { kind: 'formula', sheetId: 'sheet-1', row: 1, column: 2, formula: '=SUM(A1:A2)' },
      {
        kind: 'style',
        sheetId: 'sheet-1',
        startRow: 0,
        startColumn: 0,
        endRow: 1,
        endColumn: 1,
        style: { bl: 1 },
      },
    ]
    const edit = session.append('document/edit', {
      documentId,
      actor: 'agent',
      fromVersion: 0,
      toVersion: 1,
      operation,
      inverse: operation,
    })

    expect(edit.data.operation).toEqual(operation)
    expect(snapshotJsonValue(edit.data)).toEqual(edit.data)
  })

  it('records a correlated read or edit command without a snapshot', () => {
    const session = Session.create(SessionId('document-command'))
    const documentId = DocumentId('document-1')
    const read = session.append('document/command', {
      requestId: 'request-read',
      documentId,
      operation: { kind: 'read' },
      fromVersion: 3,
    })
    const edit = session.append('document/command', {
      requestId: 'request-edit',
      documentId,
      operation: {
        kind: 'edit',
        steps: [{ kind: 'cell', sheetId: 'sheet-1', row: 0, column: 0, value: 'x' }],
      },
      fromVersion: 4,
    })

    expect(read.data).toEqual({
      requestId: 'request-read',
      documentId,
      operation: { kind: 'read' },
      fromVersion: 3,
    })
    expect(edit.data.operation).toEqual({
      kind: 'edit',
      steps: [{ kind: 'cell', sheetId: 'sheet-1', row: 0, column: 0, value: 'x' }],
    })
    expect(snapshotJsonValue(edit.data)).toEqual(edit.data)
    expectTypeOf<SessionEventMap['document/command']>().not.toHaveProperty('snapshot')
    expect(KNOWN_SESSION_EVENT_TYPES.has('document/command')).toBe(true)
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
