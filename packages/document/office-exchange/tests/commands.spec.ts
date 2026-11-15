/** The correlated command broker appends one request, resolves one answer, and bounds the wait. */
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Context } from '@deepseek-ai/cordis'
import { LocalFileSystem } from '@deepseek-ai/dsh-fs-local'
import { DocumentId, Session, SessionId } from '@deepseek-ai/dsh-session'
import type { SessionEvent } from '@deepseek-ai/dsh-session'
import SessionStore from '@deepseek-ai/dsh-session'
import type { OfficeCommandResult, OfficeDocumentSnapshot } from '../src/index.ts'
import LocalOfficeExchange, { OfficeDocumentRemote } from '../src/index.ts'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

const SIGNAL = new AbortController().signal
const DOCUMENT = DocumentId('command-test:report.xlsx')

/** A snapshot envelope sufficient for the wire result; the broker never inspects its content. */
const SNAPSHOT: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book', sheetOrder: [] } }

let ctx: Context
let root: string
let session: Session
let scope: { sessionId: SessionId; workspaceRoot: string }
const events: SessionEvent[] = []

beforeEach(async () => {
  events.length = 0
  root = await mkdtemp(join(tmpdir(), 'dsh-command-broker-'))
  ctx = new Context()
  await ctx.plugin(LocalFileSystem, { cwd: root })
  await ctx.plugin(SessionStore)
  await ctx.plugin(LocalOfficeExchange, { commandTimeoutMs: 25 })
  session = ctx.sessions.create(SessionId('command-test'), { meta: { cwd: root } })
  scope = { sessionId: session.id, workspaceRoot: root }
  ctx.on('session/event', (subject, event) => {
    if (subject === session) events.push(event)
  })
})

afterEach(async () => {
  await ctx.fiber.dispose()
  await rm(root, { recursive: true, force: true })
})

/** Read the last appended command request. */
function lastCommand(): { requestId: string; fromVersion: number } {
  const event = events.at(-1)
  if (event?.type !== 'document/command') throw new Error('the broker did not append a document/command')
  return { requestId: event.data.requestId, fromVersion: event.data.fromVersion }
}

describe('document command broker', () => {
  it('appends one request and resolves the matching correlated result', async () => {
    const pending = ctx.officeDocument.dispatchCommand({
      session,
      scope,
      documentId: DOCUMENT,
      operation: { kind: 'read' },
      fromVersion: 2,
    }, SIGNAL)
    const { requestId, fromVersion } = lastCommand()
    const result: OfficeCommandResult = { status: 'read', snapshot: SNAPSHOT, contentVersion: 3 }
    await ctx.officeDocument.completeCommand(scope, requestId, result, SIGNAL)

    await expect(pending).resolves.toEqual(result)
    expect(fromVersion).toBe(2)
  })

  it('resolves applied results and records the observed version', async () => {
    const pending = ctx.officeDocument.dispatchCommand({
      session,
      scope,
      documentId: DOCUMENT,
      operation: { kind: 'edit', steps: [{ kind: 'cell', sheetId: 's', row: 0, column: 0, value: 1 }] },
      fromVersion: 0,
    }, SIGNAL)
    const { requestId } = lastCommand()
    const result: OfficeCommandResult = { status: 'applied', snapshot: SNAPSHOT, contentVersion: 1 }
    await ctx.officeDocument.completeCommand(scope, requestId, result, SIGNAL)
    await expect(pending).resolves.toEqual(result)

    ctx.officeCommandBroker.remember(DOCUMENT, 7)
    expect(ctx.officeCommandBroker.versionOf(DOCUMENT)).toBe(7)
    expect(ctx.officeCommandBroker.versionOf(DocumentId('other'))).toBe(0)
  })

  it('fails timeout without a result', async () => {
    vi.useFakeTimers()
    try {
      const pending = ctx.officeCommandBroker.request({
        session,
        scope,
        documentId: DOCUMENT,
        operation: { kind: 'read' },
        fromVersion: 0,
      }, SIGNAL)
      await vi.advanceTimersByTimeAsync(25)
      await expect(pending).resolves.toEqual({ status: 'timeout' })
    } finally {
      vi.useRealTimers()
    }
  })

  it('ignores an unknown or late result with a diagnostic', async () => {
    const warn = vi.spyOn(ctx.logger, 'warn').mockImplementation(() => {})
    await ctx.officeDocument.completeCommand(scope, 'unknown', { status: 'pending-edits' }, SIGNAL)
    expect(warn).toHaveBeenCalledWith('Ignoring document command result for unknown request "unknown".')
    warn.mockRestore()
  })

  it('ignores a result for another Session and its request still times out', async () => {
    vi.useFakeTimers()
    try {
      const warn = vi.spyOn(ctx.logger, 'warn').mockImplementation(() => {})
      const pending = ctx.officeCommandBroker.request({
        session,
        scope,
        documentId: DOCUMENT,
        operation: { kind: 'read' },
        fromVersion: 0,
      }, SIGNAL)
      const { requestId } = lastCommand()
      await ctx.officeDocument.completeCommand(
        { sessionId: SessionId('other-session'), workspaceRoot: root },
        requestId,
        { status: 'pending-edits' },
        SIGNAL,
      )
      expect(warn).toHaveBeenCalledWith(
        `Ignoring document command result for request "${requestId}" from another Session.`,
      )
      await vi.advanceTimersByTimeAsync(25)
      await expect(pending).resolves.toEqual({ status: 'timeout' })
      warn.mockRestore()
    } finally {
      vi.useRealTimers()
    }
  })

  it('withdraws a pending request when the caller cancels', async () => {
    const warn = vi.spyOn(ctx.logger, 'warn').mockImplementation(() => {})
    const controller = new AbortController()
    const pending = ctx.officeCommandBroker.request({
      session,
      scope,
      documentId: DOCUMENT,
      operation: { kind: 'read' },
      fromVersion: 0,
    }, controller.signal)
    const { requestId } = lastCommand()
    const reason = new Error('caller cancelled')
    controller.abort(reason)
    await expect(pending).rejects.toBe(reason)
    await ctx.officeDocument.completeCommand(scope, requestId, { status: 'pending-edits' }, SIGNAL)
    expect(warn).toHaveBeenCalledTimes(1)
    warn.mockRestore()
  })

  it('refuses a caller that is already aborted before appending', async () => {
    const before = events.length
    await expect(ctx.officeCommandBroker.request({
      session,
      scope,
      documentId: DOCUMENT,
      operation: { kind: 'read' },
      fromVersion: 0,
    }, AbortSignal.abort(new Error('already cancelled')))).rejects.toThrow('already cancelled')
    expect(events.length).toBe(before)
  })

  it('drops the pending entry when the command cannot be appended', async () => {
    const append = vi.spyOn(session, 'append').mockImplementation(() => { throw new Error('append failed') })
    await expect(ctx.officeCommandBroker.request({
      session,
      scope,
      documentId: DOCUMENT,
      operation: { kind: 'read' },
      fromVersion: 0,
    }, SIGNAL)).rejects.toThrow('append failed')
    append.mockRestore()

    const warn = vi.spyOn(ctx.logger, 'warn').mockImplementation(() => {})
    await ctx.officeDocument.completeCommand(scope, 'any', { status: 'pending-edits' }, SIGNAL)
    expect(warn).toHaveBeenCalledTimes(1)
    warn.mockRestore()
  })

  it('refuses a wire result whose version is not a non-negative integer', async () => {
    await expect(ctx.officeDocument.completeCommand(scope, 'request', {
      status: 'read',
      snapshot: SNAPSHOT,
      contentVersion: -1,
    }, SIGNAL)).rejects.toThrow(/reported contentVersion -1/)
    await expect(ctx.officeDocument.completeCommand(scope, 'request', {
      status: 'conflict',
      currentVersion: 1.5,
    }, SIGNAL)).rejects.toThrow(/reported currentVersion 1.5/)
  })

  it('drops an answer delivered after the caller aborted', async () => {
    const warn = vi.spyOn(ctx.logger, 'warn').mockImplementation(() => {})
    await ctx.officeDocument.completeCommand(scope, 'late', { status: 'document-not-open' }, AbortSignal.abort())
    expect(warn).not.toHaveBeenCalled()
    warn.mockRestore()
  })

  it('fails when the command broker is not mounted with the Remote', async () => {
    const bare = new Context()
    try {
      await bare.plugin(LocalFileSystem, { cwd: root })
      await bare.plugin(SessionStore)
      await bare.plugin(OfficeDocumentRemote, { maxBytes: 1024 })
      await expect(bare.officeDocument.completeCommand(scope, 'request', {
        status: 'unsupported-feature',
        features: [{ kind: 'macro' }],
      }, SIGNAL)).rejects.toThrow(/officeCommandBroker is absent/)
    } finally {
      await bare.fiber.dispose()
    }
  })

  it('clears pending requests when the composition is disposed', async () => {
    vi.useFakeTimers()
    try {
      const broker = ctx.officeCommandBroker
      void broker.request({
        session,
        scope,
        documentId: DOCUMENT,
        operation: { kind: 'read' },
        fromVersion: 0,
      }, SIGNAL)
      expect(vi.getTimerCount()).toBe(1)
      await ctx.fiber.dispose()
      expect(vi.getTimerCount()).toBe(0)
    } finally {
      vi.useRealTimers()
    }
  })
})
