/**
 * The renderer-side `document/command` bridge.
 *
 * A fake Session event feed stands in for the controller's event source and a
 * fake Remote callback records the correlated replies, so every result path is
 * exercised without a live editor or a Host.
 */
import { describe, expect, it, vi } from 'vitest'
import type { DocumentEditOperation } from '@forwork/session/src/document-events.ts'
import type { OfficeDocumentSnapshot, UnsupportedFeature } from '@forwork/office-exchange/types'
import {
  attachSessionCommandBridge,
  createCommandBridge,
  sessionCommandSource,
  type CommandEventChange,
  type CommandEventEntry,
  type CommandEventFeed,
  type DocumentCommandRequest,
  type DocumentCommandResult,
  type OfficeApplyOutcome,
  type OfficeCommandRemote,
  type OfficeCommandTarget,
} from '../src/client/command-bridge.ts'

const SNAPSHOT: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book' } }
const STEPS: DocumentEditOperation = [{ kind: 'cell', sheetId: 'sheet-1', row: 0, column: 0, value: 1 }]

function command(overrides: Partial<DocumentCommandRequest> = {}): DocumentCommandRequest {
  return {
    requestId: 'request-1',
    documentId: 'document-1',
    operation: { kind: 'edit', steps: STEPS },
    fromVersion: 0,
    ...overrides,
  }
}

/** A controllable Session event feed. */
class FakeFeed implements CommandEventFeed {
  private change: CommandEventChange = { kind: 'replace' }
  private readonly listeners = new Set<() => void>()

  getSnapshot(): { change: CommandEventChange } {
    return { change: this.change }
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener)
    return () => { this.listeners.delete(listener) }
  }

  publish(change: CommandEventChange): void {
    this.change = change
    for (const listener of [...this.listeners]) listener()
  }

  append(entry: CommandEventEntry): void {
    this.publish({ kind: 'append', entries: [entry] })
  }

  appendEvent(type: string, data?: unknown): void {
    this.append({ type: 'event', event: data === undefined ? { type } : { type, data } })
  }

  appendCommand(request: DocumentCommandRequest): void {
    this.appendEvent('document/command', request)
  }
}

interface TargetOptions {
  readonly documentId?: string
  readonly version?: number | undefined
  readonly open?: boolean
  readonly dirty?: boolean
  readonly outcome?: OfficeApplyOutcome
  readonly snapshot?: OfficeDocumentSnapshot | undefined
}

function fakeTarget(options: TargetOptions = {}): {
  target: OfficeCommandTarget
  apply: ReturnType<typeof vi.fn>
  snapshot: ReturnType<typeof vi.fn>
} {
  const hasVersion = !Object.prototype.hasOwnProperty.call(options, 'version') || options.version !== undefined
  let version = options.version ?? 0
  const outcome = options.outcome ?? { status: 'applied' }
  const apply = vi.fn((_operation: DocumentEditOperation): OfficeApplyOutcome => {
    if (outcome.status === 'applied') version += 1
    return outcome
  })
  const snapshot = vi.fn(() => (
    Object.prototype.hasOwnProperty.call(options, 'snapshot') ? options.snapshot : SNAPSHOT
  ))
  return {
    target: {
      documentId: options.documentId ?? 'document-1',
      version: () => hasVersion ? version : undefined,
      isOpen: () => options.open ?? true,
      isDirty: () => options.dirty ?? false,
      apply,
      snapshot,
    },
    apply,
    snapshot,
  }
}

function bridgeOver(feed: FakeFeed): {
  attach: ReturnType<typeof createCommandBridge>['attach']
  complete: ReturnType<typeof vi.fn>
} {
  const complete = vi.fn<(requestId: string, result: DocumentCommandResult) => void>()
  const bridge = createCommandBridge({ source: sessionCommandSource(feed), complete })
  return { attach: target => bridge.attach(target), complete }
}

describe('createCommandBridge', () => {
  it('applies a matching edit once and returns the advanced version and snapshot', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const { target, apply, snapshot } = fakeTarget()
    attach(target)
    feed.appendCommand(command())
    expect(apply).toHaveBeenCalledOnce()
    expect(apply).toHaveBeenCalledWith(STEPS)
    expect(snapshot).toHaveBeenCalledOnce()
    expect(complete).toHaveBeenCalledOnce()
    expect(complete).toHaveBeenCalledWith('request-1', {
      status: 'applied',
      snapshot: SNAPSHOT,
      contentVersion: 1,
    })
  })

  it('reads the live unit without applying or requiring a matching version', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const { target, apply } = fakeTarget({ version: 7 })
    attach(target)
    feed.appendCommand(command({ operation: { kind: 'read' }, fromVersion: 3 }))
    expect(apply).not.toHaveBeenCalled()
    expect(complete).toHaveBeenCalledOnce()
    expect(complete).toHaveBeenCalledWith('request-1', {
      status: 'read',
      snapshot: SNAPSHOT,
      contentVersion: 7,
    })
  })

  it('refuses a read while the live unit has unsaved user edits', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const { target, apply } = fakeTarget({ dirty: true })
    attach(target)
    feed.appendCommand(command({ operation: { kind: 'read' } }))
    expect(apply).not.toHaveBeenCalled()
    expect(complete).toHaveBeenCalledOnce()
    expect(complete).toHaveBeenCalledWith('request-1', { status: 'pending-edits' })
  })

  it('ignores a command for another document without replying', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const { target, apply } = fakeTarget()
    attach(target)
    feed.appendCommand(command({ documentId: 'document-2' }))
    expect(apply).not.toHaveBeenCalled()
    expect(complete).not.toHaveBeenCalled()
  })

  it('replies with the live version and applies nothing when an edit is stale', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const { target, apply } = fakeTarget({ version: 4 })
    attach(target)
    feed.appendCommand(command({ fromVersion: 3 }))
    expect(apply).not.toHaveBeenCalled()
    expect(complete).toHaveBeenCalledOnce()
    expect(complete).toHaveBeenCalledWith('request-1', { status: 'conflict', currentVersion: 4 })
  })

  it('refuses an edit while the live unit has unsaved user edits', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const { target, apply } = fakeTarget({ dirty: true })
    attach(target)
    feed.appendCommand(command())
    expect(apply).not.toHaveBeenCalled()
    expect(complete).toHaveBeenCalledOnce()
    expect(complete).toHaveBeenCalledWith('request-1', { status: 'pending-edits' })
  })

  it('handles a duplicate command for the same request only once', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const { target, apply } = fakeTarget()
    attach(target)
    feed.appendCommand(command())
    feed.appendCommand(command())
    expect(apply).toHaveBeenCalledOnce()
    expect(complete).toHaveBeenCalledOnce()
  })

  it('returns the feature when the adapter cannot apply the operation', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const features: readonly UnsupportedFeature[] = [{ kind: 'other', detail: 'paragraph' }]
    const { target, snapshot } = fakeTarget({ outcome: { status: 'unsupported-feature', features } })
    attach(target)
    feed.appendCommand(command())
    expect(snapshot).not.toHaveBeenCalled()
    expect(complete).toHaveBeenCalledOnce()
    expect(complete).toHaveBeenCalledWith('request-1', { status: 'unsupported-feature', features })
  })

  it('reports document-not-open when no editor unit is mounted', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const { target, apply } = fakeTarget({ open: false })
    attach(target)
    feed.appendCommand(command())
    expect(apply).not.toHaveBeenCalled()
    expect(complete).toHaveBeenCalledWith('request-1', { status: 'document-not-open' })
  })

  it('reports document-not-open when the live version is unknown', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const { target, apply } = fakeTarget({ version: undefined })
    attach(target)
    feed.appendCommand(command())
    expect(apply).not.toHaveBeenCalled()
    expect(complete).toHaveBeenCalledWith('request-1', { status: 'document-not-open' })
  })

  it('reports document-not-open when a read has no live snapshot', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const { target, apply } = fakeTarget({ snapshot: undefined })
    attach(target)
    feed.appendCommand(command({ operation: { kind: 'read' } }))
    expect(apply).not.toHaveBeenCalled()
    expect(complete).toHaveBeenCalledWith('request-1', { status: 'document-not-open' })
  })

  it('reports document-not-open when an applied unit loses its snapshot or version', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const { target } = fakeTarget({ snapshot: undefined })
    attach(target)
    feed.appendCommand(command())
    expect(complete).toHaveBeenCalledWith('request-1', { status: 'document-not-open' })
  })

  it('stops handling commands once the attachment is disposed', () => {
    const feed = new FakeFeed()
    const { attach, complete } = bridgeOver(feed)
    const { target, apply } = fakeTarget()
    const dispose = attach(target)
    dispose()
    feed.appendCommand(command())
    expect(apply).not.toHaveBeenCalled()
    expect(complete).not.toHaveBeenCalled()
  })
})

describe('attachSessionCommandBridge', () => {
  it('does nothing without a retained Session binding', () => {
    const feed = new FakeFeed()
    const { target } = fakeTarget()
    const completeCommand = vi.fn()
    const registerCommandTarget = vi.fn(async () => undefined)
    const dispose = attachSessionCommandBridge(undefined, 'session-1', target, {
      read: async () => undefined,
      registerCommandTarget,
      completeCommand,
    })
    feed.appendCommand(command())
    expect(registerCommandTarget).not.toHaveBeenCalled()
    expect(completeCommand).not.toHaveBeenCalled()
    dispose()
  })

  it('registers the attached target and releases the same registration on detach', () => {
    const feed = new FakeFeed()
    const { target } = fakeTarget()
    const registerCommandTarget = vi.fn(
      async (_sessionId: string, _documentId: string, _registrationId: string) => undefined,
    )
    const unregisterCommandTarget = vi.fn(async () => undefined)
    const dispose = attachSessionCommandBridge({ eventSource: feed }, 'session-1', target, {
      read: async () => undefined,
      registerCommandTarget,
      unregisterCommandTarget,
    })

    expect(registerCommandTarget).toHaveBeenCalledOnce()
    const [sessionId, documentId, registrationId] = registerCommandTarget.mock.calls[0] ?? []
    expect(sessionId).toBe('session-1')
    expect(documentId).toBe('document-1')
    expect(typeof registrationId).toBe('string')
    expect(registrationId).not.toBe('')

    dispose()
    expect(unregisterCommandTarget).toHaveBeenCalledOnce()
    expect(unregisterCommandTarget).toHaveBeenCalledWith('session-1', 'document-1', registrationId)
  })

  it('settles through the Session-scoped Remote method', () => {
    const feed = new FakeFeed()
    const { target } = fakeTarget()
    const completeCommand = vi.fn(async () => undefined)
    const remote: OfficeCommandRemote = { read: async () => undefined, completeCommand }
    const dispose = attachSessionCommandBridge({ eventSource: feed }, 'session-1', target, remote)
    feed.appendCommand(command())
    expect(completeCommand).toHaveBeenCalledOnce()
    expect(completeCommand).toHaveBeenCalledWith('session-1', 'request-1', {
      status: 'applied',
      snapshot: SNAPSHOT,
      contentVersion: 1,
    })
    dispose()
  })

  it('leaves the Host call to its timeout when the Remote method is absent', () => {
    const feed = new FakeFeed()
    const { target } = fakeTarget()
    const dispose = attachSessionCommandBridge(
      { eventSource: feed }, 'session-1', target, { read: async () => undefined },
    )
    feed.appendCommand(command())
    dispose()
  })
})

describe('sessionCommandSource', () => {
  it('never replays the window it attaches to', () => {
    const feed = new FakeFeed()
    feed.publish({ kind: 'replace' })
    const listener = vi.fn()
    sessionCommandSource(feed).subscribe(listener)
    expect(listener).not.toHaveBeenCalled()
  })

  it('ignores deltas that are not appends', () => {
    const feed = new FakeFeed()
    const listener = vi.fn()
    sessionCommandSource(feed).subscribe(listener)
    feed.publish({ kind: 'prepend', entries: [{ type: 'event', event: { type: 'document/command', data: command() } }] })
    expect(listener).not.toHaveBeenCalled()
  })

  it('ignores an append without entries', () => {
    const feed = new FakeFeed()
    const listener = vi.fn()
    sessionCommandSource(feed).subscribe(listener)
    feed.publish({ kind: 'append' })
    expect(listener).not.toHaveBeenCalled()
  })

  it('ignores transient rows and durable events of another type', () => {
    const feed = new FakeFeed()
    const listener = vi.fn()
    sessionCommandSource(feed).subscribe(listener)
    feed.append({ type: 'transient', event: { type: 'document/command', data: command() } })
    feed.appendEvent('turn/start', { turn: 1 })
    expect(listener).not.toHaveBeenCalled()
  })

  it('ignores malformed command payloads', () => {
    const feed = new FakeFeed()
    const listener = vi.fn()
    sessionCommandSource(feed).subscribe(listener)
    feed.appendEvent('document/command')
    feed.appendEvent('document/command', null)
    feed.appendEvent('document/command', [])
    feed.appendEvent('document/command', { requestId: 1, documentId: 'document-1', operation: { kind: 'read' }, fromVersion: 0 })
    feed.appendEvent('document/command', { requestId: 'r', documentId: 2, operation: { kind: 'read' }, fromVersion: 0 })
    feed.appendEvent('document/command', { requestId: 'r', documentId: 'document-1', operation: { kind: 'read' }, fromVersion: '0' })
    feed.appendEvent('document/command', { requestId: 'r', documentId: 'document-1', fromVersion: 0 })
    feed.appendEvent('document/command', { requestId: 'r', documentId: 'document-1', operation: null, fromVersion: 0 })
    feed.appendEvent('document/command', { requestId: 'r', documentId: 'document-1', operation: [], fromVersion: 0 })
    feed.appendEvent('document/command', { requestId: 'r', documentId: 'document-1', operation: { kind: 'other' }, fromVersion: 0 })
    feed.appendEvent('document/command', { requestId: 'r', documentId: 'document-1', operation: { kind: 'edit' }, fromVersion: 0 })
    feed.appendEvent('document/command', {
      requestId: 'r', documentId: 'document-1', operation: { kind: 'edit', steps: 'none' }, fromVersion: 0,
    })
    feed.appendEvent('document/command', {
      requestId: 'r', documentId: 'document-1', operation: { kind: 'edit', steps: [null] }, fromVersion: 0,
    })
    feed.appendEvent('document/command', {
      requestId: 'r', documentId: 'document-1', operation: { kind: 'edit', steps: [{ kind: 'unknown' }] }, fromVersion: 0,
    })
    feed.appendEvent('document/command', {
      requestId: 'r',
      documentId: 'document-1',
      operation: { kind: 'edit', steps: [{ kind: 'cell', row: 0, column: 0, value: 1 }] },
      fromVersion: 0,
    })
    feed.appendEvent('document/command', {
      requestId: 'r',
      documentId: 'document-1',
      operation: { kind: 'edit', steps: [{ kind: 'cell', sheetId: 's', row: '0', column: 0, value: 1 }] },
      fromVersion: 0,
    })
    feed.appendEvent('document/command', {
      requestId: 'r',
      documentId: 'document-1',
      operation: { kind: 'edit', steps: [{ kind: 'cell', sheetId: 's', row: 0, column: 0 }] },
      fromVersion: 0,
    })
    feed.appendEvent('document/command', {
      requestId: 'r',
      documentId: 'document-1',
      operation: { kind: 'edit', steps: [{ kind: 'paragraph', index: 0 }] },
      fromVersion: 0,
    })
    feed.appendEvent('document/command', {
      requestId: 'r',
      documentId: 'document-1',
      operation: { kind: 'edit', steps: [{ kind: 'row', sheetId: 's', row: 0, count: 1, action: 'move' }] },
      fromVersion: 0,
    })
    feed.appendEvent('document/command', {
      requestId: 'r',
      documentId: 'document-1',
      operation: { kind: 'edit', steps: [{ kind: 'column', sheetId: 's', column: 0, count: 1 }] },
      fromVersion: 0,
    })
    feed.appendEvent('document/command', {
      requestId: 'r',
      documentId: 'document-1',
      operation: { kind: 'edit', steps: [{ kind: 'formula', sheetId: 's', row: 0, column: 0, formula: 1 }] },
      fromVersion: 0,
    })
    feed.appendEvent('document/command', {
      requestId: 'r',
      documentId: 'document-1',
      operation: {
        kind: 'edit',
        steps: [{ kind: 'style', sheetId: 's', startRow: 0, startColumn: 0, endRow: 0 }],
      },
      fromVersion: 0,
    })
    expect(listener).not.toHaveBeenCalled()
  })

  it('accepts read commands and every known edit step kind', () => {
    const feed = new FakeFeed()
    const listener = vi.fn()
    sessionCommandSource(feed).subscribe(listener)
    feed.appendCommand(command({ requestId: 'read', operation: { kind: 'read' } }))
    feed.appendCommand(command({
      requestId: 'steps',
      operation: {
        kind: 'edit',
        steps: [
          { kind: 'cell', sheetId: 's', row: 1, column: 2, value: 'text' },
          { kind: 'paragraph', index: 0, text: 'Hello' },
          { kind: 'row', sheetId: 's', row: 1, count: 2, action: 'insert' },
          { kind: 'column', sheetId: 's', column: 1, count: 2, action: 'delete' },
          { kind: 'formula', sheetId: 's', row: 1, column: 1, formula: '=SUM(A1:A2)' },
          {
            kind: 'style',
            sheetId: 's',
            startRow: 0,
            startColumn: 0,
            endRow: 1,
            endColumn: 1,
            style: { bl: 1 },
          },
        ],
      },
    }))
    expect(listener).toHaveBeenCalledTimes(2)
    expect(listener.mock.calls[0]?.[0]).toMatchObject({ requestId: 'read', operation: { kind: 'read' } })
    expect(listener.mock.calls[1]?.[0]).toMatchObject({ requestId: 'steps', operation: { kind: 'edit' } })
  })
})
