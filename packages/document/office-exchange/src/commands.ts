/**
 * Correlated document-command broker. Host tools record one pending request,
 * append the matching `document/command` Session event, and wait while the
 * live editor applies the operation and answers through
 * `officeDocument.completeCommand`.
 *
 * The broker is a service, not a module global: pending requests, their
 * timeouts, and the last observed content version belong to the composition
 * that mounted it and are released with it.
 * @module @deepseek-ai/dsh-office-exchange/commands
 */

import { randomUUID } from 'node:crypto'
import { Context, Service } from '@deepseek-ai/cordis'
import z from '@deepseek-ai/schemastery'
import type { WorkspaceFileScope } from '@deepseek-ai/dsh-api-workspace-files'
import type { DocumentCommandOperation, DocumentId, Session } from '@deepseek-ai/dsh-session'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from './config.ts'
import type { OfficeCommandOutcome, OfficeCommandResult } from './types.ts'

declare module '@deepseek-ai/cordis' {
  interface Context {
    /** Correlated command transport between Host document tools and the live editor. */
    officeCommandBroker: OfficeCommandBroker
  }
}

/** Deployment bounds for the correlated command transport. */
export interface CommandBrokerConfig {
  /** Longest wait for the live editor to answer one command, in milliseconds. */
  readonly commandTimeoutMs: number
}

/** Validated timeout applied to every pending command. */
export const CommandBrokerConfig: z<Partial<CommandBrokerConfig>, CommandBrokerConfig> = z.object({
  commandTimeoutMs: z.natural().min(1).max(2_147_483_647)
    .default(DEFAULT_OFFICE_EXCHANGE_CONFIG.commandTimeoutMs),
})

/** One command the Host asks the live editor to execute. */
export interface DocumentCommandRequest {
  /** Session that records the request and will receive the resulting edit fact. */
  readonly session: Session
  /** Header-derived workspace root for the Session identity on the wire. */
  readonly scope: WorkspaceFileScope
  /** Stable identity of the open document. */
  readonly documentId: DocumentId
  /** Operation to read or edit. */
  readonly operation: DocumentCommandOperation
  /** Content version the caller believes the live unit holds. */
  readonly fromVersion: number
}

/** One pending request and the timer that bounds its wait. */
interface PendingCommand {
  readonly sessionId: SessionId
  readonly resolve: (outcome: OfficeCommandOutcome) => void
  readonly timer: ReturnType<typeof setTimeout>
}

/**
 * Owns pending document commands and the last observed content version per
 * document. One instance serves every Session in its composition; a result is
 * accepted only for the Session that recorded the request.
 */
export class OfficeCommandBroker extends Service {
  static Config = CommandBrokerConfig
  private readonly pending = new Map<string, PendingCommand>()
  private readonly versions = new Map<string, number>()

  /**
   * @param ctx - Host context whose service key is `officeCommandBroker`.
   * @param config - validated maximum wait for one command.
   */
  constructor(ctx: Context, private readonly config: CommandBrokerConfig) {
    super(ctx, 'officeCommandBroker')
    ctx.effect(() => () => { this.cancelAll() })
  }

  /**
   * Append one `document/command` request and wait for the live editor's
   * answer.
   * @param request - session, document, operation, and expected version.
   * @param signal - caller cancellation; aborting withdraws the pending request.
   * @returns the editor's result, or `timeout` when no answer arrived in time.
   * @throws the reason carried by an aborted `signal`, or an append failure.
   */
  async request(request: DocumentCommandRequest, signal: AbortSignal): Promise<OfficeCommandOutcome> {
    signal.throwIfAborted()
    const requestId = randomUUID()
    const settled = new Promise<OfficeCommandOutcome>((resolve) => {
      const timer = setTimeout(() => {
        this.pending.delete(requestId)
        resolve({ status: 'timeout' })
      }, this.config.commandTimeoutMs)
      this.pending.set(requestId, { sessionId: request.session.id, resolve, timer })
    })
    try {
      request.session.append('document/command', {
        requestId,
        documentId: request.documentId,
        operation: request.operation,
        fromVersion: request.fromVersion,
      })
    } catch (cause) {
      this.cancel(requestId)
      throw cause
    }
    return await this.settle(requestId, settled, signal)
  }

  /**
   * Resolve one pending request with the live editor's result.
   *
   * A result for an unknown request id, a request from another Session, or a
   * request that already timed out is ignored and reported as a diagnostic;
   * late answers never change durable state.
   * @param scope - Session identity that produced the result.
   * @param requestId - correlation id copied from the command event.
   * @param result - the live editor's answer.
   */
  complete(scope: WorkspaceFileScope, requestId: string, result: OfficeCommandResult): void {
    const pending = this.pending.get(requestId)
    if (pending === undefined) {
      this.ctx.logger.warn(`Ignoring document command result for unknown request "${requestId}".`)
      return
    }
    if (pending.sessionId !== scope.sessionId) {
      this.ctx.logger.warn(`Ignoring document command result for request "${requestId}" from another Session.`)
      return
    }
    clearTimeout(pending.timer)
    this.pending.delete(requestId)
    pending.resolve(result)
  }

  /**
   * Read the last content version reported for one document.
   * @param documentId - stable identity of the document.
   * @returns the last reported version, or `0` before any command reported one.
   */
  versionOf(documentId: DocumentId): number {
    return this.versions.get(documentId) ?? 0
  }

  /**
   * Record the content version one document currently holds.
   * @param documentId - stable identity of the document.
   * @param version - non-negative version reported by the live editor.
   */
  remember(documentId: DocumentId, version: number): void {
    this.versions.set(documentId, version)
  }

  /**
   * Race the pending answer against the caller's cancellation and always
   * release the abort listener.
   */
  private async settle(
    requestId: string,
    settled: Promise<OfficeCommandOutcome>,
    signal: AbortSignal,
  ): Promise<OfficeCommandOutcome> {
    const aborted = Promise.withResolvers<never>()
    const onAbort = (): void => {
      this.cancel(requestId)
      aborted.reject(signal.reason)
    }
    signal.addEventListener('abort', onAbort, { once: true })
    try {
      return await Promise.race([settled, aborted.promise])
    } finally {
      signal.removeEventListener('abort', onAbort)
    }
  }

  /** Withdraw one pending request and stop its timeout. */
  private cancel(requestId: string): void {
    clearTimeout(this.pending.get(requestId)?.timer)
    this.pending.delete(requestId)
  }

  /** Withdraw every pending request during disposal. */
  private cancelAll(): void {
    for (const pending of this.pending.values()) clearTimeout(pending.timer)
    this.pending.clear()
    this.versions.clear()
  }
}

export default OfficeCommandBroker
