/** Real Host composition for the Office document tools, with a fake live editor. */
import { copyFile, mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Context } from '@forwork/cordis'
import type { Agent } from '@forwork/agent'
import { mountAgentLoopTestDependencies, mountAgentLoopTestHarness } from '@forwork/agent-loop-testkit'
import type { WorkspaceFileScope } from '@forwork/api-workspace-files'
import { LocalFileSystem } from '@forwork/fs-local'
import { ToolCallId } from '@forwork/llm'
import { SessionId } from '@forwork/session'
import type { SessionEvent } from '@forwork/session'
import { executeToolFixture } from '@forwork/tool-testkit'
import type { ToolExecutionInput, ToolExecutionResult } from '@forwork/tools'
import LocalOfficeExchange, { OfficeDocumentRemote } from '@forwork/office-exchange'
import * as ToolOffice from '../src/index.ts'

const FIXTURES = fileURLToPath(new URL('../../office-exchange/tests/fixtures/', import.meta.url))

/** One mounted Host composition and its calling agent. */
export interface ToolTestHarness {
  /** Owning Host context. */
  readonly ctx: Context
  /** Workspace root all paths resolve against. */
  readonly root: string
  /** Calling agent whose Session owns the tools. */
  readonly agent: Agent
  /** Session identity and workspace root for the Remote calls. */
  readonly scope: WorkspaceFileScope
  /** The real `officeDocument` service, spied by each test. */
  readonly server: OfficeDocumentRemote
  /** Every event appended to the calling Session after the harness was built. */
  readonly events: SessionEvent[]
  /** Copy one fixture into the workspace root and return its file name. */
  copy(name: string): Promise<string>
  /** Create one more production Agent, optionally without a workspace root. */
  createAgent(id: SessionId, cwd?: string): Promise<Agent>
  /** Execute one tool call through the real pipeline. */
  execute(name: string, arguments_: Record<string, unknown>, signal: AbortSignal, agent?: Agent): Promise<ToolExecutionResult>
}

/**
 * Mount files and sessions, Office Exchange, its Remote, and the document tools.
 * @param config - tool caps for the mounted plugin.
 * @returns the composition, cleaned up by {@link disposeHarness}.
 */
export async function createHarness(config: { summaryCells?: number; maxEditSteps?: number } = {}): Promise<ToolTestHarness> {
  const root = await mkdtemp(join(tmpdir(), 'dsh-tool-office-'))
  const ctx = new Context()
  await ctx.plugin(LocalFileSystem, { cwd: root })
  await mountAgentLoopTestDependencies(ctx)
  await ctx.plugin(LocalOfficeExchange, { commandTimeoutMs: 25 })
  await ctx.plugin(ToolOffice, config)
  const loop = await mountAgentLoopTestHarness(ctx)
  const agent = await loop.create(SessionId('tool-office-test'), {}, { cwd: root })
  const scope = { sessionId: agent.session.id, workspaceRoot: root }
  const events: SessionEvent[] = []
  ctx.on('session/event', (session, event) => {
    if (session === agent.session) events.push(event)
  })
  return {
    ctx,
    root,
    agent,
    scope,
    server: ctx.officeDocument,
    events,
    async copy(name) {
      await copyFile(join(FIXTURES, name), join(root, name))
      return name
    },
    async createAgent(id, cwd) {
      return await loop.create(id, {}, cwd === undefined ? {} : { cwd })
    },
    async execute(name, arguments_, signal, caller = agent) {
      const input: ToolExecutionInput = {
        callId: ToolCallId(`call-${name}`),
        name,
        arguments: arguments_,
        signal,
        agent: caller,
      }
      return await executeToolFixture(ctx, input)
    },
  }
}

/**
 * Dispose one harness composition and remove its workspace.
 * @param harness - composition returned by {@link createHarness}.
 */
export async function disposeHarness(harness: ToolTestHarness): Promise<void> {
  await harness.ctx.fiber.dispose()
  await rm(harness.root, { recursive: true, force: true })
}

/** The Session-scoped document id for one workspace-relative path. */
export function documentIdFor(harness: ToolTestHarness, path: string): string {
  return `${harness.agent.session.id}:${path}`
}
