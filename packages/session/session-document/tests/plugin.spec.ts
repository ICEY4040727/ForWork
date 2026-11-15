import { describe, expect, it } from 'vitest'
import { Context } from '@deepseek-ai/cordis'
import type { Fiber } from '@deepseek-ai/cordis'
import { assembleContextFor } from '@deepseek-ai/dsh-agent'
import type { Agent } from '@deepseek-ai/dsh-agent'
import { DocumentId, SessionId } from '@deepseek-ai/dsh-session'
import { mountAgentLoopTestDependencies, mountAgentLoopTestHarness } from '@deepseek-ai/dsh-agent-loop-testkit'
import * as DocumentPlugin from '../src/index.ts'

const FIRST = DocumentId('plugin:first.xlsx')
const SECOND = DocumentId('plugin:second.docx')
const SECTION = 'document:open-set'

async function mount(): Promise<{ ctx: Context; agent: Agent; fiber: Fiber }> {
  const ctx = new Context()
  await mountAgentLoopTestDependencies(ctx)
  const loop = await mountAgentLoopTestHarness(ctx)
  const agent = await loop.create(SessionId('document-plugin'))
  const fiber = await ctx.plugin(DocumentPlugin)
  return { ctx, agent, fiber }
}

/** The rendered text of the open-document section, or `undefined` when it is absent. */
async function sectionText(ctx: Context, context = {}): Promise<string | undefined> {
  const assembly = await ctx.systemPrompt.assemble(context)
  return assembly.sections.find(section => section.name === SECTION)?.text
}

describe('session-document plugin', () => {
  it('renders the open set for the assembled agent session', async () => {
    const { ctx, agent } = await mount()
    try {
      agent.session.append('document/open', {
        documentId: FIRST,
        path: 'work/first.xlsx',
        format: 'xlsx',
        contentVersion: 0,
      })
      agent.session.append('document/open', {
        documentId: SECOND,
        path: 'work/second.docx',
        format: 'docx',
        contentVersion: 0,
      })
      agent.session.append('document/missing', { documentId: SECOND, path: 'work/second.docx' })

      expect(await sectionText(ctx, assembleContextFor(agent))).toBe(
        [
          'Open documents in this session:',
          '1. first.xlsx (xlsx)',
          '2. second.docx (docx, missing)',
        ].join('\n'),
      )
    } finally {
      await ctx.fiber.dispose()
    }
  })

  it('renders nothing without an agent or without an open document', async () => {
    const { ctx, agent } = await mount()
    try {
      expect(await sectionText(ctx)).toBe('')
      expect(await sectionText(ctx, assembleContextFor(agent))).toBe('')
    } finally {
      await ctx.fiber.dispose()
    }
  })

  it('removes the projection and the prompt section when the plugin fiber is disposed', async () => {
    const { ctx, agent, fiber } = await mount()
    try {
      agent.session.append('document/open', {
        documentId: FIRST,
        path: 'work/first.xlsx',
        format: 'xlsx',
        contentVersion: 0,
      })
      expect(ctx.sessionProjections.snapshot(agent.session, ['document']).values.document).toBeDefined()
      expect(await sectionText(ctx, assembleContextFor(agent))).toContain('first.xlsx')

      await fiber.dispose()

      expect(ctx.sessionProjections.snapshot(agent.session, ['document']).values).not.toHaveProperty('document')
      expect(await sectionText(ctx, assembleContextFor(agent))).toBeUndefined()
    } finally {
      await ctx.fiber.dispose()
    }
  })
})
