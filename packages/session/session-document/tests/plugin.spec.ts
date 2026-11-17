import { describe, expect, it } from 'vitest'
import { Context } from '@forwork/cordis'
import type { Fiber } from '@forwork/cordis'
import { assembleContextFor } from '@forwork/agent'
import type { Agent } from '@forwork/agent'
import { SessionId } from '@forwork/session'
import { mountAgentLoopTestDependencies, mountAgentLoopTestHarness } from '@forwork/agent-loop-testkit'
import { documentIdFor } from '../src/document-identity.ts'
import * as DocumentPlugin from '../src/index.ts'
import { renderOpenDocuments } from '../src/plugin.ts'
import type { DocumentProjectionState } from '../src/types.ts'

const SESSION = SessionId('document-plugin')
const FIRST = documentIdFor(SESSION, 'work/first.xlsx')
const SECOND = documentIdFor(SESSION, 'work/second.docx')
const SECTION = 'document:open-set'

async function mount(): Promise<{ ctx: Context; agent: Agent; fiber: Fiber }> {
  const ctx = new Context()
  await mountAgentLoopTestDependencies(ctx)
  const loop = await mountAgentLoopTestHarness(ctx)
  const agent = await loop.create(SESSION)
  const fiber = await ctx.plugin(DocumentPlugin)
  return { ctx, agent, fiber }
}

/** The rendered text of the open-document section, or `undefined` when it is absent. */
async function sectionText(ctx: Context, context = {}): Promise<string | undefined> {
  const assembly = await ctx.systemPrompt.assemble(context)
  return assembly.sections.find(section => section.name === SECTION)?.text
}

/** One open-document state with every listed path open, in the given order. */
function stateOf(
  entries: readonly {
    readonly path: string
    readonly format: 'docx' | 'xlsx' | 'pptx'
    readonly status?: 'open' | 'missing'
  }[],
): DocumentProjectionState {
  const documents = Object.fromEntries(entries.map(({ path, format, status = 'open' }) => {
    const documentId = documentIdFor(SESSION, path)
    return [documentId, {
      documentId,
      path,
      format,
      contentVersion: 0,
      status,
      unsupported: [],
    }]
  }))
  return { order: Object.keys(documents), documents }
}

describe('session-document plugin', () => {
  it('renders the open set with an addressable id for the assembled agent session', async () => {
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

      expect(await sectionText(ctx, assembleContextFor(agent))).toMatchInlineSnapshot(`
        "Open documents in this session:
        1. document-plugin:work/first.xlsx (first.xlsx, xlsx)
        2. document-plugin:work/second.docx (second.docx, docx, missing)"
      `)
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

  it('tells two same-named files apart by their full ids', () => {
    expect(renderOpenDocuments(stateOf([
      { path: 'docs/report.docx', format: 'docx' },
      { path: 'archive/report.docx', format: 'docx' },
    ]))).toBe([
      'Open documents in this session:',
      '1. document-plugin:docs/report.docx (report.docx, docx)',
      '2. document-plugin:archive/report.docx (report.docx, docx)',
    ].join('\n'))
  })

  it('elides a long file name in the readable annotation only', () => {
    const path = `work/${'a'.repeat(50)}.xlsx`
    expect(renderOpenDocuments(stateOf([{ path, format: 'xlsx' }]))).toBe([
      'Open documents in this session:',
      `1. document-plugin:${path} (${'a'.repeat(37)}..., xlsx)`,
    ].join('\n'))
  })

  it('keeps the missing marker on an entry', () => {
    expect(renderOpenDocuments(stateOf([
      { path: 'work/report.docx', format: 'docx', status: 'missing' },
    ]))).toBe([
      'Open documents in this session:',
      '1. document-plugin:work/report.docx (report.docx, docx, missing)',
    ].join('\n'))
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
