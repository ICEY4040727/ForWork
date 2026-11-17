/**
 * Office document tools: `document_read`, `document_edit_cells`,
 * `document_edit_paragraph`, and `document_convert`.
 *
 * The tools are Host-owned. An edit validates the Session, the workspace root,
 * and the caller's version, then dispatches one correlated command to the live
 * editor; the editor applies the operation and the Host performs the single
 * version-guarded save. Without a live editor the call refuses instead of
 * editing a detached copy.
 * @module @forwork/tool-office
 */

import type {} from '@forwork/agent'
import type { Context } from '@forwork/cordis'
import z from '@forwork/schemastery'
import type { DocumentCommandRequest, OfficeCommandOutcome } from '@forwork/office-exchange'
import type { Session } from '@forwork/session'
import { defineTool } from '@forwork/tools'
import type { ToolDefinition } from '@forwork/tools'
import type { JsonValue } from '@forwork/util-values'
import { convertDocument } from './convert.ts'
import type { ConvertOutput } from './convert.ts'
import { editCells } from './edit-cells.ts'
import type { CellEditInput } from './edit-cells.ts'
import { editParagraphs } from './edit-paragraph.ts'
import type { ParagraphEditInput } from './edit-paragraph.ts'
import type { EditOutput } from './edit.ts'
import { failure } from './errors.ts'
import type { OfficeToolResult } from './errors.ts'
import { readDocument } from './read.ts'
import type { DocumentReadOutput } from './read.ts'
import type { DocumentToolCall, DocumentToolConfig, DocumentToolRuntime } from './runtime.ts'

/** Cordis plugin name used by loader diagnostics. */
export const name = 'tool-office'

/** Services required by the Office document tools. */
export const inject = ['agents', 'tools', 'officeDocument']

/** Deployment caps for the Office document tools. */
export interface Config {
  /** Largest cell preview a structure read lists; omitted uses 200. */
  summaryCells?: number
  /** Largest single edit operation; omitted uses 2000. */
  maxEditSteps?: number
}

/** Validated caps applied to every tool call. */
export const Config: z<Config> = z.object({
  summaryCells: z.natural().min(1).default(200),
  maxEditSteps: z.natural().min(1).default(2_000),
})

/** Accepted output of any document tool. */
export type DocumentToolOutput =
  | DocumentReadOutput
  | EditOutput<CellEditInput>
  | EditOutput<ParagraphEditInput>
  | ConvertOutput

/** One invocation after the tool schema validated its arguments. */
export type DocumentToolInvocation =
  | { readonly name: 'document_read'; readonly documentId: string }
  | {
    readonly name: 'document_edit_cells'
    readonly documentId: string
    readonly fromVersion: number
    readonly operations: readonly CellEditInput[]
  }
  | {
    readonly name: 'document_edit_paragraph'
    readonly documentId: string
    readonly fromVersion: number
    readonly operations: readonly ParagraphEditInput[]
  }
  | { readonly name: 'document_convert'; readonly path: string }

/** The narrow execution identity a document tool needs. */
export interface DocumentToolExecutionContext {
  /** Calling agent; only its Session is read here. */
  readonly agent?: { readonly session: Session }
  /** Caller cancellation. */
  readonly signal: AbortSignal
}

/** The document id convention shared with the editor's draft identity. */
const DOCUMENT_ID_HELP = 'Session-scoped document id in the form "<sessionId>:<path>", as returned by document_read.'

const CELL_OPERATION_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    sheetId: { type: 'string', required: true, description: 'Sheet id from document_read.' },
    row: { type: 'integer', required: true, description: 'Zero-based row index.' },
    column: { type: 'integer', required: true, description: 'Zero-based column index.' },
    value: { type: 'json', description: 'New JSON cell value; omit when setting a formula.' },
    formula: { type: 'string', description: 'New formula text; omit when setting a value.' },
  },
} as const

const PARAGRAPH_OPERATION_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    index: { type: 'integer', required: true, description: 'Zero-based paragraph index from document_read.' },
    text: { type: 'string', required: true, description: 'Paragraph text after the edit.' },
  },
} as const

const DOCUMENT_ID_PARAMETER = {
  type: 'string',
  required: true,
  description: DOCUMENT_ID_HELP,
} as const

const FROM_VERSION_PARAMETER = {
  type: 'integer',
  required: true,
  description: 'Document content version returned by document_read or the previous edit.',
} as const

const OUTPUT = {
  schema: { type: 'json' } as const,
  render: (_args: unknown, value: JsonValue) => [{ type: 'text' as const, text: JSON.stringify(value) }],
}

/**
 * Execute one document tool invocation for a calling Session.
 * @param invocation - validated tool name and arguments.
 * @param runtime - Host services consumed by the operation.
 * @param config - deployment caps.
 * @param exec - caller identity and cancellation.
 * @returns the structured result or refusal.
 * @throws when no calling agent exists or the Session has no workspace root.
 */
export async function executeDocumentTool(
  invocation: DocumentToolInvocation,
  runtime: DocumentToolRuntime,
  config: DocumentToolConfig,
  exec: DocumentToolExecutionContext,
): Promise<OfficeToolResult<DocumentToolOutput>> {
  const session = exec.agent?.session
  if (session === undefined) throw new Error('Office document tools require a calling agent')
  const workspaceRoot = session.header.cwd
  if (workspaceRoot === undefined) {
    return failure('out-of-scope', 'The Session has no workspace root, so no document path can be authorized.')
  }
  const call: DocumentToolCall = {
    runtime,
    config,
    session,
    scope: { sessionId: session.id, workspaceRoot },
    signal: exec.signal,
  }
  switch (invocation.name) {
    case 'document_read':
      return await readDocument(call, invocation.documentId)
    case 'document_edit_cells':
      return await editCells(call, {
        documentId: invocation.documentId,
        fromVersion: invocation.fromVersion,
        operations: invocation.operations,
      })
    case 'document_edit_paragraph':
      return await editParagraphs(call, {
        documentId: invocation.documentId,
        fromVersion: invocation.fromVersion,
        operations: invocation.operations,
      })
    case 'document_convert':
      return await convertDocument(call, invocation.path)
  }
}

/** Execute one invocation and materialize its result as lossless JSON. */
async function runTool(
  invocation: DocumentToolInvocation,
  runtime: DocumentToolRuntime,
  config: DocumentToolConfig,
  exec: DocumentToolExecutionContext,
): Promise<JsonValue> {
  const text = JSON.stringify(await executeDocumentTool(invocation, runtime, config, exec))
  return JSON.parse(text) as JsonValue
}

/**
 * Register the four Office document tools on a tool registry.
 * @param tools - tool registry receiving the definitions.
 * @param runtime - Host services consumed by every definition.
 * @param config - deployment caps.
 * @returns the registered definitions, in registration order.
 */
export function registerDocumentTools(
  tools: { register(definition: ToolDefinition): void },
  runtime: DocumentToolRuntime,
  config: DocumentToolConfig,
): ToolDefinition[] {
  const definitions: ToolDefinition[] = [
    defineTool({
      name: 'document_read',
      description: 'Read the structure of an open Office document: sheets, cells, formulas, paragraphs, and tables.',
      parameters: { documentId: DOCUMENT_ID_PARAMETER },
      output: OUTPUT,
      async execute(args, exec) {
        return await runTool({ name: 'document_read', documentId: args.documentId }, runtime, config, exec)
      },
    }),
    defineTool({
      name: 'document_edit_cells',
      description: 'Replace cell values or formulas in an open spreadsheet, then save the result immediately.',
      parameters: {
        documentId: DOCUMENT_ID_PARAMETER,
        fromVersion: FROM_VERSION_PARAMETER,
        operations: {
          type: 'array',
          required: true,
          description: 'Cell replacements in order.',
          items: CELL_OPERATION_SCHEMA,
        },
      },
      output: OUTPUT,
      async execute(args, exec) {
        return await runTool({
          name: 'document_edit_cells',
          documentId: args.documentId,
          fromVersion: args.fromVersion,
          operations: args.operations,
        }, runtime, config, exec)
      },
    }),
    defineTool({
      name: 'document_edit_paragraph',
      description: 'Replace paragraphs in an open text document, then save the result immediately.',
      parameters: {
        documentId: DOCUMENT_ID_PARAMETER,
        fromVersion: FROM_VERSION_PARAMETER,
        operations: {
          type: 'array',
          required: true,
          description: 'Paragraph replacements in order.',
          items: PARAGRAPH_OPERATION_SCHEMA,
        },
      },
      output: OUTPUT,
      async execute(args, exec) {
        return await runTool({
          name: 'document_edit_paragraph',
          documentId: args.documentId,
          fromVersion: args.fromVersion,
          operations: args.operations,
        }, runtime, config, exec)
      },
    }),
    defineTool({
      name: 'document_convert',
      description: 'Convert a legacy .doc or .xls file to an editable .docx or .xlsx copy inside the workspace.',
      parameters: {
        path: {
          type: 'string',
          required: true,
          description: 'Workspace-relative or absolute path of the legacy file.',
        },
      },
      output: OUTPUT,
      async execute(args, exec) {
        return await runTool({ name: 'document_convert', path: args.path }, runtime, config, exec)
      },
    }),
  ]
  for (const definition of definitions) tools.register(definition)
  return definitions
}

/** Build the Host runtime backed by `ctx.officeDocument` and its command broker. */
function runtimeOf(ctx: Context): DocumentToolRuntime {
  const documents = ctx.get('officeDocument')
  if (documents === undefined) {
    throw new Error('tool-office: officeDocument is absent: mount @forwork/office-exchange')
  }
  const broker = ctx.get('officeCommandBroker')
  return {
    documents,
    commands: {
      execute: (request: DocumentCommandRequest, signal: AbortSignal): Promise<OfficeCommandOutcome> =>
        documents.dispatchCommand(request, signal),
      currentVersion: documentId => broker?.versionOf(documentId) ?? 0,
      remember: (documentId, version) => { broker?.remember(documentId, version) },
    },
    turnOf: (session) => {
      const boundary = ctx.agents.turnBoundaryOf(session)
      return boundary.openTurnStartSeq === null ? undefined : boundary.lastTurn
    },
  }
}

/**
 * Resolve deployment caps and reject values schemastery did not normalize.
 * @param config - optional caps supplied by the composition.
 * @returns complete validated caps.
 * @throws when a cap is not a positive safe integer.
 */
export function resolveToolConfig(config: Config): DocumentToolConfig {
  const summaryCells = config.summaryCells ?? 200
  const maxEditSteps = config.maxEditSteps ?? 2_000
  if (!Number.isSafeInteger(summaryCells) || summaryCells < 1) {
    throw new Error('tool-office: summaryCells must be a positive integer')
  }
  if (!Number.isSafeInteger(maxEditSteps) || maxEditSteps < 1) {
    throw new Error('tool-office: maxEditSteps must be a positive integer')
  }
  return { summaryCells, maxEditSteps }
}

/**
 * Register the Office document tools on the Host context.
 * @param ctx - Host context carrying the tool registry and `officeDocument`.
 * @param config - deployment caps; omitted fields keep their defaults.
 */
export function apply(ctx: Context, config: Config): void {
  registerDocumentTools(ctx.tools, runtimeOf(ctx), resolveToolConfig(config))
}
