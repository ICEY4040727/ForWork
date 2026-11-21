/**
 * Contract: the Host save boundary refuses only an edit that intersects a
 * coupled construct. Preserved content the writer can carry — macros, tracked
 * changes, charts, comments elsewhere — never refuses a save by itself, and a
 * refusal names the construct, leaves the source byte-identical, and appends no
 * save fact.
 */
import { createHash } from 'node:crypto'
import { copyFile, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { strToU8, unzipSync, zipSync } from 'fflate'
import { Context } from '@forwork/cordis'
import { LocalFileSystem } from '@forwork/fs-local'
import SessionStore, { SessionId } from '@forwork/session'
import type { SessionEvent } from '@forwork/session'
import { documentIdFor, foldDocumentProjection } from '@forwork/session-document'
import type { DocumentProjectionState } from '@forwork/session-document'
import type { JsonValue } from '@forwork/util-values'
import { afterEach, beforeEach, expect, it } from 'vitest'
import LocalOfficeExchange, { type OfficeDocumentSnapshot } from '../src/index.ts'

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url))
const scope = { sessionId: SessionId('office-write-postcondition'), workspaceRoot: '' }
const signal = new AbortController().signal

let ctx: Context
let root: string
const events: SessionEvent[] = []

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'dsh-office-write-postcondition-'))
  scope.workspaceRoot = root
  events.length = 0
  ctx = new Context()
  await ctx.plugin(LocalFileSystem, { cwd: root })
  await ctx.plugin(SessionStore)
  const session = ctx.sessions.create(scope.sessionId, { meta: { cwd: root } })
  ctx.on('session/event', (subject, event) => {
    if (subject === session) events.push(event)
  })
  await ctx.plugin(LocalOfficeExchange)
})

afterEach(async () => {
  await ctx.fiber.dispose()
  await rm(root, { recursive: true, force: true })
})

async function copyFixture(name: string): Promise<string> {
  await copyFile(join(FIXTURES, name), join(root, name))
  return name
}

function hashOf(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

function jsonObject(value: JsonValue | undefined): Record<string, JsonValue> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) throw new Error('expected a JSON object')
  return value
}

function documentBody(snapshot: OfficeDocumentSnapshot): Record<string, JsonValue> {
  if (snapshot.kind !== 'document') throw new Error('expected a document snapshot')
  return jsonObject(jsonObject(snapshot.data)['body'])
}

function documentStream(snapshot: OfficeDocumentSnapshot): string {
  const stream = documentBody(snapshot)['dataStream']
  if (typeof stream !== 'string') throw new Error('expected a document data stream')
  return stream
}

function entryOf(bytes: Uint8Array, name: string): Uint8Array {
  const entry = unzipSync(bytes)[name]
  if (entry === undefined) throw new Error(`package has no ${name}`)
  return entry
}

function openedDocuments(): DocumentProjectionState {
  let state: DocumentProjectionState = { order: [], documents: {} }
  for (const event of events) state = foldDocumentProjection(state, event)
  return state
}

function saveFacts(): readonly SessionEvent[] {
  return events.filter(event => event.type === 'document/save')
}

function contentStartOf(stream: string, breakIndex: number): number {
  let index = breakIndex
  while (index > 0) {
    const previous = stream[index - 1]
    if (previous === '\r' || previous === '\n' || previous === '\x0F' || previous === '\x1C') break
    index -= 1
  }
  return index
}

/** Replace one text range and shift the model offsets the edit moves. */
function spliceStory(body: Record<string, JsonValue>, start: number, end: number, text: string): void {
  const stream = typeof body['dataStream'] === 'string' ? body['dataStream'] : ''
  const delta = text.length - (end - start)
  const deltaStart = start + text.length
  body['dataStream'] = stream.slice(0, start) + text + stream.slice(end)
  const paragraphs = Array.isArray(body['paragraphs']) ? body['paragraphs'] : []
  for (const entry of paragraphs) {
    const paragraph = jsonObject(entry)
    const index = paragraph['startIndex']
    if (typeof index === 'number' && index >= end) paragraph['startIndex'] = index + delta
  }
  const sections = Array.isArray(body['sectionBreaks']) ? body['sectionBreaks'] : []
  for (const entry of sections) {
    const section = jsonObject(entry)
    const index = section['startIndex']
    if (typeof index === 'number' && index >= end) section['startIndex'] = index + delta
  }
  const tables = Array.isArray(body['tables']) ? body['tables'] : []
  for (const entry of tables) {
    const table = jsonObject(entry)
    const startIndex = table['startIndex']
    const endIndex = table['endIndex']
    if (typeof startIndex === 'number' && startIndex >= end) table['startIndex'] = startIndex + delta
    if (typeof endIndex === 'number' && endIndex > end) table['endIndex'] = endIndex + delta
  }
  const blocks = Array.isArray(body['docxRawBlocks']) ? body['docxRawBlocks'] : []
  for (const entry of blocks) {
    const block = jsonObject(entry)
    const index = block['startIndex']
    if (typeof index === 'number' && index >= end) block['startIndex'] = index + delta
  }
  const runs = Array.isArray(body['textRuns']) ? body['textRuns'] : []
  body['textRuns'] = runs.flatMap((entry) => {
    const run = jsonObject(entry)
    const runStart = run['st']
    const runEnd = run['ed']
    if (typeof runStart !== 'number' || typeof runEnd !== 'number') return [entry]
    if (runEnd <= start) return [entry]
    if (runStart >= end) return [{ ...run, st: runStart + delta, ed: runEnd + delta }]
    const clippedStart = runStart < start ? runStart : start
    const clippedEnd = runEnd > end ? runEnd + delta : deltaStart
    return clippedEnd > clippedStart ? [{ ...run, st: clippedStart, ed: clippedEnd }] : []
  })
}

/** Replace one fixture's text and return the edited snapshot. */
function replaceText(snapshot: OfficeDocumentSnapshot, from: string, to: string): OfficeDocumentSnapshot {
  const body = documentBody(snapshot)
  const start = documentStream(snapshot).indexOf(from)
  if (start < 0) throw new Error(`fixture text "${from}" is absent`)
  spliceStory(body, start, start + from.length, to)
  return snapshot
}

/** Insert text at one paragraph's content start and return the edited snapshot. */
function insertAtParagraph(snapshot: OfficeDocumentSnapshot, paragraphIndex: number, text: string): OfficeDocumentSnapshot {
  const body = documentBody(snapshot)
  const paragraphs = Array.isArray(body['paragraphs']) ? body['paragraphs'] : []
  const paragraph = jsonObject(paragraphs[paragraphIndex])
  const startIndex = paragraph['startIndex']
  if (typeof startIndex !== 'number') throw new Error(`fixture has no paragraph ${paragraphIndex}`)
  const at = contentStartOf(documentStream(snapshot), startIndex)
  spliceStory(body, at, at, text)
  return snapshot
}

/** Write a minimal document whose paragraph carries a simple field. */
async function writeFieldFixture(name: string): Promise<string> {
  const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const packageRel = 'http://schemas.openxmlformats.org/package/2006/relationships'
  const contentTypes = 'http://schemas.openxmlformats.org/package/2006/content-types'
  const bytes = zipSync({
    '[Content_Types].xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<Types xmlns="${contentTypes}">`
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + '</Types>'),
    '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<Relationships xmlns="${packageRel}"><Relationship Id="rId1" Type="${rel}/officeDocument"`
      + ' Target="word/document.xml"/></Relationships>'),
    'word/_rels/document.xml.rels': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<Relationships xmlns="${packageRel}"><Relationship Id="rId1" Type="${rel}/styles"`
      + ' Target="styles.xml"/></Relationships>'),
    'word/styles.xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:styles xmlns:w="${word}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal">`
      + '<w:name w:val="Normal"/></w:style></w:styles>'),
    'word/document.xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:document xmlns:w="${word}" xmlns:r="${rel}"><w:body>`
      + '<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>'
      + '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>'
      + '<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>'
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>'),
  }, { level: 6 })
  await writeFile(join(root, name), bytes)
  return name
}

it('saves a macro-bearing document whose edit leaves the macro untouched', async () => {
  const path = await copyFixture('macro.docm')
  const read = await ctx.officeDocument.read(scope, path, signal)
  expect(read.unsupported.map(feature => feature.kind)).toContain('macro')
  const source = await readFile(join(root, path))
  replaceText(read.snapshot, 'Quarterly', 'Financial')

  const saved = await ctx.officeDocument.save(scope, path, read.snapshot, 'docm', read.version, { actor: 'user' }, signal)
  expect(saved.version).not.toBe(read.version)

  const written = await readFile(join(root, path))
  expect(entryOf(written, 'word/vbaProject.bin')).toEqual(entryOf(source, 'word/vbaProject.bin'))
  const writtenTypes = new TextDecoder().decode(entryOf(written, '[Content_Types].xml'))
  expect(writtenTypes).toContain('application/vnd.ms-word.document.macroEnabled.main+xml')
  expect(new TextDecoder().decode(entryOf(written, 'word/document.xml'))).toContain('Financial')
  expect(saveFacts()).toHaveLength(1)
  expect(openedDocuments().documents[documentIdFor(scope.sessionId, path)]?.contentVersion).toBe(1)
})

it('saves a tracked-change document whose edit leaves the tracked paragraph untouched', async () => {
  const path = await copyFixture('unsupported-tracked.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  expect(read.unsupported.map(feature => feature.kind)).toContain('tracked-changes')
  const sourceDocument = new TextDecoder().decode(entryOf(await readFile(join(root, path)), 'word/document.xml'))
  const trackedParagraph = /<w:p>\s*<w:ins[\s\S]*?<\/w:p>/u.exec(sourceDocument)?.[0]
  if (trackedParagraph === undefined) throw new Error('fixture has no tracked paragraph')
  replaceText(read.snapshot, 'Quarterly', 'Financial')

  await ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, { actor: 'user' }, signal)

  const writtenDocument = new TextDecoder().decode(entryOf(await readFile(join(root, path)), 'word/document.xml'))
  expect(writtenDocument).toContain(trackedParagraph)
  expect(writtenDocument).toContain('Financial Report')
  expect(saveFacts()).toHaveLength(1)
})

it('refuses a tracked change that overlaps the edited paragraph', async () => {
  const path = await copyFixture('unsupported-tracked.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  const before = await readFile(join(root, path))
  const declared = documentBody(read.snapshot)['paragraphs']
  const count = Array.isArray(declared) ? declared.length : 0
  insertAtParagraph(read.snapshot, count - 1, 'Edited')

  await expect(ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, { actor: 'user' }, signal))
    .rejects.toMatchObject({
      code: 'office-document/unsupported-feature',
      message: expect.stringMatching(/w:ins[\s\S]*tracked change/u) as string,
    })

  expect(hashOf(await readFile(join(root, path)))).toBe(hashOf(before))
  expect(saveFacts()).toHaveLength(0)
  expect(openedDocuments().documents[documentIdFor(scope.sessionId, path)]?.contentVersion).toBe(0)
})

it('refuses a comment range that overlaps the edited text', async () => {
  const path = await copyFixture('preserve-comment.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  const before = await readFile(join(root, path))
  replaceText(read.snapshot, 'Commented', 'Reviewed')

  await expect(ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, { actor: 'user' }, signal))
    .rejects.toMatchObject({
      code: 'office-document/unsupported-feature',
      message: expect.stringMatching(/w:commentRangeStart[\s\S]*comment range/u) as string,
    })

  expect(hashOf(await readFile(join(root, path)))).toBe(hashOf(before))
  expect(saveFacts()).toHaveLength(0)
  expect(openedDocuments().documents[documentIdFor(scope.sessionId, path)]?.contentVersion).toBe(0)
})

it('refuses a content control that wraps the edited text', async () => {
  const path = await copyFixture('preserve-content-control.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  const before = await readFile(join(root, path))
  insertAtParagraph(read.snapshot, 0, 'Edited')

  await expect(ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, { actor: 'user' }, signal))
    .rejects.toMatchObject({
      code: 'office-document/unsupported-feature',
      message: expect.stringContaining('w:sdt') as string,
    })

  expect(hashOf(await readFile(join(root, path)))).toBe(hashOf(before))
  expect(saveFacts()).toHaveLength(0)
  expect(openedDocuments().documents[documentIdFor(scope.sessionId, path)]?.contentVersion).toBe(0)
})

it('refuses a field whose preserved anchor would not survive the edit', async () => {
  const path = await writeFieldFixture('field-anchor.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  const before = await readFile(join(root, path))
  const body = documentBody(read.snapshot)
  body['dataStream'] = ' after\r\n'
  const paragraphs = body['paragraphs']
  if (!Array.isArray(paragraphs) || paragraphs[0] === undefined) throw new Error('fixture has no paragraph')
  jsonObject(paragraphs[0])['startIndex'] = 6

  await expect(ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, { actor: 'user' }, signal))
    .rejects.toMatchObject({
      code: 'office-document/unsupported-feature',
      message: expect.stringMatching(/w:fldSimple[\s\S]*field markup/u) as string,
    })

  expect(hashOf(await readFile(join(root, path)))).toBe(hashOf(before))
  expect(saveFacts()).toHaveLength(0)
  expect(openedDocuments().documents[documentIdFor(scope.sessionId, path)]?.contentVersion).toBe(0)
})

it('saves a document whose header text is mapped into the snapshot', async () => {
  const path = await copyFixture('unsupported-header.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  expect(read.unsupported).toEqual([])
  if (read.snapshot.kind !== 'document') throw new Error('expected a document snapshot')
  const data = jsonObject(read.snapshot.data)
  expect(JSON.stringify(data['headers'])).toContain('Confidential header')

  const saved = await ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, { actor: 'user' }, signal)
  expect(saved.version).not.toBe(read.version)

  const reread = await ctx.officeDocument.read(scope, path, signal)
  expect(JSON.stringify(reread.snapshot)).toContain('Confidential header')
  expect(saveFacts()).toHaveLength(1)
})

it('appends no document/save for a refused save', async () => {
  const path = await copyFixture('preserve-comment.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  replaceText(read.snapshot, 'Commented text', 'Reviewed text')

  await expect(ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, { actor: 'user' }, signal))
    .rejects.toMatchObject({ code: 'office-document/unsupported-feature' })

  expect(events.filter(event => event.type === 'document/open')).toHaveLength(1)
  expect(saveFacts()).toHaveLength(0)
})

it('saves a supported source and advances the file and content versions', async () => {
  const path = await copyFixture('roundtrip.docx')
  const read = await ctx.officeDocument.read(scope, path, signal)
  expect(read.unsupported).toEqual([])
  const before = await readFile(join(root, path))
  if (read.snapshot.kind !== 'document') throw new Error('expected a document snapshot')
  const body = jsonObject(jsonObject(read.snapshot.data)['body'])
  body['dataStream'] = documentStream(read.snapshot).replace('Quarterly', 'Financial')

  const saved = await ctx.officeDocument.save(scope, path, read.snapshot, 'docx', read.version, { actor: 'user' }, signal)

  expect(saved.version).not.toBe(read.version)
  expect(hashOf(await readFile(join(root, path)))).not.toBe(hashOf(before))
  const reread = await ctx.officeDocument.read(scope, path, signal)
  expect(documentStream(reread.snapshot)).toContain('Financial Report')
  expect(openedDocuments().documents[documentIdFor(scope.sessionId, path)]?.contentVersion).toBe(1)
  expect(saveFacts()).toHaveLength(1)
})
