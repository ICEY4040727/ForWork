/** DOCX provider mapping: fixture semantics, round-trip, and structural failures. */
import { readFile } from 'node:fs/promises'
import { strToU8, unzipSync, zipSync } from 'fflate'
import { expect, it, vi } from 'vitest'
import {
  BooleanNumber,
  NamedStyleType,
  TextDecoration,
  validateDocumentStructure,
  type IDocumentData,
  type IParagraph,
} from '@univerjs/core'
import { readDocument, writeDocument } from '../src/provider/docx.ts'
import { OfficeExchangeError } from '../src/errors.ts'

vi.mock('fflate', async (importOriginal) => {
  const actual = await importOriginal<typeof import('fflate')>()
  return {
    ...actual,
    unzipSync: vi.fn(actual.unzipSync),
    zipSync: vi.fn(actual.zipSync),
  }
})

const FIXTURE = new URL('./fixtures/roundtrip.docx', import.meta.url)
const MACRO_FIXTURE = new URL('./fixtures/unsupported-macro.docx', import.meta.url)
const DOCUMENT_NAMESPACE = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'

function paragraphText(data: IDocumentData, paragraph: IParagraph): string {
  const stream = data.body?.dataStream ?? ''
  let start = 0
  for (let index = paragraph.startIndex - 1; index >= 0; index -= 1) {
    if (stream[index] === '\r') {
      start = index + 1
      break
    }
  }
  return stream.slice(start, paragraph.startIndex).replaceAll(/[\u0000-\u001F]/gu, '')
}

function paragraphTexts(data: IDocumentData): readonly string[] {
  return (data.body?.paragraphs ?? []).map(paragraph => paragraphText(data, paragraph))
}

function expectFixtureSemantics(data: IDocumentData): void {
  const body = data.body
  expect(body).toBeDefined()
  if (body === undefined) throw new Error('missing document body')
  const paragraphs = body.paragraphs ?? []
  const texts = paragraphTexts(data)
  const heading = paragraphs.find(paragraph => paragraph.styleId === 'Heading1')
  expect(heading?.paragraphStyle?.namedStyleType).toBe(NamedStyleType.HEADING_1)
  expect(texts).toContain('Quarterly Report')
  expect(texts).toContain('This paragraph has bold text for round-trip checks.')

  const runs = body.textRuns ?? []
  const boldRun = runs.find(run => body.dataStream.slice(run.st, run.ed) === 'bold text')
  expect(boldRun?.ts?.bl).toBe(BooleanNumber.TRUE)

  expect(body.tables).toHaveLength(1)
  const tableId = body.tables?.[0]?.tableId
  expect(tableId).toBe('table-1')
  const table = tableId === undefined ? undefined : data.tableSource?.[tableId]
  expect(table?.tableRows).toHaveLength(2)
  expect(table?.tableRows[0]?.tableCells).toHaveLength(2)
  expect(table?.tableRows[1]?.tableCells).toHaveLength(2)
  expect(table?.tableColumns).toHaveLength(2)
  expect(texts).toEqual(expect.arrayContaining(['Region', 'Revenue', 'North', '1200']))
  expect(body.dataStream).toContain('\x1C')
  expect(body.dataStream).toContain('\x0F')
  expect(validateDocumentStructure(data)).toEqual([])
}

function packageOf(entries: Readonly<Record<string, string | Uint8Array>>): Uint8Array {
  const files: Record<string, Uint8Array> = {}
  for (const [name, content] of Object.entries(entries)) {
    files[name] = typeof content === 'string' ? strToU8(content) : content
  }
  return zipSync(files)
}

function documentOf(documentXml: string): Uint8Array {
  return packageOf({
    '[Content_Types].xml': '<Types/>',
    'word/document.xml': documentXml,
  })
}

function snapshotOf(stream: string, body: Omit<NonNullable<IDocumentData['body']>, 'dataStream'> = {}): IDocumentData {
  return {
    id: 'test-document',
    documentStyle: {},
    body: { dataStream: stream, ...body },
  }
}

it('reads the fixture heading, bold run, and 2x2 table', async () => {
  const bytes = new Uint8Array(await readFile(FIXTURE))
  expectFixtureSemantics(await readDocument(bytes))
})

it('preserves supported semantics through a write/read round-trip', async () => {
  const original = await readDocument(new Uint8Array(await readFile(FIXTURE)))
  const encoded = await writeDocument(original)
  expect(encoded.byteLength).toBeGreaterThan(0)
  expectFixtureSemantics(await readDocument(encoded))
})

it('reads a macro-bearing document without executing or retaining macros', async () => {
  const data = await readDocument(new Uint8Array(await readFile(MACRO_FIXTURE)))
  expectFixtureSemantics(data)
  expect(JSON.stringify(data)).not.toContain('vba')
})

it('maps paragraph, run, underline, and heading style variants', async () => {
  const xml = `${DOCUMENT_NAMESPACE}<w:body><w:p>`
    + '<w:pPr><w:pStyle w:val="Custom"/></w:pPr>'
    + '<w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>plain</w:t></w:r>'
    + '<w:r><w:rPr><w:i w:val="false"/><w:u w:val="double"/></w:rPr><w:t>italic</w:t></w:r>'
    + '<w:r><w:rPr><w:b w:val="off"/></w:rPr><w:t>flagged</w:t></w:r>'
    + '<w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>none</w:t></w:r>'
    + '<w:r><w:rPr><w:strike/></w:rPr><w:t>struck</w:t></w:r>'
    + '<w:r><w:rPr><w:b/></w:rPr></w:r>'
    + '<w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r>'
    + '<w:r><w:rPr><w:b/><w:i/><w:u/></w:rPr><w:t>all</w:t></w:r>'
    + '<w:hyperlink><w:r><w:t>link</w:t></w:r></w:hyperlink>'
    + '</w:p></w:body></w:document>'
  const data = await readDocument(documentOf(xml))
  const body = data.body
  expect(body?.paragraphs?.[0]?.styleId).toBe('Custom')
  expect(body?.paragraphs?.[0]?.paragraphStyle).toBeUndefined()
  expect(body?.dataStream).toContain('plainitalicflaggednonestruckA\tBall')
  const runs = body?.textRuns ?? []
  const italic = runs.find(run => body?.dataStream.slice(run.st, run.ed) === 'italic')
  expect(italic?.ts?.ul?.t).toBe(TextDecoration.DOUBLE)
  expect(italic?.ts?.it).toBeUndefined()
  expect(runs.some(run => body?.dataStream.slice(run.st, run.ed) === 'plain')).toBe(false)
  expect(runs.some(run => body?.dataStream.slice(run.st, run.ed) === 'flagged')).toBe(false)
  expect(runs.some(run => body?.dataStream.slice(run.st, run.ed) === 'none')).toBe(false)
  expect(runs.some(run => body?.dataStream.slice(run.st, run.ed) === 'struck')).toBe(false)
  expect(runs.some(run => body?.dataStream.slice(run.st, run.ed) === 'A\tB')).toBe(false)
  const all = runs.find(run => body?.dataStream.slice(run.st, run.ed) === 'all')
  expect(all?.ts?.bl).toBe(BooleanNumber.TRUE)
  expect(all?.ts?.it).toBe(BooleanNumber.TRUE)
  expect(all?.ts?.ul).toEqual({ s: BooleanNumber.TRUE, t: TextDecoration.SINGLE })
  expect(body?.dataStream).not.toContain('link')
})

it('reads an empty body as one empty paragraph with a section break', async () => {
  const data = await readDocument(documentOf(`${DOCUMENT_NAMESPACE}<w:body/></w:document>`))
  expect(data.body?.dataStream).toBe('\r\n')
  expect(data.body?.paragraphs).toEqual([{ startIndex: 0, paragraphId: 'paragraph-1' }])
  expect(data.body?.sectionBreaks).toEqual([{ sectionId: 'section-1', startIndex: 1 }])
  expect(data.body?.tables).toEqual([])
})

it('reads a leading table, empty cells, and nested tables without crashing', async () => {
  const nested = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const xml = `${DOCUMENT_NAMESPACE}<w:body><w:tbl>`
    + '<w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="auto"/><w:unknownGrid/></w:tblGrid>'
    + '<w:tr><w:trPr/>'
    + '<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>'
    + '<w:p><w:r><w:t>A</w:t></w:r></w:p>'
    + nested
    + '</w:tc>'
    + '<w:tc><w:tcPr/></w:tc>'
    + '<w:tc></w:tc>'
    + '<w:tc><w:tcPr><w:tcW w:type="dxa"/></w:tcPr></w:tc>'
    + '</w:tr></w:tbl></w:body></w:document>'
  const data = await readDocument(documentOf(xml))
  const body = data.body
  expect(body?.dataStream).not.toContain('nested')
  expect(body?.paragraphs).toHaveLength(6)
  expect(body?.dataStream.startsWith('\r\x1A\x1B\x1C')).toBe(true)
  const table = data.tableSource?.['table-1']
  expect(table?.tableRows[0]?.tableCells[0]?.size?.width.v).toBe(100)
  expect(table?.tableRows[0]?.tableCells[1]?.size).toBeUndefined()
  expect(table?.tableRows[0]?.tableCells[2]?.size).toBeUndefined()
  expect(table?.tableRows[0]?.tableCells[3]?.size).toBeUndefined()
  expect(table?.tableColumns.map(column => column.size.width.v)).toEqual([80, 160, 160, 160])
})

it('reads a table with no rows or grid columns', async () => {
  const data = await readDocument(documentOf(`${DOCUMENT_NAMESPACE}<w:body><w:tbl/></w:body></w:document>`))
  expect(data.body?.tables).toHaveLength(1)
  expect(data.tableSource?.['table-1']?.tableRows).toEqual([])
  expect(data.tableSource?.['table-1']?.tableColumns).toHaveLength(1)
})

it('writes heading, run, tab, and control-character content back to OOXML', async () => {
  const stream = 'abcde\rfgh\r\n'
  const data: IDocumentData = {
    id: 'runs',
    documentStyle: {},
    body: {
      dataStream: stream,
      paragraphs: [
        { startIndex: 5, paragraphId: 'p1', styleId: 'Custom' },
        { startIndex: 9, paragraphId: 'p2', paragraphStyle: { namedStyleType: NamedStyleType.HEADING_1 } },
      ],
      textRuns: [
        { st: 1, ed: 2, ts: { bl: BooleanNumber.TRUE, it: BooleanNumber.TRUE } },
        { st: 2, ed: 3, ts: { it: BooleanNumber.TRUE } },
        { st: 3, ed: 4, ts: { ul: { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE } } },
        { st: 4, ed: 5, ts: { ul: { s: BooleanNumber.TRUE, t: TextDecoration.DOUBLE } } },
        { st: 6, ed: 7, ts: {} },
      ],
    },
  }
  const reopened = await readDocument(await writeDocument(data))
  expect(paragraphTexts(reopened)).toEqual(['abcde', 'fgh'])
  const runs = reopened.body?.textRuns ?? []
  const styled = runs.find(run => reopened.body?.dataStream.slice(run.st, run.ed) === 'b')
  expect(styled?.ts?.bl).toBe(BooleanNumber.TRUE)
  expect(styled?.ts?.it).toBe(BooleanNumber.TRUE)
  expect(runs.find(run => reopened.body?.dataStream.slice(run.st, run.ed) === 'c')?.ts?.it)
    .toBe(BooleanNumber.TRUE)
  expect(runs.find(run => reopened.body?.dataStream.slice(run.st, run.ed) === 'd')?.ts?.ul?.t)
    .toBe(TextDecoration.SINGLE)
  expect(runs.find(run => reopened.body?.dataStream.slice(run.st, run.ed) === 'e')?.ts?.ul?.t)
    .toBe(TextDecoration.DOUBLE)

  const tabData = snapshotOf('a\tb\u0001c\r\n', {
    paragraphs: [{ startIndex: 5, paragraphId: 'p1' }],
  })
  const tabReopened = await readDocument(await writeDocument(tabData))
  expect(tabReopened.body?.dataStream).toBe('a\tbc\r\n')
})

it('writes paragraphs that carry only a named style and normal style', async () => {
  const data: IDocumentData = {
    id: 'styled',
    documentStyle: {},
    body: {
      dataStream: 'Title\rBody\r\n',
      paragraphs: [
        { startIndex: 5, paragraphId: 'p1', paragraphStyle: { namedStyleType: NamedStyleType.HEADING_1 } },
        { startIndex: 10, paragraphId: 'p2', paragraphStyle: { namedStyleType: NamedStyleType.NORMAL_TEXT } },
      ],
    },
  }
  const reopened = await readDocument(await writeDocument(data))
  expect(reopened.body?.paragraphs?.[0]?.styleId).toBe('Heading1')
  expect(reopened.body?.paragraphs?.[0]?.paragraphStyle?.namedStyleType).toBe(NamedStyleType.HEADING_1)
  expect(reopened.body?.paragraphs?.[1]?.styleId).toBeUndefined()
})

it('writes a table without table metadata using default column widths', async () => {
  const original = await readDocument(new Uint8Array(await readFile(FIXTURE)))
  const dataWithoutSource: IDocumentData = {
    id: original.id,
    documentStyle: original.documentStyle,
    body: original.body,
  }
  const reopened = await readDocument(await writeDocument(dataWithoutSource))
  const table = reopened.tableSource?.['table-1']
  expect(table?.tableRows).toHaveLength(2)
  expect(table?.tableColumns.map(column => column.size.width.v)).toEqual([160, 160])
})

it('writes an empty document body', async () => {
  const data: IDocumentData = { id: 'empty', documentStyle: {} }
  const reopened = await readDocument(await writeDocument(data))
  expect(reopened.body?.dataStream).toBe('\r\n')
})

it('skips nested table tokens inside a cell when writing', async () => {
  const stream = '\x1A\x1B\x1Couter\r\n\x1A\x1B\x1Cinner\r\n\x1D\x0E\x0F\x1D\x0E\x0F'
  const data = snapshotOf(stream, {
    tables: [{ startIndex: 0, endIndex: stream.length, tableId: 'table-1' }],
  })
  const reopened = await readDocument(await writeDocument(data))
  expect(paragraphTexts(reopened)).toEqual(['', 'outer', ''])
  expect(reopened.body?.dataStream).not.toContain('inner')
})

it('writes an empty paragraph for a cell without a paragraph break', async () => {
  const stream = '\x1A\x1B\x1C\n\x1D\x0E\x0F'
  const data = snapshotOf(stream, {
    tables: [{ startIndex: 0, endIndex: stream.length, tableId: 'table-1' }],
  })
  const reopened = await readDocument(await writeDocument(data))
  expect(reopened.body?.tables).toHaveLength(1)
  expect(reopened.tableSource?.['table-1']?.tableRows).toHaveLength(1)
})

it('rejects garbage bytes as a conversion failure', async () => {
  await expect(readDocument(new Uint8Array([0, 1, 2, 3]))).rejects.toMatchObject({
    name: 'OfficeExchangeError',
    code: 'conversion-failed',
  })
})

it('rejects bytes above the configured file bound', async () => {
  await expect(readDocument(new Uint8Array(9), {
    maxFileBytes: 8,
    maxCells: 100,
    maxConversionMs: 1_000,
  })).rejects.toMatchObject({ code: 'too-large' })
})

it('rejects a document above the configured table-cell bound', async () => {
  const bytes = new Uint8Array(await readFile(FIXTURE))
  await expect(readDocument(bytes, {
    maxFileBytes: 1_000_000,
    maxCells: 1,
    maxConversionMs: 1_000,
  })).rejects.toMatchObject({ code: 'too-large' })
})

it('rejects a package without a document part', async () => {
  const bytes = packageOf({ 'word/styles.xml': '<w:styles/>' })
  await expect(readDocument(bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('rejects malformed document XML', async () => {
  const bytes = documentOf('<w:document><w:body><w:p></w:body></w:document>')
  await expect(readDocument(bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('rejects document XML without a document or body element', async () => {
  await expect(readDocument(documentOf('<root><body/></root>'))).rejects.toMatchObject({
    code: 'conversion-failed',
  })
  await expect(readDocument(documentOf('<w:document/>'))).rejects.toMatchObject({
    code: 'conversion-failed',
  })
})

it('rejects packages with too many entries', async () => {
  const entries: Record<string, Uint8Array> = {}
  for (let index = 0; index <= 10_000; index += 1) {
    entries[`word/media/part${String(index)}.xml`] = new Uint8Array(0)
  }
  await expect(readDocument(zipSync(entries))).rejects.toMatchObject({ code: 'too-large' })
})

it('reports non-Error package read failures with their text', async () => {
  vi.mocked(unzipSync).mockImplementationOnce(() => {
    throw { toString: () => 'read rejected' }
  })
  const error: unknown = await readDocument(new Uint8Array([1])).catch((cause: unknown) => cause)
  expect(error).toMatchObject({ code: 'conversion-failed' })
  expect(String(error)).toContain('read rejected')
})

it('rejects text outside a paragraph and malformed table streams', async () => {
  const cases = [
    'stray text',
    '\x1A\x1B\x1Ctext',
    '\x1A\x1B\x1Ctext\r\n\x1D',
    '\x1A\x1B\x1Ctext\r\n\x1D\x0E',
    '\x1A\x1B\x1Ctext\n\x1D\x0E\x0F',
  ]
  for (const stream of cases) {
    await expect(writeDocument(snapshotOf(stream))).rejects.toMatchObject({ code: 'conversion-failed' })
  }
})

it('rejects a snapshot above the configured table-cell bound before writing', async () => {
  const data = await readDocument(new Uint8Array(await readFile(FIXTURE)))
  await expect(writeDocument(data, {
    maxFileBytes: 1_000_000,
    maxCells: 1,
    maxConversionMs: 1_000,
  })).rejects.toMatchObject({ code: 'too-large' })
})

it('reports document serialization failures', async () => {
  vi.mocked(zipSync).mockImplementationOnce(() => {
    throw new Error('write rejected')
  })
  await expect(writeDocument(snapshotOf('text\r\n'))).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('reports non-Error serialization failures with their text', async () => {
  vi.mocked(zipSync).mockImplementationOnce(() => {
    throw { toString: () => 'write rejected' }
  })
  const error: unknown = await writeDocument(snapshotOf('text\r\n')).catch((cause: unknown) => cause)
  expect(error).toMatchObject({ code: 'conversion-failed' })
  expect(String(error)).toContain('write rejected')
})

it('reports every document failure as an OfficeExchangeError', async () => {
  const error: unknown = await readDocument(new Uint8Array([1, 2, 3])).catch((cause: unknown) => cause)
  expect(error).toBeInstanceOf(OfficeExchangeError)
})
