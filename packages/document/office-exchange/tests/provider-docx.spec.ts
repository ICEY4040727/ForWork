/** DOCX provider mapping: fixture semantics, round-trip, and structural failures. */
import { readFile } from 'node:fs/promises'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { expect, it, vi } from 'vitest'
import {
  BooleanNumber,
  ColumnSeparatorType,
  NamedStyleType,
  PageOrientType,
  SectionType,
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

it('preserves per-section page setup and orientation', async () => {
  const xml = `${DOCUMENT_NAMESPACE.replace(
    '>',
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
  )}<w:body>`
    + '<w:p><w:r><w:t>landscape body</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:sectPr>'
    + '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
    + '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" w:header="708" w:footer="708"/>'
    + '</w:sectPr></w:pPr><w:r><w:t>A4 section</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t>letter body</w:t></w:r></w:p>'
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
    + '</w:body></w:document>'
  const data = await readDocument(documentOf(xml))
  const sections = data.body?.sectionBreaks ?? []
  expect(sections).toHaveLength(2)
  expect(sections[0]?.pageSize).toEqual({ width: 16838 / 15, height: 11906 / 15 })
  expect(sections[0]?.pageOrient).toBe(PageOrientType.LANDSCAPE)
  expect(sections[0]?.marginTop).toBe(48)
  expect(sections[0]?.marginRight).toBe(96)
  expect(sections[0]?.marginBottom).toBe(48)
  expect(sections[0]?.marginLeft).toBe(96)
  expect(sections[0]?.marginHeader).toBe(708 / 15)
  expect(sections[0]?.marginFooter).toBe(708 / 15)
  expect(sections[1]?.pageSize).toEqual({ width: 816, height: 1056 })
  expect(sections[1]?.pageOrient).toBe(PageOrientType.PORTRAIT)
  expect(sections[1]?.marginTop).toBe(96)
  expect(sections[1]?.marginLeft).toBe(96)
  expect(validateDocumentStructure(data)).toEqual([])

  const reopened = await readDocument(await writeDocument(data))
  const rewritten = reopened.body?.sectionBreaks ?? []
  expect(rewritten).toHaveLength(2)
  expect(rewritten[0]?.pageSize).toEqual(sections[0]?.pageSize)
  expect(rewritten[0]?.pageOrient).toBe(PageOrientType.LANDSCAPE)
  expect(rewritten[0]?.marginTop).toBe(48)
  expect(rewritten[0]?.marginRight).toBe(96)
  expect(rewritten[1]?.pageSize).toEqual({ width: 816, height: 1056 })
  expect(rewritten[1]?.marginTop).toBe(96)
  expect(rewritten[1]?.marginLeft).toBe(96)
  expect(validateDocumentStructure(reopened)).toEqual([])
})

it('maps header and footer text with their section references', async () => {
  const namespace = DOCUMENT_NAMESPACE.replace(
    '>',
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
  )
  const bytes = packageOf({
    'word/document.xml': `${namespace}<w:body>`
      + '<w:p><w:r><w:t>body text</w:t></w:r></w:p>'
      + '<w:sectPr>'
      + '<w:headerReference w:type="default" r:id="rId2"/>'
      + '<w:footerReference w:type="default" r:id="rId3"/>'
      + '<w:pgSz w:w="12240" w:h="15840"/>'
      + '</w:sectPr></w:body></w:document>',
    'word/_rels/document.xml.rels': '<Relationships>'
      + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"'
      + ' Target="header1.xml"/>'
      + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"'
      + ' Target="footer1.xml"/>'
      + '</Relationships>',
    'word/header1.xml': '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:p><w:r><w:t>Report header</w:t></w:r></w:p></w:hdr>',
    'word/footer1.xml': '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:p><w:r><w:t>Page footer</w:t></w:r></w:p></w:ftr>',
  })
  const data = await readDocument(bytes)
  const section = data.body?.sectionBreaks?.[0]
  expect(section?.defaultHeaderId).toBe('header1')
  expect(section?.defaultFooterId).toBe('footer1')
  expect(data.headers?.header1?.body.dataStream).toBe('Report header\r\n')
  expect(data.footers?.footer1?.body.dataStream).toBe('Page footer\r\n')
  expect(data.headers?.header1?.body.paragraphs).toEqual([{ startIndex: 13, paragraphId: 'paragraph-1' }])
  expect(validateDocumentStructure(data)).toEqual([])

  const reopened = await readDocument(await writeDocument(data))
  expect(reopened.body?.sectionBreaks?.[0]?.defaultHeaderId).toBe('header1')
  expect(reopened.body?.sectionBreaks?.[0]?.defaultFooterId).toBe('footer1')
  expect(reopened.headers?.header1?.body.dataStream).toBe('Report header\r\n')
  expect(reopened.footers?.footer1?.body.dataStream).toBe('Page footer\r\n')
  expect(validateDocumentStructure(reopened)).toEqual([])
})

it('maps first-page and even-page references, title pages, settings, and explicit columns', async () => {
  const namespace = DOCUMENT_NAMESPACE.replace(
    '>',
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
  )
  const bytes = packageOf({
    'word/document.xml': `${namespace}<w:body><w:p/>`
      + '<w:sectPr>'
      + '<w:headerReference w:type="first" r:id="rId2"/>'
      + '<w:headerReference w:type="even" r:id="rId3"/>'
      + '<w:footerReference w:type="first" r:id="rId4"/>'
      + '<w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>'
      + '<w:pgMar w:top="1000" w:left="2000"/>'
      + '<w:cols w:sep="0"><w:col w:w="3000" w:space="400"/><w:col w:w="3000"/></w:cols>'
      + '<w:titlePg/>'
      + '</w:sectPr></w:body></w:document>',
    'word/settings.xml': '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:evenAndOddHeaders/></w:settings>',
    'word/_rels/document.xml.rels': '<Relationships>'
      + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"'
      + ' Target="header1.xml"/>'
      + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"'
      + ' Target="header2.xml"/>'
      + '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"'
      + ' Target="footer1.xml"/>'
      + '</Relationships>',
    'word/header1.xml': '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:p><w:r><w:t>First header</w:t></w:r></w:p></w:hdr>',
    'word/header2.xml': '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:p><w:r><w:t>Even header</w:t></w:r></w:p></w:hdr>',
    'word/footer1.xml': '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:p><w:r><w:t>First footer</w:t></w:r></w:p></w:ftr>',
  })
  const data = await readDocument(bytes)
  const section = data.body?.sectionBreaks?.[0]
  expect(data.documentStyle.evenAndOddHeaders).toBe(BooleanNumber.TRUE)
  expect(section?.firstPageHeaderId).toBe('header1')
  expect(section?.evenPageHeaderId).toBe('header2')
  expect(section?.firstPageFooterId).toBe('footer1')
  expect(section?.defaultHeaderId).toBeUndefined()
  expect(section?.useFirstPageHeaderFooter).toBe(BooleanNumber.TRUE)
  expect(section?.pageOrient).toBe(PageOrientType.PORTRAIT)
  expect(section?.marginTop).toBe(1000 / 15)
  expect(section?.marginLeft).toBe(2000 / 15)
  expect(section?.marginRight).toBeUndefined()
  expect(section?.columnSeparatorType).toBe(ColumnSeparatorType.NONE)
  expect(section?.columnProperties).toEqual([
    { width: 200, paddingEnd: 400 / 15 },
    { width: 200, paddingEnd: 0 },
  ])

  const reopened = await readDocument(await writeDocument(data))
  const rewritten = reopened.body?.sectionBreaks?.[0]
  expect(reopened.documentStyle.evenAndOddHeaders).toBe(BooleanNumber.TRUE)
  expect(rewritten?.firstPageHeaderId).toBe('header1')
  expect(rewritten?.evenPageHeaderId).toBe('header2')
  expect(rewritten?.firstPageFooterId).toBe('footer1')
  expect(rewritten?.useFirstPageHeaderFooter).toBe(BooleanNumber.TRUE)
  expect(rewritten?.pageOrient).toBe(PageOrientType.PORTRAIT)
  expect(rewritten?.marginTop).toBe(1000 / 15)
  expect(rewritten?.marginLeft).toBe(2000 / 15)
  expect(rewritten?.columnProperties).toEqual(section?.columnProperties)
  const entries = unzipSync(await writeDocument(data))
  expect(Object.keys(entries)).toEqual(expect.arrayContaining([
    'word/header1.xml',
    'word/header2.xml',
    'word/footer1.xml',
    'word/settings.xml',
  ]))
})

it('ignores unusable section references and column declarations', async () => {
  const namespace = DOCUMENT_NAMESPACE.replace(
    '>',
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
  )
  const bytes = packageOf({
    'word/document.xml': `${namespace}<w:body>`
      + '<w:p><w:pPr><w:sectPr>'
      + '<w:headerReference w:type="default"/>'
      + '<w:headerReference w:type="default" r:id="rId99"/>'
      + '<w:headerReference w:type="default" r:id="rId2"/>'
      + '<w:pgSz w:w="12240"/>'
      + '<w:cols w:num="2"><w:col/><w:unknown/></w:cols>'
      + '</w:sectPr></w:pPr><w:r><w:t>one</w:t></w:r></w:p>'
      + '<w:p><w:pPr><w:sectPr><w:cols w:sep="1"/></w:sectPr></w:pPr><w:r><w:t>two</w:t></w:r></w:p>'
      + '<w:p><w:pPr><w:sectPr><w:pgSz w:w="15840" w:h="12240"/>'
      + '<w:cols w:num="2" w:space="20000"/></w:sectPr></w:pPr><w:r><w:t>three</w:t></w:r></w:p>'
      + '<w:p><w:r><w:t>four</w:t></w:r></w:p>'
      + '<w:sectPr><w:pgSz w:w="15840" w:h="12240"/><w:cols w:num="2"/></w:sectPr>'
      + '</w:body></w:document>',
    'word/_rels/document.xml.rels': '<Relationships>'
      + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"'
      + ' Target="headerMissing.xml"/>'
      + '</Relationships>',
  })
  const sections = (await readDocument(bytes)).body?.sectionBreaks ?? []
  expect(sections).toHaveLength(4)
  expect(sections[0]?.pageSize).toBeUndefined()
  expect(sections[0]?.defaultHeaderId).toBeUndefined()
  expect(sections[0]?.columnProperties).toBeUndefined()
  expect(sections[1]?.columnSeparatorType).toBe(ColumnSeparatorType.BETWEEN_EACH_COLUMN)
  expect(sections[1]?.columnProperties).toBeUndefined()
  expect(sections[2]?.columnProperties).toBeUndefined()
  expect(sections[3]?.pageOrient).toBe(PageOrientType.LANDSCAPE)
  expect(sections[3]?.columnProperties).toEqual([
    { width: 408, paddingEnd: 48 },
    { width: 408, paddingEnd: 0 },
  ])
})

it('omits unknown header ids and writes a separator-only restriction', async () => {
  const data: IDocumentData = {
    id: 'separator-only',
    documentStyle: {},
    body: {
      dataStream: '\r\n',
      paragraphs: [{ startIndex: 0, paragraphId: 'paragraph-1' }],
      sectionBreaks: [{
        sectionId: 'section-1',
        startIndex: 1,
        defaultHeaderId: 'missing',
        columnSeparatorType: ColumnSeparatorType.BETWEEN_EACH_COLUMN,
      }],
    },
  }
  const xml = strFromU8(unzipSync(await writeDocument(data))['word/document.xml'] ?? new Uint8Array())
  expect(xml).not.toContain('headerReference')
  expect(xml).toContain('<w:cols w:num="1" w:sep="1"/>')
})

it('reads even-and-odd settings only when they are enabled', async () => {
  const plain = `${DOCUMENT_NAMESPACE}<w:body><w:p/></w:body></w:document>`
  const unrelated = await readDocument(packageOf({
    'word/document.xml': plain,
    'word/settings.xml': '<w:settings><w:compat/></w:settings>',
  }))
  expect(unrelated.documentStyle.evenAndOddHeaders).toBeUndefined()
  const disabled = await readDocument(packageOf({
    'word/document.xml': plain,
    'word/settings.xml': '<w:settings><w:evenAndOddHeaders w:val="off"/></w:settings>',
  }))
  expect(disabled.documentStyle.evenAndOddHeaders).toBeUndefined()
})

it('round-trips a table inside a header', async () => {
  const namespace = DOCUMENT_NAMESPACE.replace(
    '>',
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
  )
  const bytes = packageOf({
    'word/document.xml': `${namespace}<w:body><w:p/>`
      + '<w:sectPr><w:headerReference w:type="default" r:id="rId2"/></w:sectPr>'
      + '</w:body></w:document>',
    'word/_rels/document.xml.rels': '<Relationships>'
      + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"'
      + ' Target="header1.xml"/>'
      + '</Relationships>',
    'word/header1.xml': '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:p><w:r><w:t>Items</w:t></w:r></w:p>'
      + '<w:tbl><w:tblGrid><w:gridCol w:w="1200"/></w:tblGrid>'
      + '<w:tr><w:tc><w:p><w:r><w:t>Item</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
      + '</w:hdr>',
  })
  const data = await readDocument(bytes)
  expect(data.headers?.header1?.body.tables).toHaveLength(1)
  expect(data.headers?.header1?.tableSource?.['table-1']?.tableRows).toHaveLength(1)
  const reopened = await readDocument(await writeDocument(data))
  expect(reopened.headers?.header1?.body.tables).toHaveLength(1)
  expect(reopened.headers?.header1?.tableSource?.['table-1']?.tableRows).toHaveLength(1)
  expect(reopened.headers?.header1?.body.dataStream).toBe(data.headers?.header1?.body.dataStream)
})

it('maps section columns and break type', async () => {
  const xml = `${DOCUMENT_NAMESPACE}<w:body><w:p/>`
    + '<w:sectPr><w:type w:val="continuous"/>'
    + '<w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
    + '<w:cols w:num="2" w:space="720" w:sep="1"/>'
    + '</w:sectPr></w:body></w:document>'
  const data = await readDocument(documentOf(xml))
  const section = data.body?.sectionBreaks?.[0]
  expect(section?.sectionType).toBe(SectionType.CONTINUOUS)
  expect(section?.columnSeparatorType).toBe(ColumnSeparatorType.BETWEEN_EACH_COLUMN)
  const columns = section?.columnProperties ?? []
  expect(columns).toHaveLength(2)
  expect(columns[0]?.paddingEnd).toBe(48)
  expect(columns[1]?.paddingEnd).toBe(0)
  expect(Math.round((columns[0]?.width ?? 0) * 15)).toBe(Math.round((columns[1]?.width ?? 0) * 15))

  const reopenedSection = (await readDocument(await writeDocument(data))).body?.sectionBreaks?.[0]
  expect(reopenedSection?.sectionType).toBe(SectionType.CONTINUOUS)
  expect(reopenedSection?.columnSeparatorType).toBe(ColumnSeparatorType.BETWEEN_EACH_COLUMN)
  expect(reopenedSection?.columnProperties).toHaveLength(2)
  expect(reopenedSection?.columnProperties?.[0]?.paddingEnd).toBe(48)
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
  const struck = runs.find(run => body?.dataStream.slice(run.st, run.ed) === 'struck')
  expect(struck?.ts?.st).toEqual({ s: BooleanNumber.TRUE, t: TextDecoration.SINGLE })
  expect(runs.some(run => body?.dataStream.slice(run.st, run.ed) === 'A\tB')).toBe(false)
  const all = runs.find(run => body?.dataStream.slice(run.st, run.ed) === 'all')
  expect(all?.ts?.bl).toBe(BooleanNumber.TRUE)
  expect(all?.ts?.it).toBe(BooleanNumber.TRUE)
  expect(all?.ts?.ul).toEqual({ s: BooleanNumber.TRUE, t: TextDecoration.SINGLE })
  expect(body?.dataStream).toContain('link')
})

it('reads an empty body as one empty paragraph with a section break', async () => {
  const data = await readDocument(documentOf(`${DOCUMENT_NAMESPACE}<w:body/></w:document>`))
  expect(data.body?.dataStream).toBe('\r\n')
  expect(data.body?.paragraphs).toEqual([{ startIndex: 0, paragraphId: 'paragraph-1' }])
  expect(data.body?.sectionBreaks).toEqual([{ sectionId: 'section-1', startIndex: 1 }])
  expect(data.body?.tables).toEqual([])
})

it('reads a leading table, empty cells, and nested tables', async () => {
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
  expect(body?.dataStream).toContain('nested')
  expect(body?.paragraphs).toHaveLength(8)
  expect(body?.tables).toHaveLength(2)
  expect(body?.dataStream.startsWith('\r\x1A\x1B\x1C')).toBe(true)
  const table = data.tableSource?.['table-1']
  expect(table?.tableRows[0]?.tableCells[0]?.size?.width.v).toBe(100)
  expect(table?.tableRows[0]?.tableCells[1]?.size).toBeUndefined()
  expect(table?.tableRows[0]?.tableCells[2]?.size).toBeUndefined()
  expect(table?.tableRows[0]?.tableCells[3]?.size).toBeUndefined()
  expect(table?.tableColumns.map(column => column.size.width.v)).toEqual([80, 160, 160, 160])
  expect(data.tableSource?.['table-2']?.tableRows[0]?.tableCells).toHaveLength(1)
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
    ...(original.body === undefined ? {} : { body: original.body }),
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

it('writes a nested table from stream tokens without table metadata', async () => {
  const stream = '\x1A\x1B\x1Couter\r\n\x1A\x1B\x1Cinner\r\n\x1D\x0E\x0F\x1D\x0E\x0F'
  const data = snapshotOf(stream, {
    tables: [{ startIndex: 0, endIndex: stream.length, tableId: 'table-1' }],
  })
  const reopened = await readDocument(await writeDocument(data))
  expect(paragraphTexts(reopened)).toEqual(['', 'outer', 'inner', '', ''])
  expect(reopened.body?.dataStream).toContain('inner')
  expect(reopened.body?.tables).toHaveLength(2)
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
    commandTimeoutMs: 5_000,
  })).rejects.toMatchObject({ code: 'too-large' })
})

it('rejects a document above the configured table-cell bound', async () => {
  const bytes = new Uint8Array(await readFile(FIXTURE))
  await expect(readDocument(bytes, {
    maxFileBytes: 1_000_000,
    maxCells: 1,
    maxConversionMs: 1_000,
    commandTimeoutMs: 5_000,
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
    commandTimeoutMs: 5_000,
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
