/** In-place DOCX save: span splicing, property merge, and auxiliary appends. */
import { readFile } from 'node:fs/promises'
import type { IDocumentBody, IDocumentData, IParagraph } from '@univerjs/core'
import { CustomRangeType } from '@univerjs/core'
import { strFromU8, strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { OfficeExchangeError } from '../src/errors.ts'
import { readDocumentWithSource, writeDocumentFromSource } from '../src/provider/docx.ts'
import type { OfficeSourceState, UnitAnchor } from '../src/source/source-state.ts'
import type { OfficeWriteResult } from '../src/snapshots.ts'
import { comparePackages, formatDiffReport, unzipEntries, validateOpc } from './ooxml-package.ts'

const WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const WORD14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
const DOCUMENT = 'word/document.xml'
const STYLES = 'word/styles.xml'
const NUMBERING = 'word/numbering.xml'
const DOCUMENT_RELS = 'word/_rels/document.xml.rels'

const DEFAULT_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${WORD}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
</w:styles>`

const DEFAULT_NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${WORD}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`

const DEFAULT_SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${WORD}"><w:zoom w:percent="100"/></w:settings>`

interface DocxFixture {
  readonly body: string
  readonly header?: string
  readonly footer?: string
  readonly styles?: string
  readonly numbering?: string
  readonly relationships?: string
}

function docxPackage(fixture: DocxFixture): Uint8Array {
  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(contentTypes(fixture)),
    '_rels/.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument"`
      + ' Target="word/document.xml"/></Relationships>',
    ),
    [DOCUMENT]: strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:document xmlns:w="${WORD}" xmlns:w14="${WORD14}" xmlns:r="${REL}"><w:body>${fixture.body}</w:body></w:document>`,
    ),
    [DOCUMENT_RELS]: strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<Relationships xmlns="${PKG_REL}">`
      + `<Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>`
      + (fixture.numbering === undefined
        ? ''
        : `<Relationship Id="rId9" Type="${REL}/numbering" Target="numbering.xml"/>`)
      + (fixture.relationships ?? '')
      + '</Relationships>',
    ),
    [STYLES]: strToU8(fixture.styles ?? DEFAULT_STYLES),
    'word/settings.xml': strToU8(DEFAULT_SETTINGS),
  }
  if (fixture.header !== undefined) {
    entries['word/header1.xml'] = strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:hdr xmlns:w="${WORD}" xmlns:r="${REL}">${fixture.header}</w:hdr>`,
    )
  }
  if (fixture.footer !== undefined) {
    entries['word/footer1.xml'] = strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:ftr xmlns:w="${WORD}" xmlns:r="${REL}">${fixture.footer}</w:ftr>`,
    )
  }
  if (fixture.numbering !== undefined) entries[NUMBERING] = strToU8(fixture.numbering)
  return zipSync(entries, { level: 6 })
}

function contentTypes(fixture: DocxFixture): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<Types xmlns="${CT}">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '<Override PartName="/word/settings.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    + (fixture.numbering === undefined
      ? ''
      : '<Override PartName="/word/numbering.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>')
    + (fixture.header === undefined
      ? ''
      : '<Override PartName="/word/header1.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>')
    + (fixture.footer === undefined
      ? ''
      : '<Override PartName="/word/footer1.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>')
    + '</Types>'
}

const SECTION = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
  + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
  + ' w:header="720" w:footer="720"/></w:sectPr>'

function paragraph(text: string, properties = ''): string {
  return `<w:p>${properties}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
}

interface SidecarCarrier {
  officeExchange?: {
    readonly source?: unknown
  }
}

function sourceStateOf(data: IDocumentData): OfficeSourceState | undefined {
  return (data as SidecarCarrier).officeExchange?.source as OfficeSourceState | undefined
}

function anchorsOf(data: IDocumentData): readonly UnitAnchor[] {
  const state = sourceStateOf(data)
  if (state?.anchorMode !== 'per-unit') throw new Error('fixture has no per-unit source anchors')
  return state.anchors
}

function anchorOf(
  data: IDocumentData,
  kind: UnitAnchor['kind'],
  key: string,
  part?: string,
): UnitAnchor {
  const anchor = anchorsOf(data).find(candidate => (
    candidate.kind === kind && candidate.key === key && (part === undefined || candidate.part === part)
  ))
  if (anchor === undefined) throw new Error(`fixture has no ${kind} anchor ${key}`)
  return anchor
}

function paragraphById(data: IDocumentData, paragraphId: string): IParagraph {
  const paragraph = (data.body?.paragraphs ?? []).find(candidate => candidate.paragraphId === paragraphId)
  if (paragraph === undefined) throw new Error(`fixture has no paragraph ${paragraphId}`)
  return paragraph
}

function bodyOf(data: IDocumentData): IDocumentBody {
  const body = data.body
  if (body === undefined) throw new Error('fixture has no body')
  return body
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

function spliceStory(body: IDocumentBody, start: number, end: number, text: string): void {
  const stream = body.dataStream ?? ''
  const delta = text.length - (end - start)
  const deltaStart = start + text.length
  body.dataStream = stream.slice(0, start) + text + stream.slice(end)
  for (const paragraph of body.paragraphs ?? []) {
    if (paragraph.startIndex >= end) paragraph.startIndex += delta
  }
  for (const section of body.sectionBreaks ?? []) {
    if (section.startIndex >= end) section.startIndex += delta
  }
  for (const range of body.tables ?? []) {
    if (range.startIndex >= end) range.startIndex += delta
    if (range.endIndex > end) range.endIndex += delta
  }
  for (const block of body.docxRawBlocks ?? []) {
    if (block.startIndex >= end) block.startIndex += delta
  }
  body.textRuns = (body.textRuns ?? []).flatMap((run) => {
    if (run.ed <= start) return [run]
    if (run.st >= end) return [{ ...run, st: run.st + delta, ed: run.ed + delta }]
    const clippedStart = run.st < start ? run.st : start
    const clippedEnd = run.ed > end ? run.ed + delta : deltaStart
    return clippedEnd > clippedStart ? [{ ...run, st: clippedStart, ed: clippedEnd }] : []
  })
}

function replaceStoryText(body: IDocumentBody, from: string, to: string): void {
  const stream = body.dataStream ?? ''
  const start = stream.indexOf(from)
  if (start < 0) throw new Error(`fixture text ${from} is absent`)
  spliceStory(body, start, start + from.length, to)
}

function replaceText(data: IDocumentData, from: string, to: string): void {
  replaceStoryText(bodyOf(data), from, to)
}

function insertParagraph(
  data: IDocumentData,
  after: IParagraph,
  text: string,
  extra: Partial<IParagraph> = {},
): IParagraph {
  const body = bodyOf(data)
  const at = after.startIndex + 1
  spliceStory(body, at, at, `${text}\r`)
  const inserted: IParagraph = { paragraphId: 'para_inserted', startIndex: at + text.length, ...extra }
  body.paragraphs = [...(body.paragraphs ?? []), inserted].sort((left, right) => left.startIndex - right.startIndex)
  return inserted
}

function removeParagraph(data: IDocumentData, paragraph: IParagraph): void {
  const body = bodyOf(data)
  const start = contentStartOf(body.dataStream ?? '', paragraph.startIndex)
  spliceStory(body, start, paragraph.startIndex + 1, '')
  body.paragraphs = (body.paragraphs ?? []).filter(candidate => candidate !== paragraph)
}

function entryOf(entries: ReadonlyMap<string, Uint8Array>, name: string): Uint8Array {
  const bytes = entries.get(name)
  if (bytes === undefined) throw new Error(`package has no ${name}`)
  return bytes
}

function expectOnlySpan(
  sourcePart: Uint8Array,
  writtenPart: Uint8Array,
  span: { readonly start: number; readonly end: number },
): void {
  expect(writtenPart.subarray(0, span.start)).toEqual(sourcePart.subarray(0, span.start))
  const tail = sourcePart.subarray(span.end)
  expect(writtenPart.subarray(writtenPart.byteLength - tail.byteLength)).toEqual(tail)
}

function expectDispositionOf(result: OfficeWriteResult, changed: readonly string[]): void {
  const sorted = [...changed].sort()
  expect(result.disposition.spliced).toEqual({ count: sorted.length, names: sorted })
  expect(result.disposition.added).toEqual({ count: 0, names: [] })
  expect(result.disposition.removed).toEqual({ count: 0, names: [] })
  expect(result.disposition.nameLimit).toBe(256)
}

function expectDiff(
  source: Uint8Array,
  written: Uint8Array,
  changed: readonly string[],
): ReturnType<typeof comparePackages> {
  const diff = comparePackages(source, written, { changed })
  expect(diff.ok, formatDiffReport(diff)).toBe(true)
  return diff
}

function appendOffset(part: string, closingTag: string): { start: number; end: number } {
  const character = part.lastIndexOf(closingTag)
  expect(character).toBeGreaterThanOrEqual(0)
  const start = strToU8(part.slice(0, character)).byteLength
  return { start, end: start }
}

function spanOf(anchor: UnitAnchor): { start: number; end: number } {
  return { start: anchor.start, end: anchor.end }
}

async function refusalOf(result: Promise<unknown>): Promise<OfficeExchangeError> {
  try {
    await result
  } catch (error) {
    if (error instanceof OfficeExchangeError) return error
    throw error
  }
  throw new Error('Expected the writer to refuse')
}

describe('writeDocumentFromSource', () => {
  it('splices an edited paragraph and copies every other entry byte-for-byte', async () => {
    const source = docxPackage({
      body: `${paragraph('Alpha paragraph')}${paragraph('Beta paragraph')}${SECTION}`,
    })
    const data = await readDocumentWithSource(source, 'docx')
    const anchor = anchorOf(data, 'paragraph', 'para_2')
    replaceText(data, 'Beta paragraph', 'Gamma paragraph')
    const result = await writeDocumentFromSource(data, source)
    const diff = expectDiff(source, result.bytes, [DOCUMENT])
    expect(result.disposition.verbatimCount).toBe(diff.entries.filter(entry => entry.status === 'identical').length)
    expectDispositionOf(result, [DOCUMENT])
    const written = unzipEntries(result.bytes)
    expectOnlySpan(entryOf(unzipEntries(source), DOCUMENT), entryOf(written, DOCUMENT), spanOf(anchor))
    const documentXml = strFromU8(entryOf(written, DOCUMENT))
    expect(documentXml).toContain('Gamma paragraph')
    expect(documentXml).not.toContain('Beta paragraph')
    expect(documentXml).toContain('Alpha paragraph')
    expect(validateOpc(written)).toEqual([])
    const repeated = await writeDocumentFromSource(data, source)
    expect(repeated.bytes).toEqual(result.bytes)
  })

  it('writes an unchanged snapshot back as the identical source package', async () => {
    const source = docxPackage({
      body: `${paragraph('Alpha paragraph')}${paragraph('Beta paragraph')}${SECTION}`,
    })
    const data = await readDocumentWithSource(source, 'docx')
    const result = await writeDocumentFromSource(data, source)
    expect(result.bytes).toEqual(source)
    expectDispositionOf(result, [])
    expect(result.disposition.verbatimCount).toBe(unzipEntries(source).size)
  })

  it('splices a paragraph inside a table cell and re-emits its unmapped cell property', async () => {
    const source = docxPackage({
      body: `${paragraph('Alpha paragraph')}<w:tbl>`
        + '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        + '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>'
        + '<w:tr>'
        + '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:tcFitText/></w:tcPr>'
        + `${paragraph('Region')}</w:tc>`
        + '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
        + `${paragraph('Revenue')}</w:tc>`
        + '</w:tr></w:tbl>'
        + SECTION,
    })
    const data = await readDocumentWithSource(source, 'docx')
    const anchor = anchorOf(data, 'tableCell', 'table-1/0/0')
    replaceText(data, 'Region', 'Area')
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT])
    expectDispositionOf(result, [DOCUMENT])
    const written = unzipEntries(result.bytes)
    expectOnlySpan(entryOf(unzipEntries(source), DOCUMENT), entryOf(written, DOCUMENT), spanOf(anchor))
    const documentXml = strFromU8(entryOf(written, DOCUMENT))
    expect(documentXml).toContain('>Area</w:t>')
    expect(documentXml).toContain('<w:tcFitText/>')
    expect(documentXml).toContain('Revenue')
    expect(validateOpc(written)).toEqual([])
  })

  it('splices header and footer paragraphs in their own parts', async () => {
    const source = docxPackage({
      header: paragraph('Confidential header'),
      footer: paragraph('Page footer'),
      body: `${paragraph('Body paragraph')}<w:sectPr>`
        + '<w:headerReference w:type="default" r:id="rId2"/>'
        + '<w:footerReference w:type="default" r:id="rId3"/>'
        + '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
      relationships: `<Relationship Id="rId2" Type="${REL}/header" Target="header1.xml"/>`
        + `<Relationship Id="rId3" Type="${REL}/footer" Target="footer1.xml"/>`,
    })
    const data = await readDocumentWithSource(source, 'docx')
    const headerAnchor = anchorOf(data, 'paragraph', 'para_1', 'word/header1.xml')
    const footerAnchor = anchorOf(data, 'paragraph', 'para_1', 'word/footer1.xml')
    const header = data.headers?.['header1']
    const footer = data.footers?.['footer1']
    if (header === undefined || footer === undefined) throw new Error('fixture has no header or footer')
    replaceStoryText(header.body, 'Confidential header', 'Restricted header')
    replaceStoryText(footer.body, 'Page footer', 'Page number')
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, ['word/header1.xml', 'word/footer1.xml'])
    expectDispositionOf(result, ['word/footer1.xml', 'word/header1.xml'])
    const written = unzipEntries(result.bytes)
    const sourceEntries = unzipEntries(source)
    expectOnlySpan(entryOf(sourceEntries, 'word/header1.xml'), entryOf(written, 'word/header1.xml'), spanOf(headerAnchor))
    expectOnlySpan(entryOf(sourceEntries, 'word/footer1.xml'), entryOf(written, 'word/footer1.xml'), spanOf(footerAnchor))
    expect(strFromU8(entryOf(written, 'word/header1.xml'))).toContain('Restricted header')
    expect(strFromU8(entryOf(written, 'word/footer1.xml'))).toContain('Page number')
    expect(validateOpc(written)).toEqual([])
  })

  it('re-emits a raw drawing block exactly once inside an edited paragraph', async () => {
    const source = new Uint8Array(await readFile(new URL('./fixtures/preserve-image.docx', import.meta.url)))
    const data = await readDocumentWithSource(source, 'docx')
    const anchor = anchorOf(data, 'paragraph', 'para_1')
    replaceText(data, 'after image', 'later image')
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT])
    const written = unzipEntries(result.bytes)
    expectOnlySpan(entryOf(unzipEntries(source), DOCUMENT), entryOf(written, DOCUMENT), spanOf(anchor))
    const documentXml = strFromU8(entryOf(written, DOCUMENT))
    expect(documentXml.match(/<w:drawing>/gu) ?? []).toHaveLength(1)
    expect(documentXml).toContain('later image')
    expect(entryOf(written, 'word/media/image1.jpeg')).toEqual(entryOf(unzipEntries(source), 'word/media/image1.jpeg'))
    expect(validateOpc(written)).toEqual([])
  })

  it('inserts a new paragraph at a neighbouring anchor boundary', async () => {
    const source = docxPackage({
      body: `${paragraph('Alpha paragraph')}${paragraph('Beta paragraph')}${SECTION}`,
    })
    const data = await readDocumentWithSource(source, 'docx')
    const first = paragraphById(data, 'para_1')
    const at = anchorOf(data, 'paragraph', 'para_1').end
    insertParagraph(data, first, 'Inserted paragraph')
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT])
    const written = unzipEntries(result.bytes)
    expectOnlySpan(entryOf(unzipEntries(source), DOCUMENT), entryOf(written, DOCUMENT), { start: at, end: at })
    const documentXml = strFromU8(entryOf(written, DOCUMENT))
    expect(documentXml.indexOf('Alpha paragraph')).toBeLessThan(documentXml.indexOf('Inserted paragraph'))
    expect(documentXml.indexOf('Inserted paragraph')).toBeLessThan(documentXml.indexOf('Beta paragraph'))
    expect(validateOpc(written)).toEqual([])
  })

  it('deletes one paragraph span and leaves its neighbours intact', async () => {
    const source = docxPackage({
      body: `${paragraph('Alpha paragraph')}${paragraph('Beta paragraph')}${paragraph('Gamma paragraph')}${SECTION}`,
    })
    const data = await readDocumentWithSource(source, 'docx')
    const anchor = anchorOf(data, 'paragraph', 'para_2')
    removeParagraph(data, paragraphById(data, 'para_2'))
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT])
    const written = unzipEntries(result.bytes)
    expectOnlySpan(entryOf(unzipEntries(source), DOCUMENT), entryOf(written, DOCUMENT), spanOf(anchor))
    const documentXml = strFromU8(entryOf(written, DOCUMENT))
    expect(documentXml).not.toContain('Beta paragraph')
    expect(documentXml).toContain('Alpha paragraph')
    expect(documentXml).toContain('Gamma paragraph')
    expect(validateOpc(written)).toEqual([])
  })

  it('does not insert the reader synthetic spacer when the body ends with a table', async () => {
    const source = docxPackage({
      body: paragraph('Alpha paragraph')
        + '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        + '<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>'
        + `<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${paragraph('Cell')}</w:tc></w:tr>`
        + '</w:tbl>'
        + SECTION,
    })
    const data = await readDocumentWithSource(source, 'docx')
    expect(data.body?.paragraphs?.length).toBe(3)
    const result = await writeDocumentFromSource(data, source)
    expect(result.bytes).toEqual(source)
  })

  it('moves an insertion that follows a table cell to the table boundary', async () => {
    const source = docxPackage({
      body: paragraph('Alpha paragraph')
        + '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        + '<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>'
        + `<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${paragraph('Cell')}</w:tc></w:tr>`
        + '</w:tbl>'
        + paragraph('Omega paragraph')
        + SECTION,
    })
    const data = await readDocumentWithSource(source, 'docx')
    const body = bodyOf(data)
    const omega = paragraphById(data, 'para_3')
    const at = contentStartOf(body.dataStream ?? '', omega.startIndex)
    spliceStory(body, at, at, 'Inserted paragraph\r')
    body.paragraphs = [
      ...(body.paragraphs ?? []),
      { paragraphId: 'para_inserted', startIndex: at + 'Inserted paragraph'.length },
    ].sort((left, right) => left.startIndex - right.startIndex)
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT])
    const written = unzipEntries(result.bytes)
    const documentXml = strFromU8(entryOf(written, DOCUMENT))
    expect(documentXml).toContain('</w:tbl><w:p>')
    expect(documentXml.indexOf('Inserted paragraph')).toBeGreaterThan(documentXml.indexOf('</w:tbl>'))
    expect(documentXml.indexOf('Inserted paragraph')).toBeLessThan(documentXml.indexOf('Omega paragraph'))
    expect(validateOpc(written)).toEqual([])
  })

  it('refuses unit reordering without producing bytes', async () => {
    const source = docxPackage({
      body: `${paragraph('Alpha paragraph')}${paragraph('Beta paragraph')}${SECTION}`,
    })
    const data = await readDocumentWithSource(source, 'docx')
    const first = paragraphById(data, 'para_1')
    const second = paragraphById(data, 'para_2')
    const swap = first.paragraphId
    first.paragraphId = second.paragraphId
    second.paragraphId = swap
    const error = await refusalOf(writeDocumentFromSource(data, source))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('reordering')
  })

  it('re-emits unmapped paragraph and run properties the serializer does not model', async () => {
    const source = docxPackage({
      body: '<w:p><w:pPr><w:suppressLineNumbers/></w:pPr>'
        + '<w:r><w:rPr><w:b/><w:noProof/><w:lang w:val="en-US"/></w:rPr>'
        + '<w:t>Styled paragraph</w:t></w:r></w:p>'
        + SECTION,
    })
    const data = await readDocumentWithSource(source, 'docx')
    const anchor = anchorOf(data, 'paragraph', 'para_1')
    replaceText(data, 'Styled paragraph', 'Styled section')
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT])
    const written = unzipEntries(result.bytes)
    expectOnlySpan(entryOf(unzipEntries(source), DOCUMENT), entryOf(written, DOCUMENT), spanOf(anchor))
    const documentXml = strFromU8(entryOf(written, DOCUMENT))
    expect(documentXml).toContain('<w:suppressLineNumbers/>')
    expect(documentXml).toContain('<w:noProof/>')
    expect(documentXml).toContain('<w:lang w:val="en-US"/>')
    expect(documentXml).toContain('<w:b/>')
    expect(documentXml).toContain('Styled section')
    expect(validateOpc(written)).toEqual([])
  })

  it('carries Word root metadata attributes through an edited paragraph', async () => {
    const body = '<w:p w:rsidR="00AB12CD" w:rsidRDefault="00AB12CD" w:rsidP="00AB12CD"'
      + ' w14:paraId="1A2B3C4D" w14:textId="77777777">'
      + '<w:r><w:t>Rsid paragraph</w:t></w:r></w:p>'
      + SECTION
    const source = docxPackage({ body })
    const data = await readDocumentWithSource(source, 'docx')
    replaceText(data, 'Rsid paragraph', 'Rsid section')
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT])
    const documentXml = strFromU8(entryOf(unzipEntries(result.bytes), DOCUMENT))
    for (const attribute of [
      'w:rsidR="00AB12CD"',
      'w:rsidRDefault="00AB12CD"',
      'w:rsidP="00AB12CD"',
      'w14:paraId="1A2B3C4D"',
      'w14:textId="77777777"',
    ]) {
      expect(documentXml.split(attribute)).toHaveLength(2)
    }
    expect(documentXml).toContain('Rsid section')
    expect(validateOpc(unzipEntries(result.bytes))).toEqual([])
  })

  it('merges an unmapped run property after non-ASCII text', async () => {
    const body = '<w:p><w:r><w:t>日本語</w:t></w:r>'
      + '<w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space="preserve">After</w:t></w:r></w:p>'
      + SECTION
    const source = docxPackage({ body })
    const data = await readDocumentWithSource(source, 'docx')
    replaceText(data, 'After', 'Later')
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT])
    const documentXml = strFromU8(entryOf(unzipEntries(result.bytes), DOCUMENT))
    expect(documentXml).toContain('<w:noProof/>')
    expect(documentXml).toContain('日本語')
    expect(documentXml).toContain('Later')
    expect(validateOpc(unzipEntries(result.bytes))).toEqual([])
  })

  it('preserves a run-level object exactly once through an edit', async () => {
    const body = '<w:p><w:r><w:t>Before</w:t></w:r>'
      + '<w:r><w:object/></w:r>'
      + '<w:r><w:t xml:space="preserve">After</w:t></w:r></w:p>'
      + SECTION
    const source = docxPackage({ body })
    const data = await readDocumentWithSource(source, 'docx')
    replaceText(data, 'After', 'Later')
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT])
    const documentXml = strFromU8(entryOf(unzipEntries(result.bytes), DOCUMENT))
    expect(documentXml.match(/<w:object\/>/gu) ?? []).toHaveLength(1)
    expect(documentXml).toContain('Before')
    expect(documentXml).toContain('Later')
    expect(validateOpc(unzipEntries(result.bytes))).toEqual([])
  })

  it('keeps model keys aligned across repeated stale saves', async () => {
    const source = docxPackage({
      body: `${paragraph('Alpha paragraph')}${paragraph('Beta paragraph')}${paragraph('Gamma paragraph')}${SECTION}`,
    })
    const data = await readDocumentWithSource(source, 'docx')
    insertParagraph(data, paragraphById(data, 'para_1'), 'Inserted paragraph')
    replaceText(data, 'Beta paragraph', 'Beta revised')
    const first = await writeDocumentFromSource(data, source)
    const second = await writeDocumentFromSource(data, first.bytes)
    expect(second.bytes).toEqual(first.bytes)
    expectDispositionOf(second, [])

    replaceText(data, 'Gamma paragraph', 'Gamma revised')
    const third = await writeDocumentFromSource(data, second.bytes)
    expectDiff(second.bytes, third.bytes, [DOCUMENT])
    const secondXml = strFromU8(entryOf(unzipEntries(second.bytes), DOCUMENT))
    const thirdXml = strFromU8(entryOf(unzipEntries(third.bytes), DOCUMENT))
    const secondParagraph = secondXml.indexOf('<w:p><w:r><w:t xml:space="preserve">Gamma paragraph')
    const thirdParagraph = thirdXml.indexOf('<w:p><w:r><w:t xml:space="preserve">Gamma revised')
    expect(thirdParagraph).toBeGreaterThanOrEqual(0)
    expect(thirdXml.slice(0, thirdParagraph)).toBe(secondXml.slice(0, secondParagraph))
    const secondEnd = secondXml.indexOf('</w:p>', secondParagraph) + '</w:p>'.length
    const thirdEnd = thirdXml.indexOf('</w:p>', thirdParagraph) + '</w:p>'.length
    expect(thirdXml.slice(thirdEnd)).toBe(secondXml.slice(secondEnd))
    expect(thirdXml).toContain('Alpha paragraph')
    expect(thirdXml).toContain('Inserted paragraph')
    expect(thirdXml).toContain('Beta revised')
    expect(thirdXml).toContain('Gamma revised')
    expect(validateOpc(unzipEntries(third.bytes))).toEqual([])
  })

  it('re-anchors a stale sidecar whose spans exceed the current part length', async () => {
    const long = docxPackage({
      body: `${' '.repeat(4096)}${paragraph('Alpha paragraph')}${paragraph('Beta paragraph')}${SECTION}`,
    })
    const short = docxPackage({
      body: `${paragraph('Alpha paragraph')}${paragraph('Beta paragraph')}${SECTION}`,
    })
    const data = await readDocumentWithSource(long, 'docx')
    const result = await writeDocumentFromSource(data, short)
    const documentXml = strFromU8(entryOf(unzipEntries(result.bytes), DOCUMENT))
    expect(documentXml).toContain('Alpha paragraph')
    expect(documentXml).toContain('Beta paragraph')
    expect(validateOpc(unzipEntries(result.bytes))).toEqual([])
  })

  it('appends a missing style before the styles closing tag', async () => {
    const source = docxPackage({
      body: `${paragraph('Alpha paragraph')}${SECTION}`,
    })
    const data = await readDocumentWithSource(source, 'docx')
    insertParagraph(data, paragraphById(data, 'para_1'), 'Heading paragraph', { styleId: 'Heading3' })
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT, STYLES])
    expectDispositionOf(result, [DOCUMENT, STYLES])
    const written = unzipEntries(result.bytes)
    const stylesXml = strFromU8(entryOf(written, STYLES))
    expect(stylesXml).toContain('w:styleId="Heading3"')
    expect(stylesXml).toContain('<w:outlineLvl w:val="2"/>')
    expectOnlySpan(
      entryOf(unzipEntries(source), STYLES),
      entryOf(written, STYLES),
      appendOffset(strFromU8(entryOf(unzipEntries(source), STYLES)), '</w:styles>'),
    )
    const documentXml = strFromU8(entryOf(written, DOCUMENT))
    expect(documentXml).toContain('<w:pStyle w:val="Heading3"/>')
    expect(validateOpc(written)).toEqual([])
  })

  it('inserts a new abstract definition before existing instances and appends its instance', async () => {
    const source = docxPackage({
      body: `${paragraph('Alpha paragraph')}${SECTION}`,
      numbering: DEFAULT_NUMBERING,
    })
    const data = await readDocumentWithSource(source, 'docx')
    insertParagraph(data, paragraphById(data, 'para_1'), 'List paragraph', {
      bullet: { listType: 'numbering-9', listId: 'num-9', nestingLevel: 0 },
    })
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT, NUMBERING])
    const written = unzipEntries(result.bytes)
    const numberingXml = strFromU8(entryOf(written, NUMBERING))
    expect(numberingXml).toContain('<w:abstractNum w:abstractNumId="1">')
    expect(numberingXml).toContain('<w:num w:numId="2">')
    expect(numberingXml.indexOf('<w:abstractNum w:abstractNumId="1">'))
      .toBeLessThan(numberingXml.indexOf('<w:num w:numId="1">'))
    expect(numberingXml.indexOf('<w:num w:numId="2">'))
      .toBeGreaterThan(numberingXml.indexOf('<w:num w:numId="1">'))
    const sourceNumbering = strFromU8(entryOf(unzipEntries(source), NUMBERING))
    const abstractOffset = strToU8(sourceNumbering.slice(0, sourceNumbering.indexOf('<w:num '))).byteLength
    const sourceBytes = entryOf(unzipEntries(source), NUMBERING)
    const writtenBytes = entryOf(written, NUMBERING)
    expect(writtenBytes.subarray(0, abstractOffset)).toEqual(sourceBytes.subarray(0, abstractOffset))
    const closing = appendOffset(sourceNumbering, '</w:numbering>')
    const tail = sourceBytes.subarray(closing.start)
    expect(writtenBytes.subarray(writtenBytes.byteLength - tail.byteLength)).toEqual(tail)
    const documentXml = strFromU8(entryOf(written, DOCUMENT))
    expect(documentXml).toContain('<w:numId w:val="2"/>')
    expect(validateOpc(written)).toEqual([])
  })

  it('appends a list definition at the closing tag when numbering has no instance', async () => {
    const numbering = `<w:numbering xmlns:w="${WORD}">`
      + '<w:abstractNum w:abstractNumId="0">'
      + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>'
      + '</w:abstractNum>'
      + '</w:numbering>'
    const source = docxPackage({
      body: `${paragraph('Alpha paragraph')}${SECTION}`,
      numbering,
    })
    const data = await readDocumentWithSource(source, 'docx')
    insertParagraph(data, paragraphById(data, 'para_1'), 'List paragraph', {
      bullet: { listType: 'numbering-9', listId: 'num-9', nestingLevel: 0 },
    })
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT, NUMBERING])
    const numberingXml = strFromU8(entryOf(unzipEntries(result.bytes), NUMBERING))
    expect(numberingXml).toContain('<w:abstractNum w:abstractNumId="1">')
    expect(numberingXml).toContain('<w:num w:numId="1">')
    expect(numberingXml.indexOf('<w:abstractNum w:abstractNumId="1">'))
      .toBeLessThan(numberingXml.indexOf('<w:num w:numId="1">'))
    expect(validateOpc(unzipEntries(result.bytes))).toEqual([])
  })

  it('appends a hyperlink relationship when an edit introduces a new URL', async () => {
    const source = docxPackage({
      body: `${paragraph('Alpha paragraph')}${SECTION}`,
    })
    const data = await readDocumentWithSource(source, 'docx')
    const editedParagraph = paragraphById(data, 'para_1')
    const start = contentStartOf(bodyOf(data).dataStream ?? '', editedParagraph.startIndex)
    replaceText(data, 'Alpha paragraph', 'Linked paragraph')
    const ranges = bodyOf(data).customRanges
    if (ranges === undefined) throw new Error('fixture has no custom ranges')
    ranges.push({
      startIndex: start,
      endIndex: editedParagraph.startIndex - 1,
      rangeId: 'link-1',
      rangeType: CustomRangeType.HYPERLINK,
      properties: { url: 'https://example.test/new' },
    })
    const result = await writeDocumentFromSource(data, source)
    expectDiff(source, result.bytes, [DOCUMENT, DOCUMENT_RELS])
    const written = unzipEntries(result.bytes)
    const relsXml = strFromU8(entryOf(written, DOCUMENT_RELS))
    expect(relsXml).toContain('Target="https://example.test/new" TargetMode="External"')
    expectOnlySpan(
      entryOf(unzipEntries(source), DOCUMENT_RELS),
      entryOf(written, DOCUMENT_RELS),
      appendOffset(strFromU8(entryOf(unzipEntries(source), DOCUMENT_RELS)), '</Relationships>'),
    )
    const documentXml = strFromU8(entryOf(written, DOCUMENT))
    expect(documentXml).toContain('<w:hyperlink r:id="rId2">')
    expect(validateOpc(written)).toEqual([])
  })

  it('refuses a tracked-change wrapper with a typed unsupported-feature error', async () => {
    const source = docxPackage({
      body: '<w:p><w:r><w:t>Edited text</w:t></w:r>'
        + '<w:ins w:id="1" w:author="reviewer" w:date="2026-01-01T00:00:00Z">'
        + '<w:r><w:t>Tracked text</w:t></w:r></w:ins></w:p>'
        + SECTION,
    })
    const data = await readDocumentWithSource(source, 'docx')
    replaceText(data, 'Edited text', 'Changed text')
    const error = await refusalOf(writeDocumentFromSource(data, source))
    expect(error.code).toBe('unsupported-feature')
    expect(error.message).toContain('w:ins')
  })

  it('refuses a snapshot without source anchors', async () => {
    const source = docxPackage({
      body: `${paragraph('Alpha paragraph')}${SECTION}`,
    })
    const data = await readDocumentWithSource(source, 'docx')
    delete (data as SidecarCarrier).officeExchange
    const error = await refusalOf(writeDocumentFromSource(data, source))
    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('no source anchors')
  })
})
