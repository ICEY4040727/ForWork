/**
 * DOCX lists, numbering, and hyperlink mapping: OOXML-to-Univer metadata and
 * the write/read round-trip that keeps paragraphs editable.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { expect, it } from 'vitest'
import {
  BooleanNumber,
  CustomRangeType,
  ListGlyphType,
  validateDocumentStructure,
  type IDocumentData,
} from '@univerjs/core'
import { readDocument, writeDocument } from '../src/provider/docx.ts'

const DOCUMENT_XML = 'word/document.xml'
const NUMBERING_XML = 'word/numbering.xml'
const DOCUMENT_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const HYPERLINK_RELATIONSHIP = 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'

function packageOf(entries: Readonly<Record<string, string>>): Uint8Array {
  const files: Record<string, Uint8Array> = {}
  for (const [name, content] of Object.entries(entries)) files[name] = strToU8(content)
  return zipSync(files)
}

function documentPackage(body: string, extra: Readonly<Record<string, string>> = {}): Uint8Array {
  return packageOf({
    '[Content_Types].xml': '<Types/>',
    [DOCUMENT_XML]: `<w:document ${DOCUMENT_NAMESPACE}><w:body>${body}</w:body></w:document>`,
    ...extra,
  })
}

function writtenXml(bytes: Uint8Array, part: string): string {
  const content = unzipSync(bytes)[part]
  return content === undefined ? '' : strFromU8(content)
}

function numberedParagraph(text: string, level: number, numId = 1): string {
  return `<w:p><w:pPr><w:numPr><w:ilvl w:val="${String(level)}"/>`
    + `<w:numId w:val="${String(numId)}"/></w:numPr></w:pPr>`
    + `<w:r><w:t>${text}</w:t></w:r></w:p>`
}

function twoLevelNumbering(): string {
  return `<w:numbering ${DOCUMENT_NAMESPACE}>`
    + '<w:abstractNum w:abstractNumId="0">'
    + '<w:multiLevelType w:val="hybridMultilevel"/>'
    + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>'
    + '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
    + '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/>'
    + '<w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>'
    + '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/></w:lvl>'
    + '</w:abstractNum>'
    + '<w:num w:numId="1"><w:abstractNumId w:val="0"/>'
    + '<w:lvlOverride w:ilvl="1"><w:startOverride w:val="3"/></w:lvlOverride></w:num>'
    + '</w:numbering>'
}

function hyperlink(attributes: string, text: string, bold = false): string {
  const properties = bold ? '<w:rPr><w:b/></w:rPr>' : ''
  return `<w:hyperlink ${attributes}><w:r>${properties}<w:t>${text}</w:t></w:r></w:hyperlink>`
}

it('maps a numbered list with two levels and its per-instance start', async () => {
  const bytes = documentPackage(
    numberedParagraph('First', 0) + numberedParagraph('Nested', 1),
    { [NUMBERING_XML]: twoLevelNumbering() },
  )
  const data = await readDocument(bytes)
  const levels = data.lists?.['numbering-1']?.nestingLevel ?? []
  expect(data.lists?.['numbering-1']?.listType).toBe('numbering-1')
  expect(levels).toHaveLength(3)
  expect(levels[0]?.glyphType).toBe(ListGlyphType.DECIMAL)
  expect(levels[0]?.glyphFormat).toBe('%1.')
  expect(levels[0]?.startNumber).toBe(1)
  expect(levels[0]?.paragraphProperties?.indentStart?.v).toBe(48)
  expect(levels[1]?.glyphType).toBe(ListGlyphType.LOWER_LETTER)
  expect(levels[1]?.glyphFormat).toBe('%2)')
  expect(levels[1]?.startNumber).toBe(3)
  const paragraphs = data.body?.paragraphs ?? []
  expect(paragraphs[0]?.bullet).toMatchObject({ listType: 'numbering-1', listId: 'num-1', nestingLevel: 0 })
  expect(paragraphs[1]?.bullet).toMatchObject({
    listType: 'numbering-1',
    listId: 'num-1',
    nestingLevel: 1,
    startNumber: 3,
  })
  expect(validateDocumentStructure(data)).toEqual([])

  const encoded = await writeDocument(data)
  expect(writtenXml(encoded, DOCUMENT_XML)).toContain('<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>')
  const numbering = writtenXml(encoded, NUMBERING_XML)
  expect(numbering).toContain('<w:numFmt w:val="decimal"/>')
  expect(numbering).toContain('<w:numFmt w:val="lowerLetter"/>')

  const reopened = await readDocument(encoded)
  expect(reopened.lists?.['numbering-1']?.nestingLevel[1]?.glyphType).toBe(ListGlyphType.LOWER_LETTER)
  expect(reopened.body?.paragraphs?.[1]?.bullet).toEqual(paragraphs[1]?.bullet)
  expect(validateDocumentStructure(reopened)).toEqual([])
})

it('maps a legal numbering level with its paragraph and run properties', async () => {
  const numbering = `<w:numbering ${DOCUMENT_NAMESPACE}>`
    + '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
    + '<w:start w:val="1"/><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1、"/>'
    + '<w:isLgl/><w:pPr><w:ind w:left="360"/></w:pPr><w:rPr><w:b/></w:rPr>'
    + '</w:lvl></w:abstractNum>'
    + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    + '</w:numbering>'
  const bytes = documentPackage(numberedParagraph('Item', 0), { [NUMBERING_XML]: numbering })
  const data = await readDocument(bytes)
  const level = data.lists?.['numbering-1']?.nestingLevel[0]
  expect(level?.isLegal).toBe(true)
  expect(level?.glyphType).toBe(ListGlyphType.CHINESE_COUNTING)
  expect(level?.paragraphProperties?.indentStart?.v).toBe(24)
  expect(level?.textStyle?.bl).toBe(BooleanNumber.TRUE)

  const encoded = await writeDocument(data)
  const written = writtenXml(encoded, NUMBERING_XML)
  expect(written).toContain('<w:isLgl/>')
  expect(written).toContain('<w:rPr><w:b/></w:rPr>')
  expect((await readDocument(encoded)).lists?.['numbering-1']?.nestingLevel[0]?.textStyle?.bl)
    .toBe(BooleanNumber.TRUE)
})

it('maps a bullet list glyph and keeps it through a round trip', async () => {
  const numbering = `<w:numbering ${DOCUMENT_NAMESPACE}>`
    + '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
    + '<w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\u25CF"/>'
    + '</w:lvl></w:abstractNum>'
    + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    + '</w:numbering>'
  const bytes = documentPackage(numberedParagraph('Point', 0), { [NUMBERING_XML]: numbering })
  const data = await readDocument(bytes)
  const level = data.lists?.['numbering-1']?.nestingLevel[0]
  expect(level?.glyphType).toBe(ListGlyphType.BULLET)
  expect(level?.glyphSymbol).toBe('\u25CF')
  expect(data.body?.paragraphs?.[0]?.bullet?.listType).toBe('numbering-1')

  const reopened = await readDocument(await writeDocument(data))
  expect(reopened.lists?.['numbering-1']?.nestingLevel[0]?.glyphType).toBe(ListGlyphType.BULLET)
  expect(reopened.lists?.['numbering-1']?.nestingLevel[0]?.glyphSymbol).toBe('\u25CF')
  expect(validateDocumentStructure(reopened)).toEqual([])
})

it('maps an external hyperlink and a bookmark anchor into custom ranges', async () => {
  const target = 'https://example.test/a?b=1'
  const bytes = documentPackage(
    '<w:p><w:r><w:t>Visit </w:t></w:r>' + hyperlink('r:id="rId9"', 'site', true)
    + '<w:r><w:t> now</w:t></w:r></w:p>'
    + '<w:p>' + hyperlink('w:anchor="section-1"', 'jump')
    + hyperlink('w:anchor="section-2"', 'next') + '</w:p>',
    {
      'word/_rels/document.xml.rels': '<Relationships>'
        + `<Relationship Id="rId9" ${HYPERLINK_RELATIONSHIP} Target="${target}" TargetMode="External"/>`
        + '</Relationships>',
    },
  )
  const data = await readDocument(bytes)
  const stream = data.body?.dataStream ?? ''
  expect(stream).toContain('\x1Fsite\x1E')
  expect(stream).toContain('\x1Fjump\x1E')
  const ranges = data.body?.customRanges ?? []
  expect(ranges).toHaveLength(3)
  expect(ranges[0]).toMatchObject({
    rangeType: CustomRangeType.HYPERLINK,
    properties: { url: target },
  })
  expect(stream[ranges[0]?.startIndex ?? 0]).toBe('\x1F')
  expect(stream[ranges[0]?.endIndex ?? 0]).toBe('\x1E')
  expect(ranges[1]).toMatchObject({
    rangeType: CustomRangeType.HYPERLINK,
    properties: { bookmarkId: 'section-1' },
  })
  expect(ranges[2]).toMatchObject({
    rangeType: CustomRangeType.HYPERLINK,
    properties: { bookmarkId: 'section-2' },
  })
  const bold = data.body?.textRuns?.find(run => stream.slice(run.st, run.ed) === 'site')
  expect(bold?.ts?.bl).toBe(BooleanNumber.TRUE)
  expect(validateDocumentStructure(data)).toEqual([])

  const encoded = await writeDocument(data)
  const documentXml = writtenXml(encoded, DOCUMENT_XML)
  expect(documentXml).toContain('r:id="rId')
  expect(documentXml).toContain('w:anchor="section-1"')
  expect(writtenXml(encoded, 'word/_rels/document.xml.rels')).toContain(`Target="${target}" TargetMode="External"`)

  const reopened = await readDocument(encoded)
  expect(reopened.body?.customRanges).toEqual(data.body?.customRanges)
  expect(reopened.body?.dataStream).toBe(data.body?.dataStream)
  expect(validateDocumentStructure(reopened)).toEqual([])
})

it('maps header numbering and a header hyperlink with the header relationship part', async () => {
  const target = 'https://example.test/header'
  const bytes = packageOf({
    '[Content_Types].xml': '<Types/>',
    [DOCUMENT_XML]: `<w:document ${DOCUMENT_NAMESPACE}><w:body><w:p/>`
      + '<w:sectPr><w:headerReference w:type="default" r:id="rId2"/></w:sectPr>'
      + '</w:body></w:document>',
    'word/_rels/document.xml.rels': '<Relationships>'
      + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"'
      + ' Target="header1.xml"/>'
      + '</Relationships>',
    'word/header1.xml': `<w:hdr ${DOCUMENT_NAMESPACE}>`
      + numberedParagraph('Header item', 0)
      + `<w:p>${hyperlink('r:id="rId1"', 'header link')}</w:p></w:hdr>`,
    'word/_rels/header1.xml.rels': '<Relationships>'
      + `<Relationship Id="rId1" ${HYPERLINK_RELATIONSHIP} Target="${target}" TargetMode="External"/>`
      + '</Relationships>',
    [NUMBERING_XML]: twoLevelNumbering(),
  })
  const data = await readDocument(bytes)
  const header = data.headers?.header1?.body
  expect(header?.paragraphs?.[0]?.bullet?.listType).toBe('numbering-1')
  expect(header?.customRanges?.[0]?.properties).toEqual({ url: target })

  const encoded = await writeDocument(data)
  const headerXml = writtenXml(encoded, 'word/header1.xml')
  expect(headerXml).toContain('<w:numPr>')
  expect(headerXml).toContain('<w:hyperlink r:id="rId1">')
  expect(writtenXml(encoded, 'word/_rels/header1.xml.rels')).toContain(`Target="${target}"`)

  const reopened = await readDocument(encoded)
  expect(reopened.headers?.header1?.body.customRanges?.[0]?.properties).toEqual({ url: target })
  expect(reopened.headers?.header1?.body.paragraphs?.[0]?.bullet?.listType).toBe('numbering-1')
  expect(validateDocumentStructure(reopened)).toEqual([])
})

it('synthesizes a list definition when a paragraph references missing numbering', async () => {
  const bytes = documentPackage(numberedParagraph('Item', 0, 7))
  const data = await readDocument(bytes)
  expect(data.lists?.['numbering-7']?.listType).toBe('numbering-7')
  expect(data.body?.paragraphs?.[0]?.bullet).toMatchObject({ listType: 'numbering-7', nestingLevel: 0 })
  const encoded = await writeDocument(data)
  expect(writtenXml(encoded, DOCUMENT_XML)).toContain('<w:numId w:val="1"/>')
  expect(writtenXml(encoded, NUMBERING_XML)).toContain('<w:numFmt w:val="decimal"/>')
  expect((await readDocument(encoded)).body?.paragraphs?.[0]?.bullet?.listType).toBe('numbering-1')
})

it('ignores malformed numbering instances and unnumbered paragraphs', async () => {
  const numbering = `<w:numbering ${DOCUMENT_NAMESPACE}>`
    + '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
    + '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>'
    + '</w:lvl></w:abstractNum>'
    + '<w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"/>'
    + '<w:lvlOverride><w:startOverride w:val="5"/></w:lvlOverride>'
    + '<w:lvlOverride w:ilvl="5"><w:startOverride w:val="2"/></w:lvlOverride></w:num>'
    + '<w:num/><w:abstractNum><w:lvl w:ilvl="0"/></w:abstractNum></w:numbering>'
  const bytes = documentPackage(
    '<w:p><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>none</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>missing</w:t></w:r></w:p>'
    + numberedParagraph('item', 0),
    { [NUMBERING_XML]: numbering },
  )
  const data = await readDocument(bytes)
  expect(data.body?.paragraphs?.[0]?.bullet).toBeUndefined()
  expect(data.body?.paragraphs?.[1]?.bullet).toBeUndefined()
  expect(data.body?.paragraphs?.[2]?.bullet?.listType).toBe('numbering-1')
})

it('writes a bullet whose list definition is absent from the snapshot', async () => {
  const data: IDocumentData = {
    id: 'orphan-list',
    documentStyle: {},
    body: {
      dataStream: 'Item\r\n',
      paragraphs: [{
        startIndex: 4,
        paragraphId: 'p1',
        bullet: { listType: 'numbering-9', listId: 'num-9', nestingLevel: 0 },
      }],
      customRanges: [
        { startIndex: 0, endIndex: 0, rangeId: 'bookmark-1', rangeType: CustomRangeType.CUSTOM },
        { startIndex: 0, endIndex: 0, rangeId: 'link-1', rangeType: CustomRangeType.HYPERLINK, properties: {} },
      ],
    },
  }
  const encoded = await writeDocument(data)
  expect(writtenXml(encoded, DOCUMENT_XML)).toContain('<w:numPr>')
  expect(writtenXml(encoded, NUMBERING_XML)).toContain('<w:numFmt w:val="decimal"/>')
})

it('writes no numbering part for a snapshot without lists', async () => {
  const data: IDocumentData = {
    id: 'plain',
    documentStyle: {},
    body: { dataStream: 'text\r\n', paragraphs: [{ startIndex: 4, paragraphId: 'p1' }] },
  }
  const encoded = await writeDocument(data)
  expect(Object.keys(unzipSync(encoded))).not.toContain(NUMBERING_XML)
  expect(writtenXml(encoded, DOCUMENT_XML)).not.toContain('w:numPr')
})
