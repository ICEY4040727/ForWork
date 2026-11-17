/**
 * DOCX run, paragraph, and styles.xml formatting: OOXML-to-Univer mapping and
 * the write/read round-trip that keeps every mapped property.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { expect, it } from 'vitest'
import {
  BaselineOffset,
  BorderStyleTypes,
  BooleanNumber,
  DashStyleType,
  HorizontalAlign,
  SpacingRule,
  TabStopAlignment,
  TabStopLeader,
  TextDecoration,
  type IDocumentData,
  type IParagraph,
} from '@univerjs/core'
import { readDocument, writeDocument } from '../src/provider/docx.ts'

const DOCUMENT_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const DOCUMENT_XML = 'word/document.xml'
const STYLES_XML = 'word/styles.xml'

function packageOf(entries: Readonly<Record<string, string>>): Uint8Array {
  const files: Record<string, Uint8Array> = {}
  for (const [name, content] of Object.entries(entries)) files[name] = strToU8(content)
  return zipSync(files)
}

function documentPackage(body: string, styles?: string): Uint8Array {
  const entries: Record<string, string> = {
    '[Content_Types].xml': '<Types/>',
    [DOCUMENT_XML]: `<w:document ${DOCUMENT_NAMESPACE}><w:body>${body}</w:body></w:document>`,
  }
  if (styles !== undefined) entries[STYLES_XML] = `<w:styles ${DOCUMENT_NAMESPACE}>${styles}</w:styles>`
  return packageOf(entries)
}

function textRunStyle(data: IDocumentData, text: string): object | undefined {
  const stream = data.body?.dataStream ?? ''
  return data.body?.textRuns?.find(run => stream.slice(run.st, run.ed) === text)?.ts
}

function firstParagraph(data: IDocumentData): IParagraph | undefined {
  return data.body?.paragraphs?.[0]
}

function writtenXml(bytes: Uint8Array, part: string): string {
  const entries = unzipSync(bytes)
  const content = entries[part]
  return content === undefined ? '' : strFromU8(content)
}

it('maps run formatting into the text run style', async () => {
  const bytes = documentPackage(
    '<w:p><w:r><w:rPr>'
    + '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="SimSun"/>'
    + '<w:sz w:val="24"/><w:color w:val="FF0000"/><w:highlight w:val="yellow"/>'
    + '<w:dstrike/><w:vertAlign w:val="superscript"/><w:caps/><w:smallCaps/>'
    + '<w:bdr w:val="double" w:sz="8" w:color="00FF00"/>'
    + '<w:u w:val="double" w:color="0000FF"/>'
    + '</w:rPr><w:t>styled</w:t></w:r></w:p>',
  )
  const data = await readDocument(bytes)
  expect(textRunStyle(data, 'styled')).toEqual({
    ff: 'Arial',
    eastAsiaFontFamily: 'SimSun',
    fs: 12,
    cl: { rgb: '#FF0000' },
    bg: { rgb: '#FFFF00' },
    st: { s: BooleanNumber.TRUE, t: TextDecoration.DOUBLE },
    va: BaselineOffset.SUPERSCRIPT,
    caps: true,
    smallCaps: true,
    bd: {
      t: { s: BorderStyleTypes.DOUBLE, cl: { rgb: '#00FF00' } },
      r: { s: BorderStyleTypes.DOUBLE, cl: { rgb: '#00FF00' } },
      b: { s: BorderStyleTypes.DOUBLE, cl: { rgb: '#00FF00' } },
      l: { s: BorderStyleTypes.DOUBLE, cl: { rgb: '#00FF00' } },
    },
    ul: { s: BooleanNumber.TRUE, t: TextDecoration.DOUBLE, cl: { rgb: '#0000FF' } },
  })
})

it('round-trips mapped run formatting', async () => {
  const bytes = documentPackage(
    '<w:p><w:r><w:rPr>'
    + '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="SimSun"/>'
    + '<w:sz w:val="24"/><w:color w:val="FF0000"/><w:highlight w:val="yellow"/>'
    + '<w:dstrike/><w:vertAlign w:val="superscript"/><w:caps/><w:smallCaps/>'
    + '<w:bdr w:val="double" w:sz="8" w:color="00FF00"/>'
    + '<w:u w:val="double" w:color="0000FF"/>'
    + '</w:rPr><w:t>styled</w:t></w:r></w:p>',
  )
  const original = await readDocument(bytes)
  const reopened = await readDocument(await writeDocument(original))
  expect(textRunStyle(reopened, 'styled')).toEqual(textRunStyle(original, 'styled'))
})

it('maps additional run property variants', async () => {
  const bytes = documentPackage(
    '<w:p>'
    + '<w:r><w:rPr><w:rFonts w:hAnsi="Tahoma"/><w:szCs w:val="18"/>'
    + '<w:shd w:val="clear" w:color="auto" w:fill="ABCDEF"/></w:rPr><w:t>fill</w:t></w:r>'
    + '<w:r><w:rPr><w:rFonts w:cs="Courier New"/><w:highlight w:val="none"/></w:rPr><w:t>clear</w:t></w:r>'
    + '<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>sub</w:t></w:r>'
    + '<w:r><w:rPr><w:vertAlign w:val="baseline"/></w:rPr><w:t>baseline</w:t></w:r>'
    + '<w:r><w:rPr><w:strike w:val="false"/></w:rPr><w:t>plain</w:t></w:r>'
    + '<w:r><w:rPr><w:u w:val="wavyHeavy"/></w:rPr><w:t>wavy</w:t></w:r>'
    + '<w:r><w:rPr><w:u/></w:rPr><w:t>plain-underline</w:t></w:r>'
    + '<w:r><w:rPr><w:bdr w:val="bogus"/></w:rPr><w:t>bordered</w:t></w:r>'
    + '</w:p>',
  )
  const data = await readDocument(bytes)
  expect(textRunStyle(data, 'fill')).toEqual({
    ff: 'Tahoma',
    fs: 9,
    bg: { rgb: '#ABCDEF' },
  })
  expect(textRunStyle(data, 'clear')).toEqual({ ff: 'Courier New' })
  expect(textRunStyle(data, 'sub')).toEqual({ va: BaselineOffset.SUBSCRIPT })
  expect(textRunStyle(data, 'baseline')).toEqual({ va: BaselineOffset.NORMAL })
  expect(textRunStyle(data, 'plain')).toBeUndefined()
  expect(textRunStyle(data, 'wavy')).toEqual({
    ul: { s: BooleanNumber.TRUE, t: TextDecoration.WAVY_HEAVY },
  })
  expect(textRunStyle(data, 'plain-underline')).toEqual({
    ul: { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE },
  })
  expect(textRunStyle(data, 'bordered')).toEqual({
    bd: {
      t: { s: BorderStyleTypes.THIN, cl: { rgb: '#000000' } },
      r: { s: BorderStyleTypes.THIN, cl: { rgb: '#000000' } },
      b: { s: BorderStyleTypes.THIN, cl: { rgb: '#000000' } },
      l: { s: BorderStyleTypes.THIN, cl: { rgb: '#000000' } },
    },
  })
})

it('round-trips additional run property variants', async () => {
  const bytes = documentPackage(
    '<w:p>'
    + '<w:r><w:rPr><w:rFonts w:hAnsi="Tahoma"/><w:szCs w:val="18"/>'
    + '<w:shd w:val="clear" w:color="auto" w:fill="ABCDEF"/></w:rPr><w:t>fill</w:t></w:r>'
    + '<w:r><w:rPr><w:vertAlign w:val="baseline"/></w:rPr><w:t>baseline</w:t></w:r>'
    + '<w:r><w:rPr><w:u w:val="wavyHeavy"/></w:rPr><w:t>wavy</w:t></w:r>'
    + '</w:p>',
  )
  const original = await readDocument(bytes)
  const reopened = await readDocument(await writeDocument(original))
  for (const text of ['fill', 'baseline', 'wavy']) {
    expect(textRunStyle(reopened, text)).toEqual(textRunStyle(original, text))
  }
})

it('maps paragraph formatting into the paragraph style', async () => {
  const bytes = documentPackage(
    '<w:p><w:pPr>'
    + '<w:jc w:val="both"/>'
    + '<w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>'
    + '<w:ind w:left="720" w:right="360" w:firstLine="240" w:hanging="0"/>'
    + '<w:tabs><w:tab w:val="center" w:pos="1440" w:leader="dot"/>'
    + '<w:tab w:val="right" w:pos="2880"/><w:tab w:val="clear" w:pos="480"/></w:tabs>'
    + '<w:pBdr><w:top w:val="single" w:sz="8" w:space="6" w:color="FF0000"/>'
    + '<w:bottom w:val="dotted" w:sz="4" w:space="0" w:color="auto"/>'
    + '<w:left w:val="none"/></w:pBdr>'
    + '<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/>'
    + '<w:outlineLvl w:val="0"/>'
    + '</w:pPr><w:r><w:t>styled</w:t></w:r></w:p>',
  )
  const data = await readDocument(bytes)
  expect(firstParagraph(data)?.paragraphStyle).toEqual({
    horizontalAlign: HorizontalAlign.BOTH,
    spaceAbove: { v: 8 },
    spaceBelow: { v: 16 },
    lineSpacing: 1.5,
    spacingRule: SpacingRule.AUTO,
    indentStart: { v: 48 },
    indentEnd: { v: 24 },
    indentFirstLine: { v: 16 },
    hanging: { v: 0 },
    tabStops: [
      { offset: 96, alignment: TabStopAlignment.CENTER, leader: TabStopLeader.DOT },
      { offset: 192, alignment: TabStopAlignment.END },
      { offset: 32, alignment: TabStopAlignment.START, clear: true },
    ],
    borderTop: {
      color: { rgb: '#FF0000' },
      width: 8 / 6,
      dashStyle: DashStyleType.SOLID,
      padding: 8,
    },
    borderBottom: {
      color: { rgb: '#000000' },
      width: 4 / 6,
      dashStyle: DashStyleType.DOT,
      padding: 0,
    },
    shading: { backgroundColor: { rgb: '#EEEEEE' } },
    outlineLevel: 0,
  })
})

it('round-trips mapped paragraph formatting', async () => {
  const bytes = documentPackage(
    '<w:p><w:pPr>'
    + '<w:jc w:val="both"/>'
    + '<w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>'
    + '<w:ind w:left="720" w:right="360" w:firstLine="240" w:hanging="0"/>'
    + '<w:tabs><w:tab w:val="center" w:pos="1440" w:leader="dot"/>'
    + '<w:tab w:val="right" w:pos="2880"/><w:tab w:val="clear" w:pos="480"/></w:tabs>'
    + '<w:pBdr><w:top w:val="single" w:sz="8" w:space="6" w:color="FF0000"/>'
    + '<w:bottom w:val="dotted" w:sz="4" w:space="0" w:color="auto"/>'
    + '<w:left w:val="none"/></w:pBdr>'
    + '<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/>'
    + '<w:outlineLvl w:val="0"/>'
    + '</w:pPr><w:r><w:t>styled</w:t></w:r></w:p>',
  )
  const original = await readDocument(bytes)
  const reopened = await readDocument(await writeDocument(original))
  expect(firstParagraph(reopened)?.paragraphStyle).toEqual(firstParagraph(original)?.paragraphStyle)
})

it('maps paragraph alignment, line rules, indents, and skipped tab stops', async () => {
  const bytes = documentPackage(
    '<w:p><w:pPr><w:jc w:val="right"/>'
    + '<w:spacing w:line="360" w:lineRule="atLeast"/></w:pPr><w:r><w:t>right</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:jc w:val="left"/>'
    + '<w:spacing w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>left</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:jc w:val="bogus"/>'
    + '<w:ind w:start="300" w:end="450"/>'
    + '<w:tabs><w:tab w:val="decimal" w:pos="600"/><w:tab/></w:tabs>'
    + '<w:shd w:val="clear"/></w:pPr><w:r><w:t>start</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:spacing w:line="480"/></w:pPr><w:r><w:t>default rule</w:t></w:r></w:p>',
  )
  const data = await readDocument(bytes)
  const paragraphs = data.body?.paragraphs ?? []
  expect(paragraphs[0]?.paragraphStyle).toEqual({
    horizontalAlign: HorizontalAlign.RIGHT,
    lineSpacing: 24,
    spacingRule: SpacingRule.AT_LEAST,
  })
  expect(paragraphs[1]?.paragraphStyle).toEqual({
    horizontalAlign: HorizontalAlign.LEFT,
    lineSpacing: 16,
    spacingRule: SpacingRule.EXACT,
  })
  expect(paragraphs[2]?.paragraphStyle).toEqual({
    indentStart: { v: 20 },
    indentEnd: { v: 30 },
    tabStops: [{ offset: 40, alignment: TabStopAlignment.START }],
  })
  expect(paragraphs[3]?.paragraphStyle).toEqual({
    lineSpacing: 2,
    spacingRule: SpacingRule.AUTO,
  })
})

it('round-trips atLeast and exact line rules', async () => {
  const bytes = documentPackage(
    '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="atLeast"/></w:pPr><w:r><w:t>one</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>two</w:t></w:r></w:p>',
  )
  const original = await readDocument(bytes)
  const reopened = await readDocument(await writeDocument(original))
  expect(reopened.body?.paragraphs?.map(paragraph => paragraph.paragraphStyle))
    .toEqual(original.body?.paragraphs?.map(paragraph => paragraph.paragraphStyle))
})

it('reads docDefaults and the Normal style into the document style', async () => {
  const bytes = documentPackage(
    '<w:p><w:r><w:t>text</w:t></w:r></w:p>',
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia"/><w:sz w:val="20"/>'
    + '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault>'
    + '</w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    + '<w:rPr><w:color w:val="333333"/></w:rPr></w:style>',
  )
  const data = await readDocument(bytes)
  expect(data.documentStyle).toEqual({
    textStyle: { ff: 'Georgia', fs: 10, cl: { rgb: '#333333' } },
    defaultParagraphStyle: { spaceBelow: { v: 8 } },
  })
})

it('round-trips document defaults through styles.xml', async () => {
  const bytes = documentPackage(
    '<w:p><w:r><w:t>text</w:t></w:r></w:p>',
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia"/><w:sz w:val="20"/>'
    + '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault>'
    + '</w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    + '<w:rPr><w:color w:val="333333"/></w:rPr></w:style>',
  )
  const original = await readDocument(bytes)
  const reopened = await readDocument(await writeDocument(original))
  expect(reopened.documentStyle).toEqual(original.documentStyle)
})

it('reads Normal style defaults without docDefaults', async () => {
  const bytes = documentPackage(
    '<w:p><w:r><w:t>text</w:t></w:r></w:p>',
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="x"/></w:style>'
    + '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/>'
    + '<w:pPr><w:spacing w:before="60"/></w:pPr></w:style>',
  )
  expect((await readDocument(bytes)).documentStyle).toEqual({
    defaultParagraphStyle: { spaceAbove: { v: 4 } },
  })
})

it('reads an empty or missing styles part as no document defaults', async () => {
  const withoutStyles = await readDocument(documentPackage('<w:p><w:r><w:t>text</w:t></w:r></w:p>'))
  expect(withoutStyles.documentStyle).toEqual({})
  const emptyStyles = await readDocument(documentPackage('<w:p/>', ''))
  expect(emptyStyles.documentStyle).toEqual({})
  const otherRoot = packageOf({
    [DOCUMENT_XML]: `<w:document ${DOCUMENT_NAMESPACE}><w:body><w:p/></w:body></w:document>`,
    [STYLES_XML]: '<w:other/>',
  })
  expect((await readDocument(otherRoot)).documentStyle).toEqual({})
})

it('rejects malformed styles XML', async () => {
  const bytes = packageOf({
    [DOCUMENT_XML]: `<w:document ${DOCUMENT_NAMESPACE}><w:body><w:p/></w:body></w:document>`,
    [STYLES_XML]: '<w:styles><w:docDefaults>',
  })
  await expect(readDocument(bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('writes paragraph formatting from an edited snapshot', async () => {
  const data: IDocumentData = {
    id: 'paragraph-formatting',
    documentStyle: {},
    body: {
      dataStream: 'text\r\n',
      paragraphs: [{
        startIndex: 4,
        paragraphId: 'paragraph-1',
        paragraphStyle: {
          horizontalAlign: HorizontalAlign.JUSTIFIED,
          lineSpacing: 24,
          spacingRule: SpacingRule.AT_LEAST,
          indentStart: { v: 0 },
          indentEnd: { v: 24 },
          indentFirstLine: { v: 16 },
          hanging: { v: 8 },
          tabStops: [],
          borderRight: {
            color: { rgb: '#112233' },
            width: 2,
            dashStyle: DashStyleType.DASH,
            padding: 6,
          },
          outlineLevel: 1,
        },
      }],
    },
  }
  const paragraphXml = writtenXml(await writeDocument(data), DOCUMENT_XML)
  expect(paragraphXml).toContain('<w:jc w:val="both"/>')
  expect(paragraphXml).toContain('w:line="360" w:lineRule="atLeast"')
  expect(paragraphXml).toContain('w:left="0" w:right="360" w:firstLine="240" w:hanging="120"')
  expect(paragraphXml).toContain('<w:right w:val="dashed" w:sz="12" w:space="5" w:color="112233"/>')
  expect(paragraphXml).toContain('<w:outlineLvl w:val="1"/>')
})
