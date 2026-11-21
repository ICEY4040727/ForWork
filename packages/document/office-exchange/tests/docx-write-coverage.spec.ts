/** Writer branch coverage for the DOCX provider: formatting, tables, sections, lists, preserved state. */
import {
  BaselineOffset,
  BooleanNumber,
  BorderStyleTypes,
  BulletAlignment,
  ColumnSeparatorType,
  CustomRangeType,
  DashStyleType,
  HorizontalAlign,
  ListGlyphType,
  NamedStyleType,
  NumberUnitType,
  PageOrientType,
  SectionType,
  SpacingRule,
  TableAlignmentType,
  TableRowHeightRule,
  TabStopAlignment,
  TabStopLeader,
  TextDecoration,
  VerticalAlignmentType,
} from '@univerjs/core'
import type {
  ICustomRange,
  IDocumentData,
  INestingLevel,
  IParagraph,
  IParagraphBorder,
  IParagraphStyle,
  ITable,
  ITableCell,
  ITextRun,
  ITextStyle,
} from '@univerjs/core'
import { strFromU8, unzipSync } from 'fflate'
import { expect, it } from 'vitest'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from '../src/config.ts'
import { writeDocument } from '../src/provider/docx.ts'

const CONFIG = DEFAULT_OFFICE_EXCHANGE_CONFIG

interface OfficeExchangeCarrier extends IDocumentData {
  officeExchange?: unknown
}

function documentOf(
  stream: string,
  body: Omit<NonNullable<IDocumentData['body']>, 'dataStream'> = {},
  extra: Omit<IDocumentData, 'id' | 'documentStyle' | 'body'> = {},
): IDocumentData {
  return { id: 'writer', documentStyle: {}, body: { dataStream: stream, ...body }, ...extra }
}

async function wrote(data: IDocumentData): Promise<Record<string, Uint8Array>> {
  return unzipSync(await writeDocument(data, CONFIG))
}

function documentXml(entries: Record<string, Uint8Array>): string {
  return strFromU8(entries['word/document.xml'] ?? new Uint8Array())
}

function stylesXml(entries: Record<string, Uint8Array>): string {
  return strFromU8(entries['word/styles.xml'] ?? new Uint8Array())
}

it('writes every paragraph alignment, spacing, indentation, tab, and border variant', async () => {
  const alignments = [
    HorizontalAlign.LEFT,
    HorizontalAlign.CENTER,
    HorizontalAlign.RIGHT,
    HorizontalAlign.BOTH,
    HorizontalAlign.JUSTIFIED,
    HorizontalAlign.DISTRIBUTED,
    999 as HorizontalAlign,
  ]
  const formatting: IParagraphStyle[] = alignments.map((horizontalAlign, index): IParagraphStyle => ({
    horizontalAlign,
    ...(index === 0 ? { spaceAbove: { v: 4 }, spaceBelow: { v: 8 } } : {}),
    ...(index === 1 ? { lineSpacing: 2, spacingRule: SpacingRule.AUTO } : {}),
    ...(index === 2 ? { lineSpacing: 20, spacingRule: SpacingRule.AT_LEAST } : {}),
    ...(index === 3 ? { lineSpacing: 20, spacingRule: SpacingRule.EXACT } : {}),
    ...(index === 4 ? { indentStart: { v: 8 }, indentEnd: { v: 16 } } : {}),
    ...(index === 5 ? { indentFirstLine: { v: 12 }, hanging: { v: 24 } } : {}),
    ...(index === 6
      ? {
        tabStops: [
          { offset: 10, alignment: TabStopAlignment.CENTER, leader: TabStopLeader.DOT },
          { offset: 20, alignment: 999 as TabStopAlignment },
          { offset: 30, alignment: TabStopAlignment.START, clear: true },
          { offset: 40, alignment: TabStopAlignment.START, leader: TabStopLeader.NONE },
        ],
        shading: { backgroundColor: { rgb: '#ABCDEF' } },
        outlineLevel: 1,
        borderTop: border('top'),
        borderBottom: border('bottom'),
        borderLeft: border('left'),
        borderRight: border('right'),
        borderBetween: border('between'),
      }
      : {}),
  }))
  const paragraphs: IParagraph[] = formatting.map((paragraphStyle, index) => ({
    startIndex: index * 2 + 1,
    paragraphId: `p${String(index)}`,
    paragraphStyle,
  }))
  const stream = `${'x\r'.repeat(formatting.length)}\n`
  const xml = documentXml(await wrote(documentOf(stream, { paragraphs })))
  expect(xml).toContain('<w:jc w:val="center"/>')
  expect(xml).toContain('<w:jc w:val="right"/>')
  expect(xml).toContain('<w:jc w:val="both"/>')
  expect(xml).toContain('<w:jc w:val="distribute"/>')
  expect(xml).toContain('<w:jc w:val="left"/>')
  expect(xml).toContain('<w:spacing w:before="60" w:after="120"/>')
  expect(xml).toContain('w:lineRule="atLeast"')
  expect(xml).toContain('w:lineRule="exact"')
  expect(xml).toContain('w:lineRule="auto"')
  expect(xml).toContain('<w:ind w:left="120" w:right="240"/>')
  expect(xml).toContain('w:firstLine="180" w:hanging="360"')
  expect(xml).toContain('<w:tab w:val="center" w:pos="150" w:leader="dot"/>')
  expect(xml).toContain('<w:tab w:val="left" w:pos="300"/>')
  expect(xml).toContain('<w:tab w:val="clear" w:pos="450"/>')
  expect(xml).toContain('<w:pBdr>')
  expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="ABCDEF"/>')
  expect(xml).toContain('<w:outlineLvl w:val="1"/>')
})

function border(tag: string): IParagraphBorder {
  return {
    color: { rgb: tag === 'top' ? '#GGGGGG' : '#FF0000' },
    width: tag === 'bottom' ? 2 : 4 / 6,
    dashStyle: tag === 'left' ? 999 as DashStyleType : DashStyleType.DOT,
    padding: tag === 'top' ? -1 : 2,
  }
}

it('writes every run property variant and fallback', async () => {
  const styles: (ITextStyle | undefined)[] = [
    { ff: 'Arial', eastAsiaFontFamily: 'SimSun' },
    { eastAsiaFontFamily: 'SimSun' },
    { bl: BooleanNumber.TRUE, it: BooleanNumber.TRUE, caps: true, smallCaps: true },
    { st: { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE } },
    { st: { s: BooleanNumber.TRUE, t: TextDecoration.DOUBLE } },
    { cl: { rgb: '#FF0000' }, fs: 12 },
    { cl: { rgb: 'GGGGGG' } },
    { bg: { rgb: '#FFFF00' } },
    { bg: { rgb: '#123456' } },
    { ul: { s: BooleanNumber.TRUE } },
    { ul: { s: BooleanNumber.TRUE, t: TextDecoration.DOUBLE, cl: { rgb: '#00FF00' } } },
    { ul: { s: BooleanNumber.TRUE, t: 999 as TextDecoration } },
    { va: BaselineOffset.SUPERSCRIPT },
    { va: BaselineOffset.SUBSCRIPT },
    { va: BaselineOffset.NORMAL },
    { bd: { t: { s: 999 as BorderStyleTypes, cl: { rgb: 'GGGGGG' } } } },
    { bd: { r: { s: BorderStyleTypes.THICK, cl: { rgb: '#0000FF' } } } },
    { bd: { b: { s: BorderStyleTypes.DOTTED, cl: { rgb: '#0000FF' } } } },
    { bd: { l: { s: BorderStyleTypes.DOUBLE, cl: { rgb: '#0000FF' } } } },
    {},
  ]
  const text = styles.map((_, index) => String.fromCodePoint(0x41 + index)).join('')
  const runs: ITextRun[] = []
  for (let index = 0; index < styles.length; index += 1) {
    const style = styles[index]
    if (style !== undefined) runs.push({ st: index, ed: index + 1, ts: style })
  }
  const xml = documentXml(await wrote(documentOf(
    `${text}\r\n`,
    { textRuns: runs, paragraphs: [{ startIndex: text.length, paragraphId: 'p1' }] },
  )))
  expect(xml).toContain('w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="SimSun"')
  expect(xml).toContain('<w:b/><w:i/><w:caps/><w:smallCaps/>')
  expect(xml).toContain('<w:strike/>')
  expect(xml).toContain('<w:dstrike/>')
  expect(xml).toContain('<w:color w:val="FF0000"/>')
  expect(xml).toContain('<w:sz w:val="24"/><w:szCs w:val="24"/>')
  expect(xml).toContain('<w:highlight w:val="yellow"/>')
  expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="123456"/>')
  expect(xml).toContain('<w:u w:val="single"/>')
  expect(xml).toContain('<w:u w:val="double" w:color="00FF00"/>')
  expect(xml).toContain('<w:vertAlign w:val="superscript"/>')
  expect(xml).toContain('<w:vertAlign w:val="subscript"/>')
  expect(xml).toContain('<w:vertAlign w:val="baseline"/>')
  expect(xml).toContain('<w:bdr w:val="single" w:sz="8" w:space="1" w:color="000000"/>')
  expect(xml).toContain('<w:bdr w:val="thick"')
  expect(xml).toContain('<w:bdr w:val="dotted"')
  expect(xml).toContain('<w:bdr w:val="double"')
})

it('writes bullets, list starts, and header and footer paragraph lists', async () => {
  const levels: INestingLevel[] = [
    {
      glyphType: ListGlyphType.BULLET,
      glyphFormat: '',
      startNumber: 2,
      bulletAlignment: BulletAlignment.START,
      isLegal: true,
      paragraphProperties: { horizontalAlign: HorizontalAlign.CENTER },
      textStyle: { bl: 1 },
    },
    { glyphType: ListGlyphType.DECIMAL, glyphFormat: '', startNumber: 0, bulletAlignment: BulletAlignment.START },
    { glyphType: 999 as ListGlyphType, glyphFormat: '%3.', startNumber: 1, bulletAlignment: BulletAlignment.START },
  ]
  const data = documentOf('abc\r\n', {
    paragraphs: [{ startIndex: 3, paragraphId: 'p1', bullet: { listType: 'list-a', listId: 'l1', nestingLevel: 0, startNumber: 3 } }],
  }, {
    lists: { 'list-a': { listType: 'list-a', nestingLevel: levels } },
    headers: {
      header1: {
        headerId: 'header1',
        body: {
          dataStream: 'h\r\n',
          paragraphs: [{ startIndex: 1, paragraphId: 'h1', bullet: { listType: 'list-b', listId: 'list-b', nestingLevel: 1 } }],
        },
      },
    },
    footers: {
      footer1: {
        footerId: 'footer1',
        body: {
          dataStream: 'f\r\n',
        },
      },
    },
  })
  const entries = await wrote(data)
  const numbering = strFromU8(entries['word/numbering.xml'] ?? new Uint8Array())
  expect(numbering).toContain('<w:numFmt w:val="bullet"/>')
  expect(numbering).toContain('<w:isLgl/>')
  expect(numbering).toContain('<w:pPr><w:jc w:val="center"/></w:pPr>')
  expect(numbering).toContain('<w:rPr><w:b/></w:rPr>')
  expect(numbering).toContain('<w:numFmt w:val="decimal"/>')
  expect(numbering).toContain('<w:lvlText w:val="%3."/>')
  expect(numbering).toContain('<w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride>')
  expect(documentXml(entries)).toContain('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
  expect(Object.keys(entries)).toEqual(expect.arrayContaining(['word/header1.xml', 'word/footer1.xml']))
  expect(entries['word/footer1.xml']).toBeDefined()
})

it('writes table width, alignment, layout, row height, cell spans, borders, and margins', async () => {
  const stream = '\x1A\x1B\x1C\r\x1D\x1C\r\x1D\x0E\x1B\x1C\r\x1D\x0E\x0F'
  const table: ITable = {
    tableId: 'table-1',
    tableRows: [
      {
        tableCells: [cellOne(), cellTwo()],
        trHeight: { val: { v: 20 }, hRule: TableRowHeightRule.EXACT },
        cantSplit: BooleanNumber.TRUE,
        repeatHeaderRow: BooleanNumber.TRUE,
      },
      {
        tableCells: [cellThree()],
        trHeight: { val: { v: 10 }, hRule: TableRowHeightRule.AT_LEAST },
      },
    ],
    tableColumns: [
      { size: { type: 1, width: { v: 80 } } },
      { size: { type: 1, width: { v: 120 } } },
    ],
    align: TableAlignmentType.CENTER,
    indent: { v: 8 },
    textWrap: 0,
    position: { positionH: { relativeFrom: 0, posOffset: 0 }, positionV: { relativeFrom: 0, posOffset: 0 } },
    dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
    cellMargin: { start: { v: 10 }, end: { v: 10 }, top: { v: 5 }, bottom: { v: 5 } },
    size: { type: 0, width: { v: 200 } },
    layout: 0,
  }
  const data = documentOf(stream, {
    paragraphs: [],
    tables: [{ startIndex: 0, endIndex: stream.length, tableId: 'table-1' }],
  }, { tableSource: { 'table-1': table } })
  const xml = documentXml(await wrote(data))
  expect(xml).toContain('<w:tblW w:w="0" w:type="auto"/>')
  expect(xml).toContain('<w:jc w:val="center"/>')
  expect(xml).toContain('<w:tblInd w:w="120" w:type="dxa"/>')
  expect(xml).toContain('<w:tblLayout w:type="autofit"/>')
  expect(xml).toContain('<w:gridSpan w:val="2"/>')
  expect(xml).toContain('<w:trHeight w:val="300" w:hRule="exact"/>')
  expect(xml).toContain('<w:cantSplit/>')
  expect(xml).toContain('<w:tblHeader/>')
  expect(xml).toContain('<w:trHeight w:val="150" w:hRule="atLeast"/>')
  expect(xml).toContain('<w:tcBorders>')
  expect(xml).toContain('<w:tcMar>')
  expect(xml).toContain('<w:vAlign w:val="bottom"/>')
})

function cellOne(): ITableCell {
  return {
    columnSpan: 2,
    rowSpan: 0,
    backgroundColor: { rgb: '#ABCDEF' },
    borderTop: { color: { rgb: '#GGGGGG' }, dashStyle: 999 as DashStyleType },
    borderBottom: { color: { rgb: '#112233' }, width: { v: 2 }, dashStyle: DashStyleType.DOT },
    borderLeft: { color: { rgb: '#112233' }, dashStyle: DashStyleType.DASH },
    borderRight: { color: { rgb: '#112233' }, dashStyle: DashStyleType.SOLID },
    margin: {
      start: { v: 20 }, end: { v: 10 }, top: { v: 5 }, bottom: { v: 5 },
    },
    vAlign: VerticalAlignmentType.BOTTOM,
  }
}

function cellTwo(): ITableCell {
  return {
    columnSpan: 2.5,
    rowSpan: 2,
    backgroundColor: { rgb: 'not-a-color' },
    margin: { start: { v: 10 }, end: { v: 20 }, top: { v: 7 }, bottom: { v: 9 } },
    vAlign: 999 as VerticalAlignmentType,
  }
}

function cellThree(): ITableCell {
  return { margin: { start: { v: 10 }, end: { v: 10 }, top: { v: 5 }, bottom: { v: 5 } } }
}

it('writes percent and dxa table widths and missing cell metadata', async () => {
  const percentStream = '\x1A\x1B\x1C\r\x1D\x0E\x0F'
  const percent: ITable = {
    ...baseTable(),
    size: { type: 1, width: { v: 50, u: NumberUnitType.PERCENT } },
    align: TableAlignmentType.END,
    layout: 1,
    indent: { v: 0 },
  }
  const xml = documentXml(await wrote(documentOf(percentStream, {
    tables: [{ startIndex: 0, endIndex: percentStream.length, tableId: 'table-1' }],
  }, { tableSource: { 'table-1': percent } })))
  expect(xml).toContain('<w:tblW w:w="2500" w:type="pct"/>')
  expect(xml).toContain('<w:jc w:val="end"/>')
  expect(xml).toContain('<w:tblLayout w:type="fixed"/>')
  expect(xml).not.toContain('<w:tblInd')

  const dxa = documentXml(await wrote(documentOf(percentStream, {
    tables: [{ startIndex: 0, endIndex: percentStream.length, tableId: 'table-1' }],
  })))
  expect(dxa).toContain('<w:tblW w:w="0" w:type="auto"/>')
})

function baseTable(): ITable {
  return {
    tableId: 'table-1',
    tableRows: [{ tableCells: [cellThree()], trHeight: { val: { v: 10 }, hRule: TableRowHeightRule.AT_LEAST } }],
    tableColumns: [],
    align: TableAlignmentType.CENTER,
    indent: { v: 0 },
    textWrap: 0,
    position: { positionH: { relativeFrom: 0, posOffset: 0 }, positionV: { relativeFrom: 0, posOffset: 0 } },
    dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
    cellMargin: { start: { v: 10 }, end: { v: 10 }, top: { v: 5 }, bottom: { v: 5 } },
    size: { type: 1, width: { v: 160 } },
  }
}

it('writes sections, columns, header and footer references, and settings', async () => {
  const data = documentOf('\r\n\r\n', {
    paragraphs: [{ startIndex: 0, paragraphId: 'p1' }, { startIndex: 2, paragraphId: 'p2' }],
    sectionBreaks: [
      {
        sectionId: 's2',
        startIndex: 2,
        pageSize: { width: 800, height: 600 },
        pageOrient: PageOrientType.LANDSCAPE,
        marginTop: 10,
        marginRight: 20,
        marginBottom: 30,
        marginLeft: 40,
        marginHeader: 5,
        marginFooter: 6,
        columnSeparatorType: ColumnSeparatorType.BETWEEN_EACH_COLUMN,
        columnProperties: [
          { width: 100, paddingEnd: 10 },
          { width: 200, paddingEnd: 0 },
        ],
        sectionType: SectionType.NEXT_PAGE,
        defaultHeaderId: 'header1',
        evenPageFooterId: 'footer1',
        useFirstPageHeaderFooter: BooleanNumber.TRUE,
      },
      {
        sectionId: 's1',
        startIndex: 1,
        columnSeparatorType: ColumnSeparatorType.BETWEEN_EACH_COLUMN,
        defaultHeaderId: 'missing',
      },
    ],
  }, {
    headers: {
      header1: {
        headerId: 'header1',
        body: { dataStream: 'h\r\n', paragraphs: [{ startIndex: 1, paragraphId: 'h1' }] },
      },
    },
    footers: {
      footer1: {
        footerId: 'footer1',
        body: { dataStream: 'f\r\n', paragraphs: [{ startIndex: 1, paragraphId: 'f1' }] },
      },
    },
  })
  data.documentStyle = { evenAndOddHeaders: BooleanNumber.TRUE }
  const entries = await wrote(data)
  const xml = documentXml(entries)
  expect(xml).toContain('<w:type w:val="nextPage"/>')
  expect(xml).toContain('w:orient="landscape"')
  expect(xml).toContain('<w:pgMar w:top="150" w:right="300" w:bottom="450" w:left="600" w:header="75" w:footer="90"/>')
  expect(xml).toContain('<w:cols w:num="2" w:space="150" w:sep="1">')
  expect(xml).toContain('<w:col w:w="1500" w:space="150"/>')
  expect(xml).toContain('<w:headerReference w:type="default"')
  expect(xml).toContain('<w:footerReference w:type="even"')
  expect(xml).toContain('<w:titlePg/>')
  expect(xml).toContain('<w:cols w:num="1" w:sep="1"/>')
  expect(entries['word/settings.xml']).toBeDefined()
  expect(entries['word/_rels/document.xml.rels']).toBeDefined()
})

it('writes section columns without widths and an empty section', async () => {
  const xml = documentXml(await wrote(documentOf('\r\n', {
    paragraphs: [{ startIndex: 0, paragraphId: 'p1' }],
    sectionBreaks: [{ sectionId: 's1', startIndex: 1, columnSeparatorType: ColumnSeparatorType.BETWEEN_EACH_COLUMN }],
  })))
  expect(xml).toContain('<w:cols w:num="1" w:sep="1"/>')
})

it('writes external and bookmark hyperlinks and ignores other custom ranges', async () => {
  const ranges: ICustomRange[] = [
    {
      startIndex: 0,
      endIndex: 1,
      rangeId: 'l1',
      rangeType: CustomRangeType.HYPERLINK,
      properties: { url: 'https://example.test/a' },
    },
    {
      startIndex: 1,
      endIndex: 2,
      rangeId: 'l2',
      rangeType: CustomRangeType.HYPERLINK,
      properties: { bookmarkId: 'mark' },
    },
    {
      startIndex: 2,
      endIndex: 3,
      rangeId: 'l3',
      rangeType: CustomRangeType.HYPERLINK,
      properties: {},
    },
    {
      startIndex: 3,
      endIndex: 4,
      rangeId: 'l4',
      rangeType: CustomRangeType.HYPERLINK,
      properties: { url: 5 },
    },
    {
      startIndex: 4,
      endIndex: 5,
      rangeId: 'l5',
      rangeType: CustomRangeType.COMMENT,
      properties: { url: 'https://example.test/ignored' },
    },
    {
      startIndex: 5,
      endIndex: 6,
      rangeId: 'l6',
      rangeType: CustomRangeType.HYPERLINK,
    },
    {
      startIndex: 6,
      endIndex: 6,
      rangeId: 'l7',
      rangeType: CustomRangeType.HYPERLINK,
      properties: { url: 'https://example.test/a' },
    },
  ]
  const entries = await wrote(documentOf('abcdef\r\n', {
    customRanges: ranges,
    paragraphs: [{ startIndex: 6, paragraphId: 'p1' }],
  }))
  const xml = documentXml(entries)
  expect(xml).toContain('<w:hyperlink r:id="rId2">')
  expect(xml).toContain('<w:hyperlink w:anchor="mark">')
  const rels = strFromU8(entries['word/_rels/document.xml.rels'] ?? new Uint8Array())
  expect(rels).toContain('Target="https://example.test/a" TargetMode="External"')
})

it('writes a named style and heading style fallback', async () => {
  const custom = documentXml(await wrote(documentOf('ab\r\n', {
    paragraphs: [{ startIndex: 2, paragraphId: 'p1', styleId: 'Custom' }],
  })))
  expect(custom).toContain('<w:pStyle w:val="Custom"/>')
  const heading = documentXml(await wrote(documentOf('ab\r\n', {
    paragraphs: [{ startIndex: 2, paragraphId: 'p2', paragraphStyle: { namedStyleType: NamedStyleType.HEADING_5 } }],
  })))
  expect(heading).toContain('<w:pStyle w:val="Heading5"/>')
  expect(stylesXml(await wrote(documentOf('a\r\n', {
    paragraphs: [{ startIndex: 1, paragraphId: 'p1' }],
  })))).toContain('w:styleId="Heading1"')
})

it('refuses a raw block outside a paragraph', async () => {
  await expect(writeDocument(documentOf('x\r\n', {
    paragraphs: [{ startIndex: 1, paragraphId: 'p1' }],
    docxRawBlocks: [{ startIndex: 99, xml: '<w:sdt/>' }],
  }), CONFIG)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses malformed preserved package state', async () => {
  const cases: readonly unknown[] = [
    5,
    { docx: 5 },
    { docx: { parts: 'x', relationships: {}, contentTypes: [] } },
    { docx: { parts: [{ name: 'x' }], relationships: {}, contentTypes: [] } },
    { docx: { parts: [], relationships: {}, contentTypes: [1] } },
    { docx: { parts: [], relationships: 5, contentTypes: [] } },
    { docx: { parts: [], relationships: { document: [1] }, contentTypes: [] } },
    { docx: { parts: [], relationships: { stories: 5 }, contentTypes: [] } },
    { docx: { parts: [], relationships: { stories: { header1: [1] } }, contentTypes: [] } },
  ]
  for (const officeExchange of cases) {
    const data = documentOf('x\r\n', { paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] })
    const carrier: OfficeExchangeCarrier = data
    carrier.officeExchange = officeExchange
    await expect(writeDocument(data, CONFIG)).rejects.toMatchObject({ code: 'conversion-failed' })
  }
})

it('refuses a preserved part that conflicts with a generated part', async () => {
  const data = documentOf('x\r\n', { paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] })
  const carrier: OfficeExchangeCarrier = data
  carrier.officeExchange = {
    docx: {
      parts: [{ name: 'word/document.xml', base64: 'AQ==' }],
      relationships: {},
      contentTypes: [],
    },
  }
  await expect(writeDocument(data, CONFIG)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('writes preserved media, notes, and content-type declarations', async () => {
  const data = documentOf('x\r\n', { paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] })
  const carrier: OfficeExchangeCarrier = data
  carrier.officeExchange = {
    docx: {
      parts: [
        { name: 'word/media/image1.png', base64: 'AQ==' },
        { name: 'word/media/noext', base64: 'AAE=' },
        { name: 'word/footnotes.xml', base64: 'AQID' },
        { name: 'word/endnotes.xml', base64: 'AQIDBA==' },
      ],
      relationships: {
        document: ['<Relationship Id="rId9" Type="image" Target="media/image1.png"/>'],
        stories: { header1: ['<Relationship Id="rId8" Type="image" Target="media/image1.png"/>'] },
      },
      contentTypes: [
        '<Default Extension="png" ContentType="image/png"/>',
        '<Override PartName="/word/footnotes.xml" ContentType="footnotes"/>',
        '<Override PartName="/word/document.xml" ContentType="main"/>',
      ],
    },
  }
  const entries = await wrote(data)
  const types = strFromU8(entries['[Content_Types].xml'] ?? new Uint8Array())
  expect(types).toContain('<Default Extension="png" ContentType="image/png"/>')
  expect(types).toContain('<Override PartName="/word/footnotes.xml" ContentType="footnotes"/>')
  expect(types).toContain('<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>')
  expect(types).toContain('<Override PartName="/word/document.xml" ContentType="main"/>')
  expect(entries['word/media/image1.png']).toBeDefined()
  expect(entries['word/media/noext']).toBeDefined()
  expect(entries['word/footnotes.xml']).toBeDefined()
  const rels = strFromU8(entries['word/_rels/document.xml.rels'] ?? new Uint8Array())
  expect(rels).toContain('Id="rId9"')
})

it('writes a media extension the source declares no default for', async () => {
  const data = documentOf('x\r\n', { paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] })
  const carrier: OfficeExchangeCarrier = data
  carrier.officeExchange = {
    docx: {
      parts: [{ name: 'word/media/image1.jpeg', base64: 'AQ==' }],
      relationships: {},
      contentTypes: [],
    },
  }
  const types = strFromU8((await wrote(data))['[Content_Types].xml'] ?? new Uint8Array())
  expect(types).toContain('<Default Extension="jpeg" ContentType="image/jpeg"/>')
})

it('refuses invalid base64 preserved parts', async () => {
  for (const base64 of ['A', '!!!!', 'AA==AAAA']) {
    const data = documentOf('x\r\n', { paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] })
    const carrier: OfficeExchangeCarrier = data
    carrier.officeExchange = {
      docx: {
        parts: [{ name: 'word/media/image1.png', base64 }],
        relationships: {},
        contentTypes: [],
      },
    }
    await expect(writeDocument(data, CONFIG)).rejects.toMatchObject({ code: 'conversion-failed' })
  }
})

it('writes preserved parts whose names carry no extension', async () => {
  const data = documentOf('x\r\n', { paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] })
  const carrier: OfficeExchangeCarrier = data
  carrier.officeExchange = {
    docx: {
      parts: [{ name: 'word/media/noext', base64: 'AQ==' }],
      relationships: {},
      contentTypes: [],
    },
  }
  const entries = await wrote(data)
  expect(entries['word/media/noext']).toEqual(new Uint8Array([1]))
})
