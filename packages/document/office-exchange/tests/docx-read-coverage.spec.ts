/** Reader branch coverage for the DOCX provider: formatting, numbering, tables, sections, raw blocks. */
import {
  BaselineOffset,
  BulletAlignment,
  ColumnSeparatorType,
  DashStyleType,
  HorizontalAlign,
  ListGlyphType,
  NumberUnitType,
  SpacingRule,
  TableAlignmentType,
  TableLayoutType,
  PageOrientType,
  SectionType,
  TextDecoration,
  VerticalAlignmentType,
} from '@univerjs/core'
import { strToU8, zipSync } from 'fflate'
import { expect, it } from 'vitest'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from '../src/config.ts'
import type { IDocumentData, ITextRun } from '@univerjs/core'
import { readDocument, readDocumentWithSource } from '../src/provider/docx.ts'

const WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

type Entry = string | Uint8Array

interface ExtensionView {
  readonly source?: {
    readonly anchorMode: 'per-unit' | 'sheet-summary'
    readonly anchors: readonly { readonly kind: string }[]
  }
  readonly docx?: {
    readonly parts: readonly { readonly name: string }[]
    readonly contentTypes: readonly string[]
    readonly relationships: {
      readonly document?: readonly string[]
      readonly stories?: Readonly<Record<string, readonly string[]>>
    }
  }
}

function extensionView(data: IDocumentData): ExtensionView | undefined {
  return (data as IDocumentData & { readonly officeExchange?: ExtensionView }).officeExchange
}

const DOCUMENT_NS = `<w:document xmlns:w="${WORD}"`
  + ` xmlns:r="${REL}"`
  + ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
  + ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
  + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
  + ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
  + ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
  + ' xmlns:v="urn:schemas-microsoft-com:vml"'
  + ' xmlns:o="urn:schemas-microsoft-com:office:office"'
  + ' xmlns:w10="urn:schemas-microsoft-com:office:word"'
  + ' xmlns:x="' + WORD + '">'

const STYLES = `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="${WORD}">`
  + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr>'
  + '</w:rPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
  + '<w:pPr><w:jc w:val="left"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Named"><w:name w:val="Named"/></w:style>'
  + '</w:styles>'

const SETTINGS = `<w:settings xmlns:w="${WORD}"><w:evenAndOddHeaders/></w:settings>`

const ROOT_RELS = '<?xml version="1.0" encoding="UTF-8"?>'
  + `<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1"`
  + ` Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`

function packageOf(entries: Readonly<Record<string, Entry>>): Uint8Array {
  const files: Record<string, Uint8Array> = {}
  for (const [name, content] of Object.entries(entries)) {
    files[name] = typeof content === 'string' ? strToU8(content) : content
  }
  return zipSync(files, { level: 6 })
}

interface DocxOptions {
  readonly body: string
  readonly rels?: string
  readonly contentTypes?: string
  readonly styles?: string
  readonly settings?: string
  readonly numbering?: string
  readonly extra?: Readonly<Record<string, Entry>>
}

/** One document package carrying the standard auxiliaries plus caller entries. */
function docxPackage(options: DocxOptions): Uint8Array {
  const contentTypes = options.contentTypes ?? `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CT}">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '<Override PartName="/word/settings.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    + (options.numbering === undefined ? '' : '<Override PartName="/word/numbering.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>')
    + '</Types>'
  return packageOf({
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': ROOT_RELS,
    'word/document.xml': `${DOCUMENT_NS}<w:body>${options.body}</w:body></w:document>`,
    'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PKG_REL}">`
      + `<Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>`
      + (options.numbering === undefined ? '' : `<Relationship Id="rId9" Type="${REL}/numbering" Target="numbering.xml"/>`)
      + (options.rels ?? '')
      + '</Relationships>',
    'word/styles.xml': options.styles ?? STYLES,
    'word/settings.xml': options.settings ?? SETTINGS,
    ...(options.numbering === undefined ? {} : { 'word/numbering.xml': options.numbering }),
    ...options.extra,
  })
}

const SECTION = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'

function paragraph(inner: string, properties = ''): string {
  return `<w:p>${properties}${inner}</w:p>`
}

function run(inner: string, properties = ''): string {
  return `<w:r>${properties}${inner}</w:r>`
}

function text(value: string): string {
  return `<w:t xml:space="preserve">${value}</w:t>`
}

it('maps every paragraph formatting and run property branch', async () => {
  const body = `<w:p><w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="atLeast"/>
      <w:ind w:start="120" w:end="240" w:firstLine="360" w:hanging="480"/>
      <w:tabs><w:tab w:val="center" w:pos="720" w:leader="dot"/><w:tab w:val="clear" w:pos="1440"/>
        <w:tab w:pos="2880"/><w:tab w:val="unknown" w:pos="3600" w:leader="unknown"/><w:unknown/></w:tabs>
      <w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="FF0000"/>
        <w:bottom w:val="none"/><w:left w:val="dotted"/><w:right w:val="dashed" w:color="GGGGGG"/>
        <w:between w:val="double"/></w:pBdr>
      <w:shd w:fill="00FF00"/><w:outlineLvl w:val="2"/>
    </w:pPr>
    ${run(text('styled'), '<w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="SimSun"/><w:b/><w:i/><w:caps/>'
      + '<w:smallCaps/><w:dstrike/><w:color w:val="FF0000"/><w:sz w:val="24"/><w:highlight w:val="yellow"/>'
      + '<w:u w:val="wave" w:color="00FF00"/><w:vertAlign w:val="superscript"/>'
      + '<w:bdr w:val="dashed" w:color="0000FF"/></w:rPr>')}
    ${run(text('fallback'), '<w:rPr><w:rFonts w:eastAsia="SimSun" w:cs="Courier"/><w:strike/><w:sz w:val="0"/>'
      + '<w:highlight w:val="none"/><w:u w:val="none"/><w:vertAlign w:val="baseline"/>'
      + '<w:bdr w:val="none"/></w:rPr>')}
    ${run(text('other'), '<w:rPr><w:vertAlign w:val="subscript"/><w:bdr w:val="single" w:color="GGGGGG"/></w:rPr>')}
    ${run(text('plain'), '<w:rPr><w:rFonts/></w:rPr>')}
  </w:p>${SECTION}`
  const data = await readDocument(docxPackage({ body }))
  const style = data.body?.paragraphs?.[0]?.paragraphStyle
  expect(style?.horizontalAlign).toBe(HorizontalAlign.CENTER)
  expect(style?.spaceAbove?.v).toBe(8)
  expect(style?.spaceBelow?.v).toBe(16)
  expect(style?.lineSpacing).toBe(24)
  expect(style?.spacingRule).toBe(SpacingRule.AT_LEAST)
  expect(style?.indentStart?.v).toBe(8)
  expect(style?.indentEnd?.v).toBe(16)
  expect(style?.indentFirstLine?.v).toBe(24)
  expect(style?.hanging?.v).toBe(32)
  expect(style?.tabStops).toHaveLength(4)
  expect(style?.borderTop?.width).toBe(8 / 6)
  expect(style?.borderBottom).toBeUndefined()
  expect(style?.borderLeft?.dashStyle).toBe(DashStyleType.DOT)
  expect(style?.borderRight?.color.rgb).toBe('#000000')
  expect(style?.borderBetween?.dashStyle).toBe(DashStyleType.SOLID)
  expect(style?.shading?.backgroundColor.rgb).toBe('#00FF00')
  expect(style?.outlineLevel).toBe(2)
  const runs = data.body?.textRuns ?? []
  const styled = runs.find(candidate => data.body?.dataStream.slice(candidate.st, candidate.ed) === 'styled')
  expect(styled?.ts?.bl).toBe(1)
  expect(styled?.ts?.ff).toBe('Arial')
  expect(styled?.ts?.eastAsiaFontFamily).toBe('SimSun')
  expect(styled?.ts?.fs).toBe(12)
  expect(styled?.ts?.ul?.t).toBe(TextDecoration.WAVE)
  expect(styled?.ts?.va).toBe(BaselineOffset.SUPERSCRIPT)
  const fallback = runs.find(candidate => data.body?.dataStream.slice(candidate.st, candidate.ed) === 'fallback')
  expect(fallback?.ts?.ff).toBe('Courier')
  expect(fallback?.ts?.st?.t).toBe(12)
  expect(fallback?.ts?.fs).toBeUndefined()
  expect(fallback?.ts?.bg).toBeUndefined()
  expect(fallback?.ts?.ul).toBeUndefined()
  expect(fallback?.ts?.bd).toBeUndefined()
  const other = runs.find(candidate => data.body?.dataStream.slice(candidate.st, candidate.ed) === 'other')
  expect(other?.ts?.va).toBe(BaselineOffset.SUBSCRIPT)
  expect(other?.ts?.bd?.t?.cl.rgb).toBe('#000000')
  expect(runs.some(candidate => data.body?.dataStream.slice(candidate.st, candidate.ed) === 'plain')).toBe(false)
})

it('maps the remaining alignment, spacing, indentation, tab, and border variants', async () => {
  const variants = [
    '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:jc w:val="end"/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:t>c</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:jc w:val="justified"/><w:spacing w:line="360" w:lineRule="exact"/></w:pPr><w:r><w:t>d</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:jc w:val="distribute"/><w:spacing w:before="60"/></w:pPr><w:r><w:t>e</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:jc w:val="start"/><w:ind w:left="120"/></w:pPr><w:r><w:t>f</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:ind w:right="240"/></w:pPr><w:r><w:t>g</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:tabs><w:unknown/></w:tabs></w:pPr><w:r><w:t>h</w:t></w:r></w:p>',
    '<w:p><w:r><w:rPr><w:rFonts w:hAnsi="Arial"/><w:u w:val="unknown"/></w:rPr><w:t>i</w:t></w:r></w:p>',
    '<w:p><w:r><w:rPr><w:u/></w:rPr><w:t>j</w:t></w:r></w:p>',
    '<w:p><w:r><w:rPr><w:color w:val="GGGGGG"/><w:shd w:fill="FF0000"/></w:rPr><w:t>k</w:t></w:r></w:p>',
  ].join('')
  const data = await readDocument(docxPackage({ body: `${variants}${SECTION}` }))
  const paragraphs = data.body?.paragraphs ?? []
  expect(paragraphs).toHaveLength(11)
  expect(paragraphs[0]?.paragraphStyle?.horizontalAlign).toBe(HorizontalAlign.RIGHT)
  expect(paragraphs[1]?.paragraphStyle?.horizontalAlign).toBe(HorizontalAlign.RIGHT)
  expect(paragraphs[2]?.paragraphStyle).toMatchObject({
    horizontalAlign: HorizontalAlign.BOTH,
    lineSpacing: 1,
    spacingRule: SpacingRule.AUTO,
  })
  expect(paragraphs[3]?.paragraphStyle).toMatchObject({
    horizontalAlign: HorizontalAlign.JUSTIFIED,
    lineSpacing: 24,
    spacingRule: SpacingRule.EXACT,
  })
  expect(paragraphs[4]?.paragraphStyle).toMatchObject({ horizontalAlign: HorizontalAlign.DISTRIBUTED, spaceAbove: { v: 4 } })
  expect(paragraphs[5]?.paragraphStyle?.horizontalAlign).toBe(HorizontalAlign.LEFT)
  expect(paragraphs[5]?.paragraphStyle?.indentStart?.v).toBe(8)
  expect(paragraphs[6]?.paragraphStyle?.indentEnd?.v).toBe(16)
  expect(paragraphs[7]?.paragraphStyle).toBeUndefined()
  const runs = data.body?.textRuns ?? []
  const runAt = (value: string): ITextRun | undefined => (
    runs.find(candidate => data.body?.dataStream.slice(candidate.st, candidate.ed) === value)
  )
  expect(runAt('i')?.ts?.ff).toBe('Arial')
  expect(runAt('i')?.ts?.ul?.t).toBe(12)
  expect(runAt('j')?.ts?.ul?.t).toBe(12)
  expect(runAt('k')?.ts?.cl).toBeUndefined()
  expect(runAt('k')?.ts?.bg).toEqual({ rgb: '#FF0000' })
})

it('maps numbering levels, overrides, and paragraph bullets', async () => {
  const numbering = `<w:numbering xmlns:w="${WORD}">`
    + '<w:abstractNum w:abstractNumId="7">'
    + '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="0x2022"/><w:start w:val="3"/>'
    + '<w:isLgl/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/></w:rPr></w:lvl>'
    + '<w:lvl w:ilvl="2"><w:lvlText w:val=""/></w:lvl>'
    + '<w:lvl><w:numFmt w:val="customFormat"/></w:lvl>'
    + '</w:abstractNum>'
    + '<w:abstractNum><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>'
    + '<w:num w:numId="4"><w:abstractNumId w:val="7"/>'
    + '<w:lvlOverride w:ilvl="2"><w:startOverride w:val="9"/></w:lvlOverride>'
    + '<w:lvlOverride w:ilvl="3"/></w:num>'
    + '<w:num><w:abstractNumId w:val="0"/></w:num>'
    + '<w:num w:numId="6"><w:abstractNumId w:val="99"/></w:num>'
    + '</w:numbering>'
  const body = paragraph(run(text('one')), '<w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="4"/></w:numPr></w:pPr>')
    + paragraph(run(text('two')), '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="6"/></w:numPr></w:pPr>')
    + paragraph(run(text('three')), '<w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr>')
    + paragraph(run(text('four')), '<w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr>')
    + paragraph(run(text('five')), '<w:pPr><w:numPr/></w:pPr>')
    + SECTION
  const data = await readDocument(docxPackage({ body, numbering }))
  expect(data.lists?.['numbering-4']?.nestingLevel[0]).toMatchObject({
    glyphFormat: ' %1',
    glyphSymbol: '0x2022',
    glyphType: ListGlyphType.BULLET,
    isLegal: true,
    startNumber: 3,
    bulletAlignment: BulletAlignment.START,
  })
  expect(data.lists?.['numbering-4']?.nestingLevel[0]?.paragraphProperties?.horizontalAlign).toBe(HorizontalAlign.CENTER)
  expect(data.lists?.['numbering-4']?.nestingLevel[0]?.textStyle?.bl).toBe(1)
  expect(data.lists?.['numbering-4']?.nestingLevel[1]?.glyphFormat).toBe('%2.')
  expect(data.lists?.['numbering-4']?.nestingLevel[2]?.glyphType).toBe(ListGlyphType.CUSTOM)
  expect(data.lists?.['numbering-6']?.nestingLevel[0]?.glyphType).toBe(ListGlyphType.DECIMAL)
  const bullets = (data.body?.paragraphs ?? []).map(paragraph => paragraph.bullet)
  expect(bullets[0]).toEqual({ listType: 'numbering-4', listId: 'num-4', nestingLevel: 2, startNumber: 9 })
  expect(bullets[1]?.listType).toBe('numbering-6')
  expect(bullets[2]?.listType).toBe('numbering-5')
  expect(bullets[3]).toBeUndefined()
  expect(bullets[4]).toBeUndefined()
})

it('maps table widths, spans, merges, borders, margins, and row properties', async () => {
  const table = '<w:tbl>'
    + '<w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/><w:tblInd w:w="120"/>'
    + '<w:tblLayout w:type="fixed"/></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="auto"/><w:unknown/></w:tblGrid>'
    + '<w:tr><w:trPr><w:trHeight w:val="300" w:hRule="exact"/><w:cantSplit/><w:tblHeader/></w:trPr>'
    + '<w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>'
    + '<w:tcBorders><w:top w:val="single"/><w:insideH w:val="dotted"/><w:start w:val="dashed"/>'
    + '<w:right w:val="none"/><w:insideV w:val="double"/></w:tcBorders>'
    + '<w:shd w:fill="ABCDEF"/><w:tcMar><w:top w:w="120"/><w:start w:w="240"/></w:tcMar>'
    + '<w:vAlign w:val="center"/></w:tcPr>'
    + `${runInParagraph('a')}</w:tc>`
    + '<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p/></w:tc>'
    + '</w:tr>'
    + '<w:tr><w:trPr><w:trHeight w:val="240" w:hRule="atLeast"/></w:trPr>'
    + '<w:tc><w:tcPr><w:vMerge/><w:vAlign w:val="bottom"/></w:tcPr><w:p/></w:tc>'
    + '<w:tc><w:tcPr><w:vAlign w:val="top"/></w:tcPr><w:p/></w:tc>'
    + '</w:tr>'
    + '<w:tr><w:tc><w:p/></w:tc></w:tr>'
    + '</w:tbl>'
  const other = '<w:tbl><w:tblPr><w:tblW w:w="100" w:type="auto"/><w:jc w:val="end"/>'
    + '<w:tblLayout w:type="autofit"/></w:tblPr>'
    + '<w:tr><w:trPr><w:trHeight w:val="300"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="abc"/></w:tcPr><w:p/></w:tc></w:tr>'
    + '</w:tbl>'
  const third = '<w:tbl><w:tblPr><w:jc w:val="right"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
  const fourth = '<w:tbl><w:tblPr><w:jc w:val="left"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
  const data = await readDocument(docxPackage({ body: `${table}${other}${third}${fourth}${SECTION}` }))
  const model = data.tableSource?.['table-1']
  expect(model?.size).toEqual({ type: 1, width: { v: 100, u: NumberUnitType.PERCENT } })
  expect(model?.align).toBe(TableAlignmentType.CENTER)
  expect(model?.indent?.v).toBe(8)
  expect(model?.layout).toBe(TableLayoutType.FIXED)
  expect(model?.tableColumns.map(column => column.size.width.v)).toEqual([80, 160, 160])
  const firstRow = model?.tableRows[0]
  expect(firstRow?.trHeight).toEqual({ val: { v: 20 }, hRule: 2 })
  expect(firstRow?.cantSplit).toBe(1)
  expect(firstRow?.repeatHeaderRow).toBe(1)
  expect(firstRow?.tableCells[0]?.columnSpan).toBe(2)
  expect(firstRow?.tableCells[0]?.rowSpan).toBe(2)
  expect(firstRow?.tableCells[0]?.borderTop?.width).toEqual({ v: 4 / 6 })
  expect(firstRow?.tableCells[0]?.borderBottom?.dashStyle).toBe(DashStyleType.DOT)
  expect(firstRow?.tableCells[0]?.borderLeft?.dashStyle).toBe(DashStyleType.DASH)
  expect(firstRow?.tableCells[0]?.borderRight).toBeUndefined()
  expect(firstRow?.tableCells[0]?.backgroundColor?.rgb).toBe('#ABCDEF')
  expect(firstRow?.tableCells[0]?.margin?.top.v).toBe(8)
  expect(firstRow?.tableCells[0]?.margin?.start.v).toBe(16)
  expect(firstRow?.tableCells[0]?.margin?.end.v).toBe(10)
  expect(firstRow?.tableCells[0]?.vAlign).toBe(VerticalAlignmentType.CENTER)
  expect(model?.tableRows[1]?.tableCells[0]?.vAlign).toBe(VerticalAlignmentType.BOTTOM)
  expect(model?.tableRows[1]?.tableCells[1]?.vAlign).toBe(VerticalAlignmentType.TOP)
  expect(data.tableSource?.['table-2']?.layout).toBe(0)
  expect(data.tableSource?.['table-2']?.align).toBe(TableAlignmentType.END)
  expect(data.tableSource?.['table-3']?.align).toBe(TableAlignmentType.END)
  expect(data.tableSource?.['table-4']?.align).toBe(TableAlignmentType.START)
})

function runInParagraph(value: string): string {
  return `<w:p>${run(text(value))}</w:p>`
}

it('maps a vertical merge chain with a restart spanning two rows', async () => {
  const table = '<w:tbl><w:tblGrid><w:gridCol w:w="1200"/></w:tblGrid>'
    + '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p/></w:tc></w:tr>'
    + '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>'
    + '<w:tr><w:tc><w:p/></w:tc></w:tr>'
    + '</w:tbl>'
  const data = await readDocument(docxPackage({ body: `${table}${SECTION}` }))
  const rows = data.tableSource?.['table-1']?.tableRows ?? []
  expect(rows[0]?.tableCells[0]?.rowSpan).toBe(2)
  expect(rows[1]?.tableCells[0]?.rowSpan).toBe(0)
  expect(rows[2]?.tableCells[0]?.rowSpan).toBeUndefined()
})

it('maps section sizes, margins, columns, references, and titles', async () => {
  const body = '<w:p><w:pPr><w:sectPr>'
    + '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
    + '<w:pgMar w:bottom="30" w:right="60" w:header="90" w:footer="120"/>'
    + '<w:cols w:sep="1"><w:col w:w="3000" w:space="400"/><w:col w:w="3000"/><w:unknown/></w:cols>'
    + '<w:type w:val="continuous"/>'
    + '<w:headerReference w:type="first" r:id="rId2"/>'
    + '<w:headerReference w:type="default" r:id="rId3"/>'
    + '<w:headerReference w:type="default" r:id="rId4"/>'
    + '<w:footerReference w:type="even" r:id="rId5"/>'
    + '<w:footerReference w:type="default" r:id="rId6"/>'
    + '<w:titlePg/>'
    + '</w:sectPr></w:pPr><w:r><w:t>one</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:sectPr><w:pgSz w:w="15840" w:h="12240"/>'
    + '<w:cols w:num="2" w:space="20000"/></w:sectPr></w:pPr></w:p>'
    + '<w:p><w:pPr><w:sectPr><w:pgSz w:h="15840"/>'
    + '<w:cols w:num="3" w:space="720" w:sep="0"/></w:sectPr></w:pPr></w:p>'
    + '<w:p><w:pPr><w:sectPr><w:cols w:num="1"/></w:sectPr></w:pPr></w:p>'
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>'
    + '<w:cols w:sep="true"/><w:type w:val="nextColumn"/></w:sectPr>'
  const rels = `<Relationship Id="rId2" Type="${REL}/header" Target="header1.xml"/>`
    + `<Relationship Id="rId3" Type="${REL}/header" Target="header2.xml"/>`
    + `<Relationship Id="rId4" Type="${REL}/header" Target="headerMissing.xml"/>`
    + `<Relationship Id="rId5" Type="${REL}/footer" Target="footer1.xml"/>`
    + `<Relationship Id="rId6" Type="${REL}/footer" Target="footer1.xml"/>`
    + `<Relationship Id="rId7" Type="${REL}/hyperlink" Target=""/>`
  const extra = {
    'word/header1.xml': `<w:hdr xmlns:w="${WORD}"><w:p><w:r><w:t>first</w:t></w:r></w:p></w:hdr>`,
    'word/header2.xml': `<w:hdr xmlns:w="${WORD}"><w:p><w:r><w:t>default</w:t></w:r></w:p></w:hdr>`,
    'word/footer1.xml': `<w:ftr xmlns:w="${WORD}"><w:p><w:r><w:t>footer</w:t></w:r></w:p></w:ftr>`,
  }
  const data = await readDocument(docxPackage({ body, rels, extra }))
  const sections = data.body?.sectionBreaks ?? []
  expect(sections[0]?.pageSize).toEqual({ width: 16838 / 15, height: 11906 / 15 })
  expect(sections[0]?.pageOrient).toBe(PageOrientType.LANDSCAPE)
  expect(sections[0]?.marginBottom).toBe(2)
  expect(sections[0]?.marginRight).toBe(4)
  expect(sections[0]?.marginHeader).toBe(6)
  expect(sections[0]?.marginFooter).toBe(8)
  expect(sections[0]?.marginTop).toBeUndefined()
  expect(sections[0]?.columnSeparatorType).toBe(ColumnSeparatorType.BETWEEN_EACH_COLUMN)
  expect(sections[0]?.columnProperties).toEqual([
    { width: 200, paddingEnd: 400 / 15 },
    { width: 200, paddingEnd: 0 },
  ])
  expect(sections[0]?.firstPageHeaderId).toBe('header1')
  expect(sections[0]?.defaultHeaderId).toBe('header2')
  expect(sections[0]?.evenPageFooterId).toBe('footer1')
  expect(sections[0]?.defaultFooterId).toBe('footer1')
  expect(sections[0]?.useFirstPageHeaderFooter).toBe(1)
  expect(sections[0]?.sectionType).toBe(SectionType.CONTINUOUS)
  expect(sections[1]?.columnProperties).toBeUndefined()
  expect(sections[2]?.columnProperties).toBeUndefined()
  expect(sections[2]?.columnSeparatorType).toBe(ColumnSeparatorType.NONE)
  expect(sections[3]?.columnProperties).toBeUndefined()
  expect(sections[4]?.pageOrient).toBe(PageOrientType.PORTRAIT)
  expect(sections[4]?.columnSeparatorType).toBe(ColumnSeparatorType.BETWEEN_EACH_COLUMN)
  expect(sections[4]?.columnProperties).toBeUndefined()
  expect(sections[4]?.sectionType).toBe(SectionType.NEXT_COLUMN)
})

it('maps hyperlink targets, anchors, and plain hyperlinks', async () => {
  const body = paragraph(
    '<w:hyperlink r:id="rId7">' + run(text('external')) + '</w:hyperlink>'
    + '<w:hyperlink w:anchor="mark">' + run(text('anchor')) + '</w:hyperlink>'
    + '<w:hyperlink r:id="rId8">' + run(text('missing')) + '</w:hyperlink>'
    + '<w:hyperlink r:id="rId9">' + run(text('plain')) + '</w:hyperlink>'
    + '<w:hyperlink>' + run(text('bare')) + '<w:proofErr/></w:hyperlink>'
    + '<w:hyperlink r:id="rId10">' + run(text('empty')) + '</w:hyperlink>'
    + '<w:hyperlink w:anchor="mark2">' + run(text('two')) + '<w:proofErr/>' + '</w:hyperlink>',
  ) + SECTION
  const rels = `<Relationship Id="rId7" Type="${REL}/hyperlink" Target="https://example.test/a"/>`
    + `<Relationship Id="rId9" Type="${REL}/styles" Target="styles.xml"/>`
    + `<Relationship Id="rId10" Type="${REL}/hyperlink" Target=""/>`
  const data = await readDocument(docxPackage({ body, rels }))
  const ranges = data.body?.customRanges ?? []
  expect(ranges).toHaveLength(3)
  expect(ranges[0]?.properties?.['url']).toBe('https://example.test/a')
  expect(ranges[1]?.properties?.['bookmarkId']).toBe('mark')
  expect(ranges[2]?.properties).toEqual({ bookmarkId: 'mark2' })
})

it('carries paragraph-level preserved elements and raw runs', async () => {
  const body = paragraph(
    '<w:r><w:drawing><wp:inline/></w:drawing></w:r>'
    + '<w:r><w:object/></w:r>'
    + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:instrText>PAGE</w:instrText></w:r>'
    + '<w:r><w:footnoteReference w:id="1"/></w:r>'
    + '<w:r><w:endnoteReference w:id="1"/></w:r>'
    + '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing/></mc:Choice></mc:AlternateContent></w:r>'
    + '<w:r><w:r/><w:drawing/></w:r>'
    + '<w:r/>'
    + '<w:pict/>'
    + '<w:sdt><w:sdt/></w:sdt>'
    + '<w:fldSimple w:instr="DATE"/>'
    + '<w:bookmarkStart w:id="0" w:name="mark"/>'
    + '<w:bookmarkEnd w:id="0"/>',
  ) + SECTION
  const data = await readDocument(docxPackage({ body }))
  const blocks = data.body?.docxRawBlocks ?? []
  expect(blocks.length).toBeGreaterThanOrEqual(12)
  expect(blocks.some(block => block.xml.includes('<w:drawing>'))).toBe(true)
  expect(blocks.some(block => block.xml.includes('<w:sdt>'))).toBe(true)
  expect(blocks.some(block => block.xml.includes('<w:fldSimple'))).toBe(true)
  expect(blocks.some(block => block.xml.includes('<w:bookmarkStart'))).toBe(true)
})

it('refuses a preserved run that has no source record when its prefix is not w', async () => {
  const body = '<w:p><x:r><w:drawing/></x:r></w:p>'
  await expect(readDocument(docxPackage({ body }))).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses a preserved paragraph element that has no source record when its prefix is not w', async () => {
  const body = '<w:p><x:sdt/></w:p>'
  await expect(readDocument(docxPackage({ body }))).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('maps a raw record fallback and then refuses the unmatched duplicate record', async () => {
  const body = '<w:p><x:r><w:drawing/></x:r>'
    + '<w:r><w:drawing r:embed="rId5"/></w:r></w:p>'
  await expect(readDocument(docxPackage({ body }))).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('reads header and footer story parts including tables and content controls', async () => {
  const body = paragraph(run(text('body')))
    + '<w:sectPr><w:headerReference w:type="default" r:id="rId2"/>'
    + '<w:footerReference w:type="default" r:id="rId3"/></w:sectPr>'
  const rels = `<Relationship Id="rId2" Type="${REL}/header" Target="header1.xml"/>`
    + `<Relationship Id="rId3" Type="${REL}/footer" Target="footer1.xml"/>`
  const extra = {
    'word/header1.xml': `<w:hdr xmlns:w="${WORD}"><w:p><w:r><w:t>header</w:t></w:r></w:p>`
      + '<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:sdt/></w:hdr>',
    'word/footer1.xml': `<w:ftr xmlns:w="${WORD}"><w:p><w:r><w:t>footer</w:t></w:r></w:p></w:ftr>`,
  }
  const data = await readDocument(docxPackage({ body, rels, extra }))
  expect(data.headers?.['header1']?.body.tables).toHaveLength(1)
  expect(data.headers?.['header1']?.body.paragraphs).toHaveLength(3)
  expect(data.footers?.['footer1']?.body.dataStream).toBe('footer\r\n')
})

it('reads forced-linear numbering ids, bare attributes, and invalid ids', async () => {
  const numbering = `<w:numbering xmlns:w="${WORD}">`
    + '<w:abstractNum abstractNumId="5"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>'
    + '<w:num w:numId="bad"><w:abstractNumId w:val="5"/></w:num>'
    + '<w:num w:numId="8" numId="9"><w:abstractNumId w:val="5"/></w:num>'
    + '</w:numbering>'
  const body = paragraph(run(text('item')), '<w:pPr><w:numPr><w:numId w:val="8"/></w:numPr></w:pPr>') + SECTION
  const data = await readDocument(docxPackage({ body, numbering }))
  expect(data.lists?.['numbering-8']?.nestingLevel).toHaveLength(1)
})

it('maps package entries: preserved media, notes, declarations, and incomplete relationships', async () => {
  const body = paragraph(run(text('image'))) + SECTION
  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CT}">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Default ContentType="application/octet-stream"/>'
    + '<Default Extension="jpeg" ContentType="image/jpeg"/>'
    + '<Override ContentType="text/plain"/>'
    + '<Override PartName="/word/document.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '<Override PartName="/word/settings.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    + '<Override PartName="/word/footnotes.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
    + '<Override PartName="/word/endnotes.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
    + '<Override PartName="/customXml/itemProps1.xml" ContentType="application/xml"/>'
    + '</Types>'
  const rels = `<Relationship Id="rId5" Type="${REL}/image" Target="media/noext"/>`
    + '<Relationship Id="rId6" Target="media/image2.png"/>'
    + `<Relationship Id="rId7" Type="${REL}/image" Target="media/image3.png"/>`
    + '<Relationship Type="x" Target="y"/>'
  const extra = {
    'word/media/noext': new Uint8Array([1]),
    'word/media/image2.png': new Uint8Array([1, 2]),
    'word/media/image3.png': new Uint8Array([1, 2, 3, 4]),
    'word/footnotes.xml': `<w:footnotes xmlns:w="${WORD}"><w:footnote w:id="1"><w:p/></w:footnote></w:footnotes>`,
    'word/endnotes.xml': `<w:endnotes xmlns:w="${WORD}"><w:endnote w:id="1"><w:p/></w:endnote></w:endnotes>`,
    'word/_rels/footnotes.xml.rels': '<Relationships><Relationship Id="rId1" Target="media/image3.png"/></Relationships>',
    'word/header1.xml': `<w:hdr xmlns:w="${WORD}"><w:p/></w:hdr>`,
    'word/_rels/header1.xml.rels': '<Relationships><Relationship Id="rId1" Target="media/image2.png"/></Relationships>',
    'customXml/itemProps1.xml': '<props/>',
    'word/': new Uint8Array(0),
  }
  const data = await readDocument(docxPackage({ body, rels, contentTypes, extra }))
  const preserved = extensionView(data)?.docx
  expect(preserved?.parts.map(part => part.name)).toEqual([
    'word/_rels/footnotes.xml.rels',
    'word/endnotes.xml',
    'word/footnotes.xml',
    'word/media/image2.png',
    'word/media/image3.png',
    'word/media/noext',
  ])
  expect(preserved?.contentTypes.length).toBeGreaterThan(0)
  expect(preserved?.relationships.document?.length).toBeGreaterThan(0)
  expect(preserved?.relationships.stories).toBeUndefined()
})

it('maps a styles part with default and named styles, and rejects malformed styles', async () => {
  const styled = `<w:styles xmlns:w="${WORD}">`
    + '<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="240"/></w:pPr></w:pPrDefault>'
    + '<w:rPrDefault><w:rPr><w:i/></w:rPr></w:rPrDefault></w:docDefaults>'
    + '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/>'
    + '<w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Other"/>'
    + '</w:styles>'
  const data = await readDocument(docxPackage({
    body: paragraph(run(text('x'))) + SECTION,
    styles: styled,
  }))
  expect(data.documentStyle.textStyle?.bl).toBe(1)
  expect(data.documentStyle.textStyle?.it).toBe(1)
  expect(data.documentStyle.defaultParagraphStyle?.horizontalAlign).toBe(HorizontalAlign.CENTER)
  expect(data.documentStyle.defaultParagraphStyle?.spaceBelow?.v).toBe(16)

  const missing = `<w:styles xmlns:w="${WORD}"><w:docDefaults/></w:styles>`
  const plain = await readDocument(docxPackage({ body: `${paragraph(run(text('x')))}${SECTION}`, styles: missing }))
  expect(plain.documentStyle.textStyle).toBeUndefined()

  await expect(readDocument(docxPackage({
    body: paragraph(run(text('x'))) + SECTION,
    styles: '<w:styles><w:style></w:styles>',
  }))).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('reads numeric attributes with non-finite values as absent', async () => {
  const body = '<w:tbl><w:tblPr><w:tblW w:w="abc" w:type="dxa"/></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="abc"/></w:tblGrid>'
    + '<w:tr><w:trPr><w:trHeight w:val="abc"/></w:trPr>'
    + '<w:tc><w:tcPr><w:gridSpan w:val="abc"/><w:tcW w:w="abc"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>'
    + '<w:p><w:pPr><w:ind w:start="abc"/><w:spacing w:before="abc"/></w:pPr></w:p>'
    + '<w:p><w:pPr><w:sectPr><w:pgSz w:w="abc" w:h="abc"/><w:pgMar w:top="abc"/>'
    + '<w:cols w:num="abc"/></w:sectPr></w:pPr></w:p>'
    + SECTION
  const data = await readDocument(docxPackage({ body }))
  expect(data.tableSource?.['table-1']?.size?.width.v).toBe(160)
  expect(data.tableSource?.['table-1']?.tableRows[0]?.trHeight).toEqual({ val: { v: 30 }, hRule: 0 })
  expect(data.tableSource?.['table-1']?.tableRows[0]?.tableCells[0]?.columnSpan).toBeUndefined()
  expect(data.body?.sectionBreaks?.[0]?.pageSize).toBeUndefined()
  expect(data.body?.sectionBreaks?.[0]?.marginTop).toBeUndefined()
})

it('reads an empty content-types part and absent document relationships', async () => {
  const data = await readDocument(packageOf({
    '[Content_Types].xml': '<Types/>',
    'word/document.xml': `${DOCUMENT_NS}<w:body>${paragraph(run(text('body')))}${SECTION}</w:body></w:document>`,
  }))
  expect(data.body?.dataStream).toContain('body')
  expect(extensionView(data)).toBeUndefined()
})

it('leaves the source sidecar absent when anchors exceed the configured budget', async () => {
  const source = docxPackage({ body: `${paragraph(run(text('one')))}${paragraph(run(text('two')))}${SECTION}` })
  const data = await readDocumentWithSource(source, 'docx', {
    ...DEFAULT_OFFICE_EXCHANGE_CONFIG,
    maxSourceAnchors: 1,
    maxSourceAnchorBytes: 1024,
  })
  expect(extensionView(data)?.source).toBeUndefined()
})

it('anchors a nested table and its cells', async () => {
  const nested = '<w:tbl><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="1200"/></w:tblGrid>'
    + '<w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc>'
    + '<w:tc><w:p><w:r><w:t>two</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const outer = '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>'
    + '<w:tr><w:tc><w:p><w:r><w:t>outer</w:t></w:r></w:p>'
    + nested
    + '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>sibling</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    + '</w:tc></w:tr></w:tbl>'
  const data = await readDocumentWithSource(docxPackage({ body: `${outer}${SECTION}` }), 'docx')
  const source = extensionView(data)?.source
  expect(source?.anchorMode).toBe('per-unit')
  if (source?.anchorMode !== 'per-unit') throw new Error('fixture has no per-unit anchors')
  expect(source.anchors.filter(anchor => anchor.kind === 'tableCell').length).toBe(1)
})

it('leaves the source sidecar absent when a scan or model count disagrees', async () => {
  const extraParagraph = docxPackage({
    body: `${paragraph(run(text('one')))}<w:hyperlink><w:p/></w:hyperlink>${SECTION}`,
  })
  const noAnchors = await readDocumentWithSource(extraParagraph, 'docx')
  expect(extensionView(noAnchors)?.source).toBeUndefined()

  const alternate = docxPackage({
    body: `<w:p><mc:AlternateContent><mc:Choice Requires="wps">${run(text('x'))}</mc:Choice>`
      + '</mc:AlternateContent></w:p>' + SECTION,
  })
  const refused = await readDocumentWithSource(alternate, 'docx')
  expect(extensionView(refused)?.source).toBeUndefined()
})

it('leaves the source sidecar absent when the opened package has duplicate folded names', async () => {
  const source = docxPackage({ body: `${paragraph(run(text('one')))}${SECTION}` })
  const duplicate = packageOf({
    '[Content_Types].xml': '<Types/>',
    'word/document.xml': `${DOCUMENT_NS}<w:body>${paragraph(run(text('one')))}${SECTION}</w:body></w:document>`,
    'WORD/DOCUMENT.XML': '<w:document><w:body/></w:document>',
  })
  const data = await readDocumentWithSource(duplicate, 'docx')
  expect(extensionView(data)?.source).toBeUndefined()
  expect(source.byteLength).toBeGreaterThan(0)
})








it('maps bullet levels with no text, partial borders, and unknown highlight names', async () => {
  const numbering = `<w:numbering xmlns:w="${WORD}">`
    + '<w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/>'
    + '</w:lvl></w:abstractNum>'
    + '<w:num w:numId="3"><w:abstractNumId w:val="3"/></w:num>'
    + '</w:numbering>'
  const body = paragraph(
    run(text('plain'), '<w:rPr><w:highlight w:val="chartreuse"/></w:rPr>'),
    '<w:pPr><w:pBdr><w:bottom w:val="unknown"/></w:pBdr></w:pPr>',
  )
    + paragraph(run(text('item')), '<w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr>')
    + SECTION
  const data = await readDocument(docxPackage({ body, numbering }))
  expect(data.lists?.['numbering-3']?.nestingLevel[0]?.glyphType).toBe(ListGlyphType.BULLET)
  expect(data.lists?.['numbering-3']?.nestingLevel[0]?.glyphSymbol).toBeUndefined()
  const paragraphStyle = data.body?.paragraphs?.[0]?.paragraphStyle
  expect(paragraphStyle?.borderTop).toBeUndefined()
  expect(paragraphStyle?.borderBottom?.dashStyle).toBe(DashStyleType.SOLID)
  const plain = data.body?.textRuns?.[0]
  expect(plain?.ts?.bg).toBeUndefined()
})

it('maps partial cell borders, margins, and table widths without a type', async () => {
  const table = '<w:tbl><w:tblPr><w:tblW w:w="1000"/></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="1200"/></w:tblGrid>'
    + '<w:tr><w:tc><w:tcPr><w:tcBorders><w:insideH w:val="dotted"/></w:tcBorders>'
    + '<w:tcMar><w:end w:w="240"/></w:tcMar></w:tcPr><w:p/></w:tc></w:tr>'
    + '<w:tr><w:tc><w:tcPr><w:tcBorders><w:insideV w:val="unknown"/></w:tcBorders></w:tcPr><w:p/></w:tc></w:tr>'
    + '</w:tbl>'
  const data = await readDocument(docxPackage({ body: `${table}${SECTION}` }))
  const model = data.tableSource?.['table-1']
  expect(model?.size?.width.v).toBe(1000 / 15)
  const first = model?.tableRows[0]?.tableCells[0]
  expect(first?.borderTop?.dashStyle).toBe(DashStyleType.DOT)
  expect(first?.borderBottom?.dashStyle).toBe(DashStyleType.DOT)
  expect(first?.margin?.start.v).toBe(10)
  expect(first?.margin?.top.v).toBe(5)
  expect(first?.margin?.end.v).toBe(16)
  const second = model?.tableRows[1]?.tableCells[0]
  expect(second?.borderRight?.dashStyle).toBe(DashStyleType.SOLID)
})

it('ignores header and footer references whose relationship type disagrees', async () => {
  const body = paragraph(run(text('body')))
    + '<w:sectPr><w:headerReference w:type="even" r:id="rId5"/>'
    + '<w:footerReference w:type="even" r:id="rId2"/></w:sectPr>'
  const rels = `<Relationship Id="rId2" Type="${REL}/header" Target="header1.xml"/>`
    + `<Relationship Id="rId5" Type="${REL}/footer" Target="footer1.xml"/>`
  const extra = {
    'word/header1.xml': `<w:hdr xmlns:w="${WORD}"><w:p/></w:hdr>`,
    'word/footer1.xml': `<w:ftr xmlns:w="${WORD}"><w:p/></w:ftr>`,
  }
  const data = await readDocument(docxPackage({ body, rels, extra }))
  const section = data.body?.sectionBreaks?.[0]
  expect(section?.evenPageHeaderId).toBeUndefined()
  expect(section?.evenPageFooterId).toBeUndefined()
})
