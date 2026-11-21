/** Unsupported-feature detection over in-memory OOXML packages. */
import { readFile } from 'node:fs/promises'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { PageOrientType } from '@univerjs/core'
import { expect, it } from 'vitest'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from '../src/config.ts'
import { detectPreserved, detectUnsupportedFeatures } from '../src/detect.ts'
import type { OfficeExchangeErrorCode } from '../src/errors.ts'
import { OfficeExchangeError } from '../src/errors.ts'
import { readDocument, writeDocument } from '../src/provider/docx.ts'
import { readWorkbook, writeWorkbook } from '../src/provider/xlsx.ts'
import type { CoupledConstructMarker, PreservedConstructKind } from '../src/types.ts'

const DOCUMENT_XML = 'word/document.xml'
const PLAIN_DOCUMENT = '<w:document><w:body><w:p/></w:body></w:document>'
const WORKSHEET_XML = 'xl/worksheets/sheet1.xml'
const PLAIN_WORKSHEET = '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'
const DOCX_FIXTURE = new URL('./fixtures/roundtrip.docx', import.meta.url)
const XLSX_FIXTURE = new URL('./fixtures/roundtrip.xlsx', import.meta.url)
const MACRO_FIXTURE = new URL('./fixtures/unsupported-macro.docx', import.meta.url)
const TRACKED_FIXTURE = new URL('./fixtures/unsupported-tracked.docx', import.meta.url)

function preservedKindsOf(bytes: Uint8Array): readonly PreservedConstructKind[] {
  return [...detectPreserved(bytes).preserved]
}

function markersOf(bytes: Uint8Array): readonly CoupledConstructMarker[] {
  return detectPreserved(bytes).coupled
}

function bodyOf(inner: string): string {
  return `<w:document><w:body>${inner}</w:body></w:document>`
}

function packageOf(entries: Readonly<Record<string, string | Uint8Array>>): Uint8Array {
  const files: Record<string, Uint8Array> = {}
  for (const [name, content] of Object.entries(entries)) {
    files[name] = typeof content === 'string' ? strToU8(content) : content
  }
  return zipSync(files)
}

function codeOfThrown(run: () => void): OfficeExchangeErrorCode {
  try {
    run()
  } catch (error) {
    if (!(error instanceof OfficeExchangeError)) throw error
    return error.code
  }
  throw new Error('expected the call to throw')
}

it('reports no features for a clean document', () => {
  const bytes = packageOf({
    '[Content_Types].xml': '<Types/>',
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/styles.xml': '<styles/>',
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('detects macro parts', () => {
  for (const path of ['word/vbaProject.bin', 'xl/vbaProject.bin', 'ppt/vbaProject.bin']) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: new Uint8Array([0]) })
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'macro', detail: path }])
  }
})

it('detects tracked changes in the document body', () => {
  const ins = packageOf({ [DOCUMENT_XML]: '<w:document><w:ins w:id="1"/></w:document>' })
  expect(detectUnsupportedFeatures(ins)).toEqual([{ kind: 'tracked-changes', detail: DOCUMENT_XML }])

  const del = packageOf({ [DOCUMENT_XML]: '<w:document><w:del w:id="2"/></w:document>' })
  expect(detectUnsupportedFeatures(del)).toEqual([{ kind: 'tracked-changes', detail: DOCUMENT_XML }])
})

it('detects comment parts across Word, Excel, and PowerPoint', () => {
  for (const path of ['word/comments.xml', 'xl/comments1.xml', 'ppt/comments/comment1.xml']) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: '<comments/>' })
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'comment', detail: path }])
  }
})

it('detects chart parts', () => {
  for (const path of ['xl/charts/chart1.xml', 'ppt/charts/chart1.xml']) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: '<chart/>' })
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'chart', detail: path }])
  }
})

it('detects drawing and VML shape parts', () => {
  const drawing = packageOf({
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/drawings/drawing1.xml': '<drawing/>',
  })
  expect(detectUnsupportedFeatures(drawing)).toEqual([
    { kind: 'drawing', detail: 'word/drawings/drawing1.xml' },
  ])

  const shape = packageOf({
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/vmlDrawing1.vml': '<vml/>',
  })
  expect(detectUnsupportedFeatures(shape)).toEqual([
    { kind: 'shape', detail: 'word/vmlDrawing1.vml' },
  ])
})

it('does not report header and footer parts', () => {
  for (const path of ['word/header1.xml', 'word/footer2.xml']) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: '<hdr/>' })
    expect(detectUnsupportedFeatures(bytes)).toEqual([])
  }
})

it('does not report mapped list and hyperlink constructs inside header and footer parts', () => {
  const parts = packageOf({
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/header1.xml': '<w:hdr>'
      + '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
      + '<w:hyperlink w:anchor="top"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p></w:hdr>',
    'word/numbering.xml': '<w:numbering><w:abstractNum w:abstractNumId="0"/><w:num w:numId="1"/></w:numbering>',
  })
  expect(detectUnsupportedFeatures(parts)).toEqual([])
})

it('does not report image relationships inside header and footer parts', () => {
  const external = packageOf({
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/footer1.xml': '<w:ftr/>',
    'word/_rels/footer1.xml.rels': '<Relationships><Relationship Id="rId1"'
      + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
      + ' Target="https://example.test/a.png" TargetMode="External"/></Relationships>',
  })
  expect(detectUnsupportedFeatures(external)).toEqual([])
})

it('detects embedded objects', () => {
  for (const path of ['word/embeddings/oleObject1.bin', 'xl/embeddings/oleObject1.bin']) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: new Uint8Array([0]) })
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'embedded-object', detail: path }])
  }
})

it('reports every kind once, ordered by the fixed kind order', () => {
  const bytes = packageOf({
    'word/embeddings/oleObject1.bin': new Uint8Array([0]),
    'word/vmlDrawing1.vml': '<vml/>',
    'ppt/charts/chart1.xml': '<chart/>',
    'word/comments.xml': '<comments/>',
    [DOCUMENT_XML]: '<w:document><w:ins w:id="1"/></w:document>',
    'xl/vbaProject.bin': new Uint8Array([0]),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([
    { kind: 'macro', detail: 'xl/vbaProject.bin' },
    { kind: 'tracked-changes', detail: DOCUMENT_XML },
    { kind: 'comment', detail: 'word/comments.xml' },
    { kind: 'chart', detail: 'ppt/charts/chart1.xml' },
    { kind: 'shape', detail: 'word/vmlDrawing1.vml' },
    { kind: 'embedded-object', detail: 'word/embeddings/oleObject1.bin' },
  ])
})

it('keeps only the first part per kind', () => {
  const bytes = packageOf({
    'xl/charts/chart1.xml': '<chart/>',
    'xl/charts/chart2.xml': '<chart/>',
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([
    { kind: 'chart', detail: 'xl/charts/chart1.xml' },
  ])
})

it('rejects input over maxFileBytes before parsing', () => {
  const bytes = new Uint8Array(32)
  expect(codeOfThrown(() => detectUnsupportedFeatures(bytes, {
    ...DEFAULT_OFFICE_EXCHANGE_CONFIG,
    maxFileBytes: 8,
  }))).toBe('too-large')
})

it('rejects a valid package larger than maxFileBytes', () => {
  const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT })
  expect(codeOfThrown(() => detectUnsupportedFeatures(bytes, {
    ...DEFAULT_OFFICE_EXCHANGE_CONFIG,
    maxFileBytes: bytes.byteLength - 1,
  }))).toBe('too-large')
})

it('rejects packages with too many entries', () => {
  const entries: Record<string, Uint8Array> = {}
  for (let index = 0; index <= 10_000; index += 1) {
    entries[`word/media/part${index}.xml`] = new Uint8Array(0)
  }
  expect(codeOfThrown(() => detectUnsupportedFeatures(zipSync(entries)))).toBe('too-large')
})

it('classifies non-zip input as conversion-failed', () => {
  expect(codeOfThrown(() => detectUnsupportedFeatures(strToU8('not a zip archive')))).toBe('conversion-failed')
  expect(codeOfThrown(() => detectUnsupportedFeatures(new Uint8Array(0)))).toBe('conversion-failed')
})

it('does not report worksheet media or preserved document media', () => {
  const worksheet = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, 'xl/media/image1.png': new Uint8Array([0]) })
  expect(detectUnsupportedFeatures(worksheet)).toEqual([])

  const document = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, 'word/media/image1.png': new Uint8Array([0]) })
  expect(detectUnsupportedFeatures(document)).toEqual([])
})

it('reports a workbook drawing that is not a cell picture', () => {
  const shape = packageOf({
    'xl/drawings/drawing1.xml': '<xdr:wsDr><xdr:sp/></xdr:wsDr>',
  })
  expect(detectUnsupportedFeatures(shape)).toEqual([
    { kind: 'drawing', detail: 'xl/drawings/drawing1.xml' },
  ])

  for (const element of ['xdr:cxnSp', 'xdr:grpSp', 'xdr:graphicFrame', 'xdr:absoluteAnchor']) {
    const bytes = packageOf({
      'xl/drawings/drawing1.xml': `<xdr:wsDr><${element}></${element}></xdr:wsDr>`,
    })
    expect(detectUnsupportedFeatures(bytes)).toEqual([
      { kind: 'drawing', detail: 'xl/drawings/drawing1.xml' },
    ])
  }
})

it('does not report workbook drawings that contain only cell pictures', () => {
  const bytes = packageOf({
    'xl/drawings/drawing1.xml': '<xdr:wsDr><xdr:twoCellAnchor><xdr:pic></xdr:pic>'
      + '</xdr:twoCellAnchor></xdr:wsDr>',
    'xl/drawings/_rels/drawing1.xml.rels': '<Relationships/>',
    'xl/media/image1.png': new Uint8Array([0]),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report images for a package without media parts', () => {
  expect(detectUnsupportedFeatures(packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT }))).toEqual([])
})

it('leaves document-body and worksheet hyperlinks alone', () => {
  const document = packageOf({
    [DOCUMENT_XML]: bodyOf('<w:p><w:hyperlink r:id="rId1"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>'),
  })
  expect(detectUnsupportedFeatures(document)).toEqual([])

  const worksheet = packageOf({
    [WORKSHEET_XML]: '<worksheet><hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks></worksheet>',
  })
  expect(detectUnsupportedFeatures(worksheet)).toEqual([])
})

it('does not report hyperlinks for plain text or plain cells', () => {
  expect(detectUnsupportedFeatures(packageOf({ [DOCUMENT_XML]: bodyOf('<w:p><w:r><w:t>link</w:t></w:r></w:p>') })))
    .toEqual([])
  expect(detectUnsupportedFeatures(packageOf({ [WORKSHEET_XML]: PLAIN_WORKSHEET }))).toEqual([])
})

it('does not report list numbering from paragraph properties or the numbering part', () => {
  const paragraph = packageOf({
    [DOCUMENT_XML]: bodyOf('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>'),
  })
  expect(detectUnsupportedFeatures(paragraph)).toEqual([])

  const part = packageOf({
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/numbering.xml': '<w:numbering><w:abstractNum w:abstractNumId="0"/><w:num w:numId="1"/></w:numbering>',
  })
  expect(detectUnsupportedFeatures(part)).toEqual([])
})

it('does not report numbering for a styled paragraph or an empty numbering part', () => {
  const styled = packageOf({
    [DOCUMENT_XML]: bodyOf('<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>plain</w:t></w:r></w:p>'),
    'word/numbering.xml': '<w:numbering/>',
  })
  expect(detectUnsupportedFeatures(styled)).toEqual([])
})

it('does not report run formatting or paragraph alignment the writer maps', () => {
  const variants = [
    '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>',
    '<w:rPr><w:sz w:val="40"/></w:rPr>',
    '<w:rPr><w:szCs w:val="40"/></w:rPr>',
    '<w:rPr><w:color w:val="FF0000"/></w:rPr>',
  ]
  for (const properties of variants) {
    const bytes = packageOf({ [DOCUMENT_XML]: bodyOf(`<w:p><w:r>${properties}<w:t>styled</w:t></w:r></w:p>`) })
    expect(detectUnsupportedFeatures(bytes)).toEqual([])
  }

  const aligned = packageOf({
    [DOCUMENT_XML]: bodyOf('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>centered</w:t></w:r></w:p>'),
  })
  expect(detectUnsupportedFeatures(aligned)).toEqual([])
})

it('does not report a run that declares only a default font element', () => {
  const bytes = packageOf({
    [DOCUMENT_XML]: bodyOf('<w:p><w:r><w:rPr><w:rFonts/></w:rPr><w:t>x</w:t></w:r></w:p>'),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report bold, italic, and underline runs as unsupported formatting', () => {
  const bytes = packageOf({
    [DOCUMENT_XML]: bodyOf(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
      + '<w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr><w:t>supported</w:t></w:r></w:p>',
    ),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report strikethrough or highlighting the writer maps', () => {
  for (const property of [
    '<w:strike/>',
    '<w:strike w:val="true"/>',
    '<w:dstrike/>',
    '<w:highlight w:val="yellow"/>',
  ]) {
    const bytes = packageOf({ [DOCUMENT_XML]: bodyOf(`<w:p><w:r><w:rPr>${property}</w:rPr><w:t>x</w:t></w:r></w:p>`) })
    expect(detectUnsupportedFeatures(bytes)).toEqual([])
  }
})

it('does not report disabled strike or a cleared highlight', () => {
  for (const property of ['<w:strike w:val="false"/>', '<w:highlight w:val="none"/>']) {
    const bytes = packageOf({ [DOCUMENT_XML]: bodyOf(`<w:p><w:r><w:rPr>${property}</w:rPr><w:t>x</w:t></w:r></w:p>`) })
    expect(detectUnsupportedFeatures(bytes)).toEqual([])
  }
})

it('does not report paragraph spacing or indentation the writer maps', () => {
  for (const property of [
    '<w:spacing w:after="240" w:line="360"/>',
    '<w:ind w:left="720" w:firstLine="360"/>',
  ]) {
    const bytes = packageOf({ [DOCUMENT_XML]: bodyOf(`<w:p><w:pPr>${property}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`) })
    expect(detectUnsupportedFeatures(bytes)).toEqual([])
  }
})

it('does not report a paragraph that only names a style', () => {
  const bytes = packageOf({
    [DOCUMENT_XML]: bodyOf('<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>plain</w:t></w:r></w:p>'),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report cell borders the provider maps, and leaves paragraph and run borders alone', () => {
  const mapped = [
    '<w:p><w:pPr><w:pBdr><w:top w:val="single"/></w:pBdr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
    '<w:p><w:r><w:rPr><w:bdr w:val="single"/></w:rPr><w:t>x</w:t></w:r></w:p>',
  ]
  for (const content of mapped) {
    expect(detectUnsupportedFeatures(packageOf({ [DOCUMENT_XML]: bodyOf(content) }))).toEqual([])
  }

  const cell = '<w:tbl><w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="single"/>'
    + '</w:tcBorders></w:tcPr><w:p/></w:tc></w:tr></w:tbl>'
  expect(detectUnsupportedFeatures(packageOf({ [DOCUMENT_XML]: bodyOf(cell) }))).toEqual([])
})

it('does not report the table borders the writer emits', () => {
  const bytes = packageOf({
    [DOCUMENT_XML]: bodyOf(
      '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr>'
      + '<w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>',
    ),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report footnote and endnote parts or body references', () => {
  for (const [path, note] of [['word/footnotes.xml', 'footnote'], ['word/endnotes.xml', 'endnote']] as const) {
    const part = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: `<w:${note}s/>` })
    expect(detectUnsupportedFeatures(part)).toEqual([])

    const reference = packageOf({
      [DOCUMENT_XML]: bodyOf(`<w:p><w:r><w:${note}Reference w:id="1"/></w:r></w:p>`),
    })
    expect(detectUnsupportedFeatures(reference)).toEqual([])
  }
})

it('does not report notes for a document without note parts or references', () => {
  expect(detectUnsupportedFeatures(packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT }))).toEqual([])
})

it('does not report multiple sections', () => {
  const bytes = packageOf({
    [DOCUMENT_XML]: bodyOf(
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>'
      + '<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
    ),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report the single default section every document carries', () => {
  const bytes = packageOf({
    [DOCUMENT_XML]: bodyOf(
      '<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
    ),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report an image relationship whose target is external', () => {
  const bytes = packageOf({
    [DOCUMENT_XML]: bodyOf('<w:p><w:r><w:drawing><a:blip r:link="rId9"/></w:drawing></w:r></w:p>'),
    'word/_rels/document.xml.rels': '<Relationships><Relationship Id="rId9"'
      + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
      + ' Target="https://example.test/a.png" TargetMode="External"/></Relationships>',
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report an external image reference without a relationship part', () => {
  const bytes = packageOf({
    [DOCUMENT_XML]: bodyOf('<w:p><w:r><w:drawing><a:blip r:link="rId9"/></w:drawing></w:r></w:p>'),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report preserved document media or external hyperlink relationships', () => {
  const embedded = packageOf({
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/media/image1.png': new Uint8Array([0]),
    'word/_rels/document.xml.rels': '<Relationships><Relationship Id="rId9"'
      + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
      + ' Target="media/image1.png"/></Relationships>',
  })
  expect(detectUnsupportedFeatures(embedded)).toEqual([])

  const hyperlink = packageOf({
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/_rels/document.xml.rels': '<Relationships><Relationship Id="rId9"'
      + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'
      + ' Target="https://example.test/" TargetMode="External"/></Relationships>',
  })
  expect(detectUnsupportedFeatures(hyperlink)).toEqual([])
})

it('does not report merged cell spans the provider maps', () => {
  const horizontal = packageOf({
    [DOCUMENT_XML]: bodyOf(
      '<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>',
    ),
  })
  expect(detectUnsupportedFeatures(horizontal)).toEqual([])

  for (const value of ['restart', 'continue']) {
    const vertical = packageOf({
      [DOCUMENT_XML]: bodyOf(
        `<w:tbl><w:tr><w:tc><w:tcPr><w:vMerge w:val="${value}"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`,
      ),
    })
    expect(detectUnsupportedFeatures(vertical)).toEqual([])
  }
})

it('does not report a single-column span or an unmerged cell', () => {
  const bytes = packageOf({
    [DOCUMENT_XML]: bodyOf(
      '<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="1"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>',
    ),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report a table nested inside a cell', () => {
  const bytes = packageOf({
    [DOCUMENT_XML]: bodyOf(
      '<w:tbl><w:tr><w:tc><w:p/>'
      + '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
      + '</w:tc></w:tr></w:tbl>',
    ),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report sibling tables as nested', () => {
  const table = '<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
  const bytes = packageOf({ [DOCUMENT_XML]: bodyOf(`${table}${table}`) })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report sibling tables after a self-closing cell as nested', () => {
  const table = '<w:tbl><w:tr><w:tc/></w:tr></w:tbl>'
  const bytes = packageOf({ [DOCUMENT_XML]: bodyOf(`${table}${table}`) })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('does not report conditional formatting or data validation', () => {
  const conditional = packageOf({
    [WORKSHEET_XML]: '<worksheet><conditionalFormatting sqref="A1:A5"/></worksheet>',
  })
  expect(detectUnsupportedFeatures(conditional)).toEqual([])

  const validation = packageOf({
    [WORKSHEET_XML]: '<worksheet><dataValidations><dataValidation type="list" sqref="B1"/></dataValidations></worksheet>',
  })
  expect(detectUnsupportedFeatures(validation)).toEqual([])
})

it('does not report frozen or split panes, which the writer now maps', () => {
  const frozen = packageOf({
    [WORKSHEET_XML]: '<worksheet><sheetViews><sheetView><pane xSplit="1" state="frozen"/></sheetView></sheetViews></worksheet>',
  })
  expect(detectUnsupportedFeatures(frozen)).toEqual([])

  const split = packageOf({
    [WORKSHEET_XML]: '<worksheet><sheetViews><sheetView><pane xSplit="1" state="split"/></sheetView></sheetViews></worksheet>',
  })
  expect(detectUnsupportedFeatures(split)).toEqual([])
})

it('does not report a sheet tab color, which the writer now maps', () => {
  const bytes = packageOf({
    [WORKSHEET_XML]: '<worksheet><sheetPr><tabColor rgb="FFFF0000"/></sheetPr></worksheet>',
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('reports a workbook drawing kind once across drawing parts', () => {
  const bytes = packageOf({
    'xl/drawings/drawing1.xml': '<xdr:wsDr><xdr:sp/></xdr:wsDr>',
    'xl/drawings/drawing2.xml': '<xdr:wsDr><xdr:sp/></xdr:wsDr>',
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([
    { kind: 'drawing', detail: 'xl/drawings/drawing1.xml' },
  ])
})

it('keeps the minimal editable web fixture free of detected losses across a round trip', async () => {
  const text = 'Office preview 中文文档'
  const bytes = packageOf({
    [DOCUMENT_XML]: '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
      + `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
      + `<w:p><w:r><w:br w:type="page"/><w:t>${text}</w:t></w:r></w:p>`
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
      + '</w:body></w:document>',
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  const original = await readDocument(bytes)
  expect(original.body?.sectionBreaks?.[0]?.pageSize).toEqual({ width: 816, height: 1056 })
  expect(original.body?.sectionBreaks?.[0]?.pageOrient).toBe(PageOrientType.PORTRAIT)
  const rewritten = await writeDocument(original)
  expect(detectUnsupportedFeatures(rewritten)).toEqual([])
  expect(strFromU8(unzipSync(rewritten)[DOCUMENT_XML] ?? new Uint8Array()))
    .toContain('<w:pgSz w:w="12240" w:h="15840"/>')
  const reopened = await readDocument(rewritten)
  expect(reopened.body?.sectionBreaks?.[0]?.pageSize).toEqual({ width: 816, height: 1056 })
  expect(reopened.body?.sectionBreaks?.[0]?.pageOrient).toBe(PageOrientType.PORTRAIT)
  expect(reopened.body?.dataStream).toBe(original.body?.dataStream)
})

it('keeps the supported round-trip fixtures free of detected losses', async () => {
  const docx = new Uint8Array(await readFile(DOCX_FIXTURE))
  expect(detectUnsupportedFeatures(docx)).toEqual([])
  const originalDocument = await readDocument(docx)
  const rewrittenDocument = await writeDocument(originalDocument)
  expect(detectUnsupportedFeatures(rewrittenDocument)).toEqual([])
  expect((await readDocument(rewrittenDocument)).body?.dataStream).toBe(originalDocument.body?.dataStream)

  const xlsx = new Uint8Array(await readFile(XLSX_FIXTURE))
  expect(detectUnsupportedFeatures(xlsx)).toEqual([])
  const originalWorkbook = await readWorkbook(xlsx)
  const rewrittenWorkbook = await writeWorkbook(originalWorkbook)
  expect(detectUnsupportedFeatures(rewrittenWorkbook)).toEqual([])
  const reopened = await readWorkbook(rewrittenWorkbook)
  expect(reopened.sheetOrder.map(id => reopened.sheets[id]?.name))
    .toEqual(originalWorkbook.sheetOrder.map(id => originalWorkbook.sheets[id]?.name))
})

it('reports no preserved or coupled constructs for a clean package', () => {
  const bytes = packageOf({
    '[Content_Types].xml': '<Types/>',
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
  })
  expect(detectPreserved(bytes)).toEqual({ preserved: new Set(), coupled: [] })
})

it('reports every preserved construct class carried by package entries', () => {
  const cases = [
    ['word/vbaProject.bin', 'macro'],
    ['xl/charts/chart1.xml', 'chart'],
    ['xl/pivotTables/pivotTable1.xml', 'pivot-table'],
    ['xl/pivotCache/pivotCacheDefinition1.xml', 'pivot-table'],
    ['xl/tables/table1.xml', 'excel-table'],
    ['xl/calcChain.xml', 'calc-chain'],
    ['word/comments.xml', 'comment'],
    ['word/commentsExtended.xml', 'comment'],
    ['xl/threadedComments/threadedComment1.xml', 'comment'],
    ['word/embeddings/oleObject1.bin', 'embedded-object'],
    ['word/drawings/drawing1.xml', 'drawing'],
    ['word/vmlDrawing1.vml', 'shape'],
    ['word/theme/theme1.xml', 'theme'],
    ['xl/theme/theme1.xml', 'theme'],
    ['word/fontTable.xml', 'font-table'],
    ['word/webSettings.xml', 'web-settings'],
    ['customXml/item1.xml', 'custom-xml'],
    ['xl/externalLinks/externalLink1.xml', 'external-link'],
    ['docProps/core.xml', 'document-properties'],
    ['docProps/app.xml', 'document-properties'],
  ] as const
  for (const [path, kind] of cases) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: '<part/>' })
    expect(preservedKindsOf(bytes), path).toEqual([kind])
  }
})

it('reports defined names from the workbook part', () => {
  const named = packageOf({
    'xl/workbook.xml': '<workbook><definedNames>'
      + '<definedName name="Total">Sheet1!$A$1</definedName></definedNames></workbook>',
  })
  expect(preservedKindsOf(named)).toEqual(['defined-names'])

  const plain = packageOf({ 'xl/workbook.xml': '<workbook><sheets><sheet name="S"/></sheets></workbook>' })
  expect(preservedKindsOf(plain)).toEqual([])
})

it('reports the macro fixture as preserved rather than as a refusal list', async () => {
  const bytes = new Uint8Array(await readFile(MACRO_FIXTURE))
  const preserved = detectPreserved(bytes)
  expect([...preserved.preserved]).toEqual(['macro'])
  expect(preserved.coupled).toEqual([])
  expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'macro', detail: 'word/vbaProject.bin' }])
})

it('reports the macro class for macro-enabled docm and xlsm fixtures', async () => {
  const cases = [['macro.docm', 'word/vbaProject.bin'], ['macro.xlsm', 'xl/vbaProject.bin']] as const
  for (const [name, part] of cases) {
    const bytes = new Uint8Array(await readFile(new URL(`./fixtures/${name}`, import.meta.url)))
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'macro', detail: part }])
    expect(preservedKindsOf(bytes)).toContain('macro')
  }
})

it('reports tracked changes as coupled markers without refusing the file', async () => {
  const bytes = new Uint8Array(await readFile(TRACKED_FIXTURE))
  const result = detectPreserved(bytes)
  expect([...result.preserved]).toEqual([])
  expect(result.coupled).toHaveLength(2)
  expect(result.coupled.map(marker => ({ kind: marker.kind, part: marker.part }))).toEqual([
    { kind: 'tracked-change', part: DOCUMENT_XML },
    { kind: 'tracked-change', part: DOCUMENT_XML },
  ])
  expect(result.coupled[0]?.offset).toBeGreaterThan(0)
  expect(result.coupled[1]?.offset).toBeGreaterThan(result.coupled[0]?.offset ?? 0)
})

it('marks each revision element as a tracked-change marker', () => {
  const xml = bodyOf(
    '<w:p><w:ins w:id="1"><w:r><w:t>a</w:t></w:r></w:ins>'
    + '<w:del w:id="2"><w:r><w:delText>b</w:delText></w:r></w:del></w:p>',
  )
  const result = detectPreserved(packageOf({ [DOCUMENT_XML]: xml }))
  expect([...result.preserved]).toEqual([])
  expect(result.coupled).toEqual([
    { kind: 'tracked-change', part: DOCUMENT_XML, offset: xml.indexOf('<w:ins ') },
    { kind: 'tracked-change', part: DOCUMENT_XML, offset: xml.indexOf('<w:del ') },
  ])
})

it('marks comment ranges and comment references as coupled and preserves the comment part', () => {
  const xml = bodyOf(
    '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>note</w:t></w:r>'
    + '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>',
  )
  const result = detectPreserved(packageOf({
    [DOCUMENT_XML]: xml,
    'word/comments.xml': '<w:comments><w:comment w:id="0"/></w:comments>',
  }))
  expect([...result.preserved]).toEqual(['comment'])
  expect(result.coupled).toEqual([
    { kind: 'comment-range', part: DOCUMENT_XML, offset: xml.indexOf('<w:commentRangeStart') },
    { kind: 'comment-range', part: DOCUMENT_XML, offset: xml.indexOf('<w:commentRangeEnd') },
    { kind: 'comment-range', part: DOCUMENT_XML, offset: xml.indexOf('<w:commentReference') },
  ])
})

it('reports a field as preserved and marks its owned text as coupled', () => {
  const xml = bodyOf(
    '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>',
  )
  const result = detectPreserved(packageOf({ [DOCUMENT_XML]: xml }))
  expect([...result.preserved]).toEqual(['field'])
  expect(result.coupled).toEqual([
    { kind: 'field-owned-text', part: DOCUMENT_XML, offset: xml.indexOf('<w:fldSimple') },
  ])
})

it('marks complex field characters as coupled', () => {
  const xml = bodyOf(
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:t>1</w:t></w:r>'
    + '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
  )
  const result = detectPreserved(packageOf({ [DOCUMENT_XML]: xml }))
  expect([...result.preserved]).toEqual(['field'])
  expect(result.coupled.map(marker => marker.offset)).toEqual([
    xml.indexOf('<w:fldChar w:fldCharType="begin"'),
    xml.indexOf('<w:fldChar w:fldCharType="end"'),
  ])
})

it('reports bookmark anchors as preserved', () => {
  const xml = bodyOf(
    '<w:p><w:bookmarkStart w:id="0" w:name="mark"/><w:r><w:t>x</w:t></w:r>'
    + '<w:bookmarkEnd w:id="0"/></w:p>',
  )
  const result = detectPreserved(packageOf({ [DOCUMENT_XML]: xml }))
  expect([...result.preserved]).toEqual(['bookmark'])
  expect(result.coupled).toEqual([])
})

it('reports a content control as preserved and its wrap as coupled', () => {
  const xml = bodyOf(
    '<w:sdt><w:sdtPr><w:alias w:val="Name"/></w:sdtPr>'
    + '<w:sdtContent><w:p><w:r><w:t>Alice</w:t></w:r></w:p></w:sdtContent></w:sdt>',
  )
  const result = detectPreserved(packageOf({ [DOCUMENT_XML]: xml }))
  expect([...result.preserved]).toEqual(['content-control'])
  expect(result.coupled).toEqual([
    { kind: 'content-control-wrap', part: DOCUMENT_XML, offset: xml.indexOf('<w:sdt>') },
  ])
})

it('marks alternate content as ambiguous and coupled', () => {
  const xml = bodyOf(
    '<w:p><mc:AlternateContent><mc:Choice Requires="wps">'
    + '<w:r><w:t>choice</w:t></w:r></mc:Choice>'
    + '<mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback>'
    + '</mc:AlternateContent></w:p>',
  )
  const result = detectPreserved(packageOf({ [DOCUMENT_XML]: xml }))
  expect(result.coupled).toEqual([
    { kind: 'ambiguous-alternate-content', part: DOCUMENT_XML, offset: xml.indexOf('<mc:AlternateContent>') },
  ])
})

it('marks rich-text shared strings as coupled but not plain ones', () => {
  const shared = '<sst><si><t>plain</t></si><si><r><rPr><b/></rPr><t>rich</t></r></si></sst>'
  const bytes = packageOf({ [WORKSHEET_XML]: PLAIN_WORKSHEET, 'xl/sharedStrings.xml': shared })
  expect(markersOf(bytes)).toEqual([
    {
      kind: 'rich-text-shared-string',
      part: 'xl/sharedStrings.xml',
      offset: shared.indexOf('<si><r>'),
    },
  ])
})

it('marks shared and array formulas as coupled', () => {
  const xml = '<worksheet><sheetData><row r="1">'
    + '<c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c>'
    + '<c r="C1"><f t="array" ref="C1:C2">A1:A2*2</f><v>4</v></c>'
    + '<c r="D1"><f>A1+1</f><v>3</v></c>'
    + '</row></sheetData></worksheet>'
  expect(markersOf(packageOf({ [WORKSHEET_XML]: xml }))).toEqual([
    { kind: 'shared-formula', part: WORKSHEET_XML, offset: xml.indexOf('<f t="shared"') },
    { kind: 'array-formula', part: WORKSHEET_XML, offset: xml.indexOf('<f t="array"') },
  ])
})

it('keeps preserved classes separate from coupled markers in one package', () => {
  const xml = bodyOf('<w:p><w:ins w:id="1"><w:r><w:t>x</w:t></w:r></w:ins></w:p>')
  const result = detectPreserved(packageOf({
    [DOCUMENT_XML]: xml,
    'word/vbaProject.bin': new Uint8Array([0]),
    'word/theme/theme1.xml': '<theme/>',
  }))
  expect([...result.preserved]).toEqual(['macro', 'theme'])
  expect(result.coupled).toEqual([
    { kind: 'tracked-change', part: DOCUMENT_XML, offset: xml.indexOf('<w:ins ') },
  ])
})

it('sorts coupled markers by part name then offset', () => {
  const documentXml = bodyOf('<w:p><w:ins w:id="1"><w:r><w:t>x</w:t></w:r></w:ins></w:p>')
  const sheetXml = '<worksheet><sheetData><row r="1">'
    + '<c r="A1"><f t="array" ref="A1" si="0">1</f></c>'
    + '</row></sheetData></worksheet>'
  const result = detectPreserved(packageOf({
    [DOCUMENT_XML]: documentXml,
    [WORKSHEET_XML]: sheetXml,
  }))
  expect(result.coupled.map(marker => marker.part))
    .toEqual([DOCUMENT_XML, WORKSHEET_XML])
})

it('scans header and footer story parts for coupled constructs', () => {
  const header = '<w:hdr><w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:hdr>'
  const result = detectPreserved(packageOf({
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/header1.xml': header,
  }))
  expect([...result.preserved]).toEqual(['field'])
  expect(result.coupled).toEqual([
    { kind: 'field-owned-text', part: 'word/header1.xml', offset: header.indexOf('<w:fldSimple') },
  ])
})

it('rejects classification input over maxFileBytes before parsing', () => {
  expect(codeOfThrown(() => detectPreserved(new Uint8Array(32), {
    ...DEFAULT_OFFICE_EXCHANGE_CONFIG,
    maxFileBytes: 8,
  }))).toBe('too-large')
})

it('classifies non-zip classification input as conversion-failed', () => {
  expect(codeOfThrown(() => detectPreserved(strToU8('not a zip archive')))).toBe('conversion-failed')
})
