/** Unsupported-feature detection over in-memory OOXML packages. */
import { readFile } from 'node:fs/promises'
import { strToU8, zipSync } from 'fflate'
import { expect, it } from 'vitest'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from '../src/config.ts'
import { detectUnsupportedFeatures } from '../src/detect.ts'
import type { OfficeExchangeErrorCode } from '../src/errors.ts'
import { OfficeExchangeError } from '../src/errors.ts'
import { readDocument, writeDocument } from '../src/provider/docx.ts'
import { readWorkbook, writeWorkbook } from '../src/provider/xlsx.ts'

const DOCUMENT_XML = 'word/document.xml'
const PLAIN_DOCUMENT = '<w:document><w:body><w:p/></w:body></w:document>'
const WORKSHEET_XML = 'xl/worksheets/sheet1.xml'
const PLAIN_WORKSHEET = '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'
const DOCX_FIXTURE = new URL('./fixtures/roundtrip.docx', import.meta.url)
const XLSX_FIXTURE = new URL('./fixtures/roundtrip.xlsx', import.meta.url)

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

it('detects header and footer parts', () => {
  for (const path of ['word/header1.xml', 'word/footer2.xml']) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: '<hdr/>' })
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'header-footer', detail: path }])
  }
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
    'word/header1.xml': '<hdr/>',
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
    { kind: 'header-footer', detail: 'word/header1.xml' },
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

it('detects inline image media parts', () => {
  for (const path of ['word/media/image1.png', 'xl/media/image1.png']) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: new Uint8Array([0]) })
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'image', detail: path }])
  }
})

it('does not report images for a package without media parts', () => {
  expect(detectUnsupportedFeatures(packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT }))).toEqual([])
})

it('detects hyperlinks in a document body and in a worksheet', () => {
  const document = packageOf({
    [DOCUMENT_XML]: bodyOf('<w:p><w:hyperlink r:id="rId1"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>'),
  })
  expect(detectUnsupportedFeatures(document)).toEqual([{ kind: 'hyperlink', detail: DOCUMENT_XML }])

  const worksheet = packageOf({
    [WORKSHEET_XML]: '<worksheet><hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks></worksheet>',
  })
  expect(detectUnsupportedFeatures(worksheet)).toEqual([{ kind: 'hyperlink', detail: WORKSHEET_XML }])
})

it('does not report hyperlinks for plain text or plain cells', () => {
  expect(detectUnsupportedFeatures(packageOf({ [DOCUMENT_XML]: bodyOf('<w:p><w:r><w:t>link</w:t></w:r></w:p>') })))
    .toEqual([])
  expect(detectUnsupportedFeatures(packageOf({ [WORKSHEET_XML]: PLAIN_WORKSHEET }))).toEqual([])
})

it('detects list numbering from paragraph properties and the numbering part', () => {
  const paragraph = packageOf({
    [DOCUMENT_XML]: bodyOf('<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>'),
  })
  expect(detectUnsupportedFeatures(paragraph)).toEqual([{ kind: 'numbering', detail: DOCUMENT_XML }])

  const part = packageOf({
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/numbering.xml': '<w:numbering><w:abstractNum w:abstractNumId="0"/><w:num w:numId="1"/></w:numbering>',
  })
  expect(detectUnsupportedFeatures(part)).toEqual([{ kind: 'numbering', detail: 'word/numbering.xml' }])
})

it('does not report numbering for a styled paragraph or an empty numbering part', () => {
  const styled = packageOf({
    [DOCUMENT_XML]: bodyOf('<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>plain</w:t></w:r></w:p>'),
    'word/numbering.xml': '<w:numbering/>',
  })
  expect(detectUnsupportedFeatures(styled)).toEqual([])
})

it('detects unsupported run formatting and paragraph alignment', () => {
  const variants = [
    '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>',
    '<w:rPr><w:sz w:val="40"/></w:rPr>',
    '<w:rPr><w:szCs w:val="40"/></w:rPr>',
    '<w:rPr><w:color w:val="FF0000"/></w:rPr>',
  ]
  for (const properties of variants) {
    const bytes = packageOf({ [DOCUMENT_XML]: bodyOf(`<w:p><w:r>${properties}<w:t>styled</w:t></w:r></w:p>`) })
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'run-formatting', detail: DOCUMENT_XML }])
  }

  const aligned = packageOf({
    [DOCUMENT_XML]: bodyOf('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>centered</w:t></w:r></w:p>'),
  })
  expect(detectUnsupportedFeatures(aligned)).toEqual([{ kind: 'run-formatting', detail: DOCUMENT_XML }])
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

it('detects merged cell spans', () => {
  const horizontal = packageOf({
    [DOCUMENT_XML]: bodyOf(
      '<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>',
    ),
  })
  expect(detectUnsupportedFeatures(horizontal)).toEqual([{ kind: 'merged-cells', detail: DOCUMENT_XML }])

  for (const value of ['restart', 'continue']) {
    const vertical = packageOf({
      [DOCUMENT_XML]: bodyOf(
        `<w:tbl><w:tr><w:tc><w:tcPr><w:vMerge w:val="${value}"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`,
      ),
    })
    expect(detectUnsupportedFeatures(vertical)).toEqual([{ kind: 'merged-cells', detail: DOCUMENT_XML }])
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

it('detects a table nested inside a cell', () => {
  const bytes = packageOf({
    [DOCUMENT_XML]: bodyOf(
      '<w:tbl><w:tr><w:tc><w:p/>'
      + '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
      + '</w:tc></w:tr></w:tbl>',
    ),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'nested-table', detail: DOCUMENT_XML }])
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

it('detects conditional formatting and data validation', () => {
  const conditional = packageOf({
    [WORKSHEET_XML]: '<worksheet><conditionalFormatting sqref="A1:A5"/></worksheet>',
  })
  expect(detectUnsupportedFeatures(conditional)).toEqual([
    { kind: 'conditional-formatting', detail: WORKSHEET_XML },
  ])

  const validation = packageOf({
    [WORKSHEET_XML]: '<worksheet><dataValidations><dataValidation type="list" sqref="B1"/></dataValidations></worksheet>',
  })
  expect(detectUnsupportedFeatures(validation)).toEqual([{ kind: 'data-validation', detail: WORKSHEET_XML }])
})

it('detects frozen panes but not split panes', () => {
  const frozen = packageOf({
    [WORKSHEET_XML]: '<worksheet><sheetViews><sheetView><pane xSplit="1" state="frozen"/></sheetView></sheetViews></worksheet>',
  })
  expect(detectUnsupportedFeatures(frozen)).toEqual([{ kind: 'freeze-panes', detail: WORKSHEET_XML }])

  const split = packageOf({
    [WORKSHEET_XML]: '<worksheet><sheetViews><sheetView><pane xSplit="1" state="split"/></sheetView></sheetViews></worksheet>',
  })
  expect(detectUnsupportedFeatures(split)).toEqual([])
})

it('detects a sheet tab color', () => {
  const bytes = packageOf({
    [WORKSHEET_XML]: '<worksheet><sheetPr><tabColor rgb="FFFF0000"/></sheetPr></worksheet>',
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'tab-color', detail: WORKSHEET_XML }])
})

it('reports a workbook kind once across worksheets', () => {
  const bytes = packageOf({
    [WORKSHEET_XML]: '<worksheet><conditionalFormatting sqref="A1:A5"/></worksheet>',
    'xl/worksheets/sheet2.xml': '<worksheet><conditionalFormatting sqref="B1:B5"/></worksheet>',
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([
    { kind: 'conditional-formatting', detail: WORKSHEET_XML },
  ])
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
