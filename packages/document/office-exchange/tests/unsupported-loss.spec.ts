/**
 * Round-trip evidence for every construct that detection refuses: each source
 * package carries the construct, detection reports its kind, and a read/write
 * round-trip through the real provider no longer contains it.
 */
import ExcelJS from 'exceljs'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { expect, it } from 'vitest'
import { detectUnsupportedFeatures } from '../src/detect.ts'
import { readDocument, writeDocument } from '../src/provider/docx.ts'
import { readWorkbook, writeWorkbook } from '../src/provider/xlsx.ts'
import type { UnsupportedFeature } from '../src/types.ts'

const DOCUMENT_XML = 'word/document.xml'
const WORKSHEET_XML = 'xl/worksheets/sheet1.xml'
const DOCUMENT_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

function documentPackage(documentXml: string, extra: Record<string, string | Uint8Array> = {}): Uint8Array {
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8('<Types/>'),
    [DOCUMENT_XML]: strToU8(`<w:document ${DOCUMENT_NAMESPACE}><w:body>${documentXml}</w:body></w:document>`),
  }
  for (const [name, content] of Object.entries(extra)) {
    files[name] = typeof content === 'string' ? strToU8(content) : content
  }
  return zipSync(files)
}

function entriesOf(bytes: Uint8Array): Record<string, string> {
  const entries = unzipSync(bytes)
  const texts: Record<string, string> = {}
  for (const [name, content] of Object.entries(entries)) texts[name] = strFromU8(content)
  return texts
}

async function documentRoundTrip(bytes: Uint8Array): Promise<Record<string, string>> {
  return entriesOf(await writeDocument(await readDocument(bytes)))
}

async function workbookRoundTrip(build: (sheet: ExcelJS.Worksheet) => void): Promise<{
  source: Uint8Array
  written: Record<string, string>
}> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Data')
  sheet.getCell('A1').value = 'kept'
  build(sheet)
  const source = new Uint8Array(await workbook.xlsx.writeBuffer())
  return { source, written: entriesOf(await writeWorkbook(await readWorkbook(source))) }
}

function sheetXml(entries: Record<string, string>): string {
  return entries[WORKSHEET_XML] ?? ''
}

function expectDetected(bytes: Uint8Array, feature: UnsupportedFeature): void {
  expect(detectUnsupportedFeatures(bytes)).toContainEqual(feature)
}

it('drops an inline image and reports it', async () => {
  const image = '<w:p><w:r><w:drawing><wp:inline xmlns:wp="urn:probe"><a:graphic'
    + ' xmlns:a="urn:probe"><a:graphicData><pic:pic xmlns:pic="urn:probe"><pic:blipFill>'
    + '<a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline>'
    + '</w:drawing></w:r></w:p>'
  const bytes = documentPackage(image, { 'word/media/image1.png': new Uint8Array([1, 2, 3]) })
  expectDetected(bytes, { kind: 'image', detail: 'word/media/image1.png' })
  const written = await documentRoundTrip(bytes)
  expect(Object.keys(written)).not.toContain('word/media/image1.png')
  expect(written[DOCUMENT_XML]).not.toContain('w:drawing')
})

it('drops a hyperlink and its target text, and reports it', async () => {
  const bytes = documentPackage(
    '<w:p><w:hyperlink r:id="rId9"><w:r><w:t>link text</w:t></w:r></w:hyperlink></w:p>',
  )
  expectDetected(bytes, { kind: 'hyperlink', detail: DOCUMENT_XML })
  const written = await documentRoundTrip(bytes)
  expect(written[DOCUMENT_XML]).not.toContain('w:hyperlink')
  expect(written[DOCUMENT_XML]).not.toContain('link text')
})

it('drops list numbering and the numbering part, and reports it', async () => {
  const bytes = documentPackage(
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
    + '<w:r><w:t>item</w:t></w:r></w:p>',
    {
      'word/numbering.xml': '<w:numbering><w:abstractNum w:abstractNumId="0"/><w:num w:numId="1"/></w:numbering>',
    },
  )
  expectDetected(bytes, { kind: 'numbering', detail: DOCUMENT_XML })
  const written = await documentRoundTrip(bytes)
  expect(Object.keys(written)).not.toContain('word/numbering.xml')
  expect(written[DOCUMENT_XML]).not.toContain('w:numPr')
})

it('drops run formatting and paragraph alignment, and reports it', async () => {
  const bytes = documentPackage(
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr>'
    + '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="40"/><w:color w:val="FF0000"/>'
    + '</w:rPr><w:t>styled</w:t></w:r></w:p>',
  )
  expectDetected(bytes, { kind: 'run-formatting', detail: DOCUMENT_XML })
  const writtenXml = (await documentRoundTrip(bytes))[DOCUMENT_XML] ?? ''
  expect(writtenXml).not.toContain('w:rFonts')
  expect(writtenXml).not.toContain('w:sz')
  expect(writtenXml).not.toContain('w:color')
  expect(writtenXml).not.toContain('w:jc')
})

it('drops merged-cell spans, and reports them', async () => {
  const bytes = documentPackage(
    '<w:tbl><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="1200"/></w:tblGrid>'
    + '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc></w:tr>'
    + '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>tall</w:t></w:r></w:p></w:tc>'
    + '<w:tc><w:p><w:r><w:t>next</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  )
  expectDetected(bytes, { kind: 'merged-cells', detail: DOCUMENT_XML })
  const writtenXml = (await documentRoundTrip(bytes))[DOCUMENT_XML] ?? ''
  expect(writtenXml).not.toContain('w:gridSpan')
  expect(writtenXml).not.toContain('w:vMerge')
})

it('drops a nested table and its content, and reports it', async () => {
  const bytes = documentPackage(
    '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>outer</w:t></w:r></w:p>'
    + '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    + '</w:tc></w:tr></w:tbl>',
  )
  expectDetected(bytes, { kind: 'nested-table', detail: DOCUMENT_XML })
  const writtenXml = (await documentRoundTrip(bytes))[DOCUMENT_XML] ?? ''
  expect(writtenXml).not.toContain('inner')
})

it('drops conditional formatting, and reports it', async () => {
  const { source, written } = await workbookRoundTrip((sheet) => {
    sheet.addConditionalFormatting({
      ref: 'A1:A5',
      rules: [{ type: 'expression', formulae: ['$A1>3'], priority: 1, style: { font: { bold: true } } }],
    })
  })
  expectDetected(source, { kind: 'conditional-formatting', detail: WORKSHEET_XML })
  expect(sheetXml(written)).not.toContain('conditionalFormatting')
})

it('drops data validation, and reports it', async () => {
  const { source, written } = await workbookRoundTrip((sheet) => {
    sheet.getCell('B1').dataValidation = { type: 'list', allowBlank: true, formulae: ['"a,b"'] }
  })
  expectDetected(source, { kind: 'data-validation', detail: WORKSHEET_XML })
  expect(sheetXml(written)).not.toContain('dataValidation')
})

it('drops frozen panes, and reports them', async () => {
  const { source, written } = await workbookRoundTrip((sheet) => {
    sheet.views = [{ state: 'frozen', xSplit: 1, ySplit: 1, topLeftCell: 'B2', activeCell: 'B2' }]
  })
  expectDetected(source, { kind: 'freeze-panes', detail: WORKSHEET_XML })
  expect(sheetXml(written)).not.toContain('<pane')
})

it('drops a sheet tab color, and reports it', async () => {
  const { source, written } = await workbookRoundTrip((sheet) => {
    sheet.properties.tabColor = { argb: 'FFFF0000' }
  })
  expectDetected(source, { kind: 'tab-color', detail: WORKSHEET_XML })
  expect(sheetXml(written)).not.toContain('tabColor')
})

it('drops a cell hyperlink target, and reports it', async () => {
  const { source, written } = await workbookRoundTrip((sheet) => {
    sheet.getCell('C1').value = { text: 'link cell', hyperlink: 'https://example.test/' }
  })
  expectDetected(source, { kind: 'hyperlink', detail: WORKSHEET_XML })
  expect(sheetXml(written)).not.toContain('<hyperlink')
  const relationships = Object.entries(written)
    .filter(([name]) => name.endsWith('.rels'))
    .map(([, text]) => text)
    .join('')
  expect(relationships).not.toContain('example.test')
})
