/**
 * Round-trip evidence for constructs that detection refuses or the provider
 * maps: each source package carries the construct, detection reports its kind
 * when it is still refused, and a read/write round-trip drops or preserves it
 * as documented.
 */
import ExcelJS from 'exceljs'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { expect, it } from 'vitest'
import { DashStyleType } from '@univerjs/core'
import type { IWorkbookData, IWorksheetData } from '@univerjs/core'
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
  bytes: Uint8Array
}> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Data')
  sheet.getCell('A1').value = 'kept'
  build(sheet)
  const source = new Uint8Array(await workbook.xlsx.writeBuffer())
  const bytes = await writeWorkbook(await readWorkbook(source))
  return { source, written: entriesOf(bytes), bytes }
}

function sheetXml(entries: Record<string, string>): string {
  return entries[WORKSHEET_XML] ?? ''
}

function sheetOf(workbook: IWorkbookData, name: string): Partial<IWorksheetData> {
  const id = workbook.sheetOrder.find(candidate => workbook.sheets[candidate]?.name === name)
  if (id === undefined) throw new Error(`missing sheet ${name}`)
  return workbook.sheets[id] ?? {}
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

it('keeps a hyperlink, its text, and its external target', async () => {
  const bytes = documentPackage(
    '<w:p><w:hyperlink r:id="rId9"><w:r><w:t>link text</w:t></w:r></w:hyperlink></w:p>',
    {
      'word/_rels/document.xml.rels': '<Relationships><Relationship Id="rId9"'
        + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'
        + ' Target="https://example.test/link" TargetMode="External"/></Relationships>',
    },
  )
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  const written = await documentRoundTrip(bytes)
  expect(written[DOCUMENT_XML]).toContain('<w:hyperlink r:id="rId')
  expect(written[DOCUMENT_XML]).toContain('link text')
  expect(written['word/_rels/document.xml.rels']).toContain('https://example.test/link')
})

it('keeps list numbering and the numbering part', async () => {
  const bytes = documentPackage(
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
    + '<w:r><w:t>item</w:t></w:r></w:p>',
    {
      'word/numbering.xml': '<w:numbering><w:abstractNum w:abstractNumId="0"/><w:num w:numId="1"/></w:numbering>',
    },
  )
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  const written = await documentRoundTrip(bytes)
  expect(Object.keys(written)).toContain('word/numbering.xml')
  expect(written[DOCUMENT_XML]).toContain('<w:numPr>')
})

it('keeps cell borders through a read/write round trip', async () => {
  const bytes = documentPackage(
    '<w:tbl><w:tr><w:tc><w:tcPr><w:tcBorders>'
    + '<w:top w:val="single" w:sz="12" w:color="FF0000"/>'
    + '<w:left w:val="dashed" w:sz="6" w:color="00FF00"/>'
    + '</w:tcBorders></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  )
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  const original = await readDocument(bytes)
  const cell = original.tableSource?.['table-1']?.tableRows[0]?.tableCells[0]
  expect(cell?.borderTop?.color).toEqual({ rgb: '#FF0000' })
  expect(cell?.borderTop?.width?.v).toBe(2)
  expect(cell?.borderLeft?.dashStyle).toBe(DashStyleType.DASH)
  const written = await writeDocument(original)
  expect(entriesOf(written)[DOCUMENT_XML]).toContain('w:tcBorders')
  const reopened = await readDocument(written)
  const rewritten = reopened.tableSource?.['table-1']?.tableRows[0]?.tableCells[0]
  expect(rewritten?.borderTop).toEqual(cell?.borderTop)
  expect(rewritten?.borderLeft).toEqual(cell?.borderLeft)
})

it('drops footnotes and endnotes with their parts, and reports them', async () => {
  const bytes = documentPackage(
    '<w:p><w:r><w:t>body</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r>'
    + '<w:r><w:endnoteReference w:id="2"/></w:r></w:p>',
    {
      'word/footnotes.xml': '<w:footnotes><w:footnote w:id="1"><w:p><w:r><w:t>footnote text</w:t></w:r></w:p></w:footnote></w:footnotes>',
      'word/endnotes.xml': '<w:endnotes><w:endnote w:id="2"><w:p><w:r><w:t>endnote text</w:t></w:r></w:p></w:endnote></w:endnotes>',
    },
  )
  expectDetected(bytes, { kind: 'footnotes', detail: 'word/footnotes.xml' })
  const written = await documentRoundTrip(bytes)
  expect(Object.keys(written)).not.toContain('word/footnotes.xml')
  expect(Object.keys(written)).not.toContain('word/endnotes.xml')
  expect(written[DOCUMENT_XML]).not.toContain('w:footnoteReference')
  expect(written[DOCUMENT_XML]).not.toContain('w:endnoteReference')
  expect(written[DOCUMENT_XML]).not.toContain('footnote text')
  expect(written[DOCUMENT_XML]).not.toContain('endnote text')
})

it('drops a linked image and its external target, and reports it', async () => {
  const bytes = documentPackage(
    '<w:p><w:r><w:drawing><wp:inline xmlns:wp="urn:probe"><a:graphic'
    + ' xmlns:a="urn:probe"><a:graphicData><pic:pic xmlns:pic="urn:probe"><pic:blipFill>'
    + '<a:blip r:link="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline>'
    + '</w:drawing></w:r></w:p>',
    {
      'word/_rels/document.xml.rels': '<Relationships><Relationship Id="rId7"'
        + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
        + ' Target="https://example.test/linked.png" TargetMode="External"/></Relationships>',
    },
  )
  expectDetected(bytes, { kind: 'external-image', detail: 'word/_rels/document.xml.rels' })
  const written = await documentRoundTrip(bytes)
  expect(written[DOCUMENT_XML]).not.toContain('r:link')
  expect(written['word/_rels/document.xml.rels']).not.toContain('example.test')
  expect(written['word/_rels/document.xml.rels']).not.toContain('TargetMode')
})

it('keeps merged-cell spans through a read/write round trip', async () => {
  const bytes = documentPackage(
    '<w:tbl><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="1200"/></w:tblGrid>'
    + '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc></w:tr>'
    + '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>tall</w:t></w:r></w:p></w:tc>'
    + '<w:tc><w:p><w:r><w:t>next</w:t></w:r></w:p></w:tc></w:tr>'
    + '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>'
    + '<w:tc><w:p><w:r><w:t>last</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  )
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  const original = await readDocument(bytes)
  const table = original.tableSource?.['table-1']
  expect(table?.tableRows[0]?.tableCells[0]?.columnSpan).toBe(2)
  expect(table?.tableRows[1]?.tableCells[0]?.rowSpan).toBe(2)
  expect(table?.tableRows[2]?.tableCells[0]?.rowSpan).toBe(0)
  const written = await writeDocument(original)
  const writtenXml = entriesOf(written)[DOCUMENT_XML] ?? ''
  expect(writtenXml).toContain('w:gridSpan')
  expect(writtenXml).toContain('w:vMerge')
  const reopened = await readDocument(written)
  const rewritten = reopened.tableSource?.['table-1']
  expect(rewritten?.tableRows[0]?.tableCells[0]?.columnSpan).toBe(2)
  expect(rewritten?.tableRows[1]?.tableCells[0]?.rowSpan).toBe(2)
  expect(rewritten?.tableRows[2]?.tableCells[0]?.rowSpan).toBe(0)
})

it('keeps a nested table and its content through a read/write round trip', async () => {
  const bytes = documentPackage(
    '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>outer</w:t></w:r></w:p>'
    + '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    + '</w:tc></w:tr></w:tbl>',
  )
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  const original = await readDocument(bytes)
  expect(original.body?.tables).toHaveLength(2)
  expect(original.body?.tables?.[1]?.tableId).toBe('table-2')
  expect(original.body?.dataStream).toContain('inner')
  const written = await writeDocument(original)
  expect(entriesOf(written)[DOCUMENT_XML]).toContain('inner')
  const reopened = await readDocument(written)
  expect(reopened.body?.tables).toHaveLength(2)
  expect(reopened.body?.dataStream).toContain('inner')
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

it('keeps frozen panes through a write/read round trip', async () => {
  const { source, written, bytes } = await workbookRoundTrip((sheet) => {
    sheet.views = [{ state: 'frozen', xSplit: 1, ySplit: 1, topLeftCell: 'B2', activeCell: 'B2' }]
  })
  expect(detectUnsupportedFeatures(source)).toEqual([])
  expect(sheetXml(written)).toContain('<pane')
  expect(sheetOf(await readWorkbook(bytes), 'Data').freeze).toEqual({
    xSplit: 1,
    ySplit: 1,
    startRow: 1,
    startColumn: 1,
  })
})

it('keeps a sheet tab color through a write/read round trip', async () => {
  const { source, written, bytes } = await workbookRoundTrip((sheet) => {
    sheet.properties.tabColor = { argb: 'FFFF0000' }
  })
  expect(detectUnsupportedFeatures(source)).toEqual([])
  expect(sheetXml(written)).toContain('tabColor')
  expect(sheetOf(await readWorkbook(bytes), 'Data').tabColor).toBe('#FF0000')
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
