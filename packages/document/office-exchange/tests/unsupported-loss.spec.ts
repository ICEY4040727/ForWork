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

const DOCUMENT_XML = 'word/document.xml'
const WORKSHEET_XML = 'xl/worksheets/sheet1.xml'
const ONE_PIXEL_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
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

async function workbookRoundTrip(build: (sheet: ExcelJS.Worksheet, workbook: ExcelJS.Workbook) => void): Promise<{
  source: Uint8Array
  written: Record<string, string>
  bytes: Uint8Array
}> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Data')
  sheet.getCell('A1').value = 'kept'
  build(sheet, workbook)
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

function resourceOf(workbook: IWorkbookData, name: string): unknown {
  const entry = workbook.resources?.find(resource => resource.name === name)
  if (entry === undefined) return undefined
  return JSON.parse(entry.data)
}

async function conditionalFormattingRulesOf(bytes: Uint8Array): Promise<Array<{
  rule: Record<string, unknown>
  ranges: unknown
}>> {
  const workbook = await readWorkbook(bytes)
  const resources = resourceOf(workbook, 'SHEET_CONDITIONAL_FORMATTING_PLUGIN') as
    Record<string, Array<{ rule: Record<string, unknown>; ranges: unknown }>> | undefined
  return resources?.['sheet-1'] ?? []
}

async function dataValidationRulesOf(bytes: Uint8Array): Promise<Array<Record<string, unknown>>> {
  const workbook = await readWorkbook(bytes)
  const resources = resourceOf(workbook, 'SHEET_DATA_VALIDATION_PLUGIN') as
    Record<string, Array<Record<string, unknown>>> | undefined
  return resources?.['sheet-1'] ?? []
}

it('preserves an inline image, its media part, and its relationship', async () => {
  const drawing = '<w:r><w:drawing><wp:inline xmlns:wp="urn:probe"><a:graphic'
    + ' xmlns:a="urn:probe"><a:graphicData><pic:pic xmlns:pic="urn:probe"><pic:blipFill>'
    + '<a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline>'
    + '</w:drawing></w:r>'
  const media = new Uint8Array([0xff, 0xd8, 0xff, 0xd9])
  const relationship = '<Relationship Id="rId5"'
    + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
    + ' Target="media/image1.jpeg"/>'
  const bytes = documentPackage(`<w:p>${drawing}</w:p>`, {
    'word/media/image1.jpeg': media,
    'word/_rels/document.xml.rels': `<Relationships>${relationship}</Relationships>`,
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  const original = await readDocument(bytes)
  expect(original.body?.docxRawBlocks).toEqual([{ startIndex: 0, xml: drawing }])
  const written = await writeDocument(original)
  const entries = unzipSync(written)
  expect(entries['word/media/image1.jpeg']).toEqual(media)
  expect(strFromU8(entries[DOCUMENT_XML] ?? new Uint8Array())).toContain(drawing)
  expect(strFromU8(entries['word/_rels/document.xml.rels'] ?? new Uint8Array())).toContain(relationship)
  expect(strFromU8(entries['[Content_Types].xml'] ?? new Uint8Array())).toContain('ContentType="image/jpeg"')
  const reopened = await readDocument(written)
  expect(reopened.body?.docxRawBlocks).toEqual([{ startIndex: 0, xml: drawing }])
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

it('preserves footnotes and endnotes with their parts and references', async () => {
  const footnoteRun = '<w:r><w:footnoteReference w:id="1"/></w:r>'
  const endnoteRun = '<w:r><w:endnoteReference w:id="2"/></w:r>'
  const footnotes = '<w:footnotes><w:footnote w:id="1"><w:p><w:r><w:t>footnote text</w:t></w:r></w:p></w:footnote></w:footnotes>'
  const endnotes = '<w:endnotes><w:endnote w:id="2"><w:p><w:r><w:t>endnote text</w:t></w:r></w:p></w:endnote></w:endnotes>'
  const footnoteRelationship = '<Relationship Id="rId7"'
    + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"'
    + ' Target="footnotes.xml"/>'
  const endnoteRelationship = '<Relationship Id="rId8"'
    + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"'
    + ' Target="endnotes.xml"/>'
  const contentTypes = '<Types>'
    + '<Override PartName="/word/footnotes.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
    + '<Override PartName="/word/endnotes.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
    + '</Types>'
  const bytes = documentPackage(
    `<w:p><w:r><w:t>body</w:t></w:r>${footnoteRun}${endnoteRun}</w:p>`,
    {
      '[Content_Types].xml': contentTypes,
      'word/footnotes.xml': footnotes,
      'word/endnotes.xml': endnotes,
      'word/_rels/document.xml.rels': `<Relationships>${footnoteRelationship}${endnoteRelationship}</Relationships>`,
    },
  )
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  const original = await readDocument(bytes)
  expect(original.body?.docxRawBlocks).toEqual([
    { startIndex: 4, xml: footnoteRun },
    { startIndex: 4, xml: endnoteRun },
  ])
  const written = await documentRoundTrip(bytes)
  expect(written['word/footnotes.xml']).toBe(footnotes)
  expect(written['word/endnotes.xml']).toBe(endnotes)
  expect(written['[Content_Types].xml']).toContain('/word/footnotes.xml')
  expect(written['[Content_Types].xml']).toContain('/word/endnotes.xml')
  expect(written['word/_rels/document.xml.rels']).toContain(footnoteRelationship)
  expect(written['word/_rels/document.xml.rels']).toContain(endnoteRelationship)
  expect(written[DOCUMENT_XML]).toContain(footnoteRun)
  expect(written[DOCUMENT_XML]).toContain(endnoteRun)
  const reopened = await readDocument(await writeDocument(original))
  expect(reopened.body?.docxRawBlocks).toEqual(original.body?.docxRawBlocks)
})

it('preserves a linked image and its external target', async () => {
  const drawing = '<w:r><w:drawing><wp:inline xmlns:wp="urn:probe"><a:graphic'
    + ' xmlns:a="urn:probe"><a:graphicData><pic:pic xmlns:pic="urn:probe"><pic:blipFill>'
    + '<a:blip r:link="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline>'
    + '</w:drawing></w:r>'
  const relationship = '<Relationship Id="rId7"'
    + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
    + ' Target="https://example.test/linked.png" TargetMode="External"/>'
  const bytes = documentPackage(`<w:p>${drawing}</w:p>`, {
    'word/_rels/document.xml.rels': `<Relationships>${relationship}</Relationships>`,
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  const original = await readDocument(bytes)
  expect(original.body?.docxRawBlocks).toEqual([{ startIndex: 0, xml: drawing }])
  const written = await documentRoundTrip(bytes)
  expect(written[DOCUMENT_XML]).toContain(drawing)
  expect(written['word/_rels/document.xml.rels']).toContain('https://example.test/linked.png')
  expect(written['word/_rels/document.xml.rels']).toContain('TargetMode="External"')
  const reopened = await readDocument(await writeDocument(original))
  expect(reopened.body?.docxRawBlocks).toEqual(original.body?.docxRawBlocks)
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

it('keeps conditional formatting through a write/read round trip', async () => {
  const { source, written, bytes } = await workbookRoundTrip((sheet) => {
    sheet.addConditionalFormatting({
      ref: 'A1:A5',
      rules: [{ type: 'expression', formulae: ['$A1>3'], priority: 1, style: { font: { bold: true } } }],
    })
  })
  expect(detectUnsupportedFeatures(source)).toEqual([])
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  expect(sheetXml(written)).toContain('conditionalFormatting')
  expect(sheetXml(written)).toContain('$A1&gt;3')
  const rules = await conditionalFormattingRulesOf(bytes)
  expect(rules).toHaveLength(1)
  expect(rules[0]?.rule).toMatchObject({ type: 'highlightCell', subType: 'formula', value: '=$A1>3' })
  expect(rules[0]?.ranges).toEqual([{ startRow: 0, startColumn: 0, endRow: 4, endColumn: 0 }])
})

it('keeps data validation through a write/read round trip', async () => {
  const { source, written, bytes } = await workbookRoundTrip((sheet) => {
    const rule = { type: 'list' as const, allowBlank: true, formulae: ['"a,b"'] }
    for (const row of [1, 2, 3]) sheet.getCell(row, 2).dataValidation = rule
  })
  expect(detectUnsupportedFeatures(source)).toEqual([])
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  expect(sheetXml(written)).toContain('dataValidation')
  expect(sheetXml(written)).toContain('sqref="B1:B3"')
  const rules = await dataValidationRulesOf(bytes)
  expect(rules).toHaveLength(1)
  expect(rules[0]).toMatchObject({
    type: 'list',
    allowBlank: true,
    formula1: '["a","b"]',
    ranges: [
      { startRow: 0, startColumn: 1, endRow: 0, endColumn: 1 },
      { startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 },
      { startRow: 2, startColumn: 1, endRow: 2, endColumn: 1 },
    ],
  })
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

it('keeps a cell hyperlink target through a write/read round trip', async () => {
  const { source, written, bytes } = await workbookRoundTrip((sheet) => {
    sheet.getCell('C1').value = { text: 'link cell', hyperlink: 'https://example.test/', tooltip: 'tip' }
  })
  expect(detectUnsupportedFeatures(source)).toEqual([])
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  expect(sheetXml(written)).toContain('<hyperlink')
  const relationships = Object.entries(written)
    .filter(([name]) => name.endsWith('.rels'))
    .map(([, text]) => text)
    .join('')
  expect(relationships).toContain('example.test')
  const link = sheetOf(await readWorkbook(bytes), 'Data').cellData?.['0']?.['2']
  expect(link?.p?.body?.customRanges?.[0]).toMatchObject({
    rangeType: 0,
    properties: { url: 'https://example.test/' },
  })
  expect(link?.p?.body?.dataStream).toBe('link cell\r\n')
})

it('keeps a cell image through a write/read round trip', async () => {
  const { source, written, bytes } = await workbookRoundTrip((sheet, workbook) => {
    const imageId = workbook.addImage({ base64: ONE_PIXEL_PNG, extension: 'png' })
    sheet.addImage(imageId, { tl: { col: 1, row: 1 }, ext: { width: 20, height: 10 } })
  })
  expect(detectUnsupportedFeatures(source)).toEqual([])
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  expect(Object.keys(written).some(name => name.startsWith('xl/media/'))).toBe(true)
  expect(Object.entries(written).find(([name]) => name.startsWith('xl/drawings/drawing'))?.[1])
    .toContain('xdr:pic')
  const drawings = resourceOf(await readWorkbook(bytes), 'SHEET_DRAWING_PLUGIN') as
    Record<string, { data: Record<string, { source: string }>; order: string[] }> | undefined
  const sheetDrawings = drawings?.['sheet-1']
  expect(sheetDrawings?.order).toHaveLength(1)
  expect(Object.values(sheetDrawings?.data ?? {})[0]?.source).toBe(ONE_PIXEL_PNG)
})

it('keeps a worksheet background image through a write/read round trip', async () => {
  const { source, written, bytes } = await workbookRoundTrip((sheet, workbook) => {
    sheet.addBackgroundImage(workbook.addImage({ base64: ONE_PIXEL_PNG, extension: 'png' }))
  })
  expect(detectUnsupportedFeatures(source)).toEqual([])
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
  expect(sheetXml(written)).toContain('<picture')
  expect(sheetOf(await readWorkbook(bytes), 'Data').backgroundImage?.source).toBe(ONE_PIXEL_PNG)
})
