/**
 * DOCX table structure: merged cells, cell formatting, row properties, table
 * layout, and nested tables mapped into the editor model and kept by a
 * write/read round-trip.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { expect, it } from 'vitest'
import {
  DashStyleType,
  NumberUnitType,
  TableAlignmentType,
  TableLayoutType,
  TableRowHeightRule,
  TableSizeType,
  VerticalAlignmentType,
  validateDocumentStructure,
  type IDocumentData,
  type ITable,
} from '@univerjs/core'
import { readDocument, writeDocument } from '../src/provider/docx.ts'

const DOCUMENT_XML = 'word/document.xml'
const DOCUMENT_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

function packageOf(entries: Readonly<Record<string, string>>): Uint8Array {
  const files: Record<string, Uint8Array> = {}
  for (const [name, content] of Object.entries(entries)) files[name] = strToU8(content)
  return zipSync(files)
}

function documentPackage(body: string): Uint8Array {
  return packageOf({
    '[Content_Types].xml': '<Types/>',
    [DOCUMENT_XML]: `<w:document ${DOCUMENT_NAMESPACE}><w:body>${body}</w:body></w:document>`,
  })
}

function writtenXml(bytes: Uint8Array): string {
  const content = unzipSync(bytes)[DOCUMENT_XML]
  return content === undefined ? '' : strFromU8(content)
}

function tableOf(data: IDocumentData, index = 0): ITable {
  const range = data.body?.tables?.[index]
  if (range === undefined) throw new Error('missing table range')
  const table = data.tableSource?.[range.tableId]
  if (table === undefined) throw new Error('missing table metadata')
  return table
}

it('maps a horizontal grid span and keeps it through a round-trip', async () => {
  const bytes = documentPackage(
    '<w:tbl><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="1200"/></w:tblGrid>'
    + '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>'
    + '<w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc></w:tr>'
    + '<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>'
    + '<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  )
  const original = await readDocument(bytes)
  expect(validateDocumentStructure(original)).toEqual([])
  const table = tableOf(original)
  expect(table.tableColumns).toHaveLength(2)
  expect(table.tableRows[0]?.tableCells[0]?.columnSpan).toBe(2)
  expect(table.tableRows[1]?.tableCells).toHaveLength(2)

  const written = await writeDocument(original)
  expect(writtenXml(written)).toContain('<w:gridSpan w:val="2"/>')
  const reopened = await readDocument(written)
  expect(validateDocumentStructure(reopened)).toEqual([])
  expect(tableOf(reopened).tableRows[0]?.tableCells[0]?.columnSpan).toBe(2)
})

it('maps a vertical merge into a start-cell row span and covered cells', async () => {
  const bytes = documentPackage(
    '<w:tbl><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="1200"/></w:tblGrid>'
    + '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>'
    + '<w:p><w:r><w:t>tall</w:t></w:r></w:p></w:tc>'
    + '<w:tc><w:p><w:r><w:t>next</w:t></w:r></w:p></w:tc></w:tr>'
    + '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>'
    + '<w:tc><w:p><w:r><w:t>two</w:t></w:r></w:p></w:tc></w:tr>'
    + '<w:tr><w:tc><w:tcPr><w:vMerge w:val="continue"/></w:tcPr><w:p/></w:tc>'
    + '<w:tc><w:p><w:r><w:t>three</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  )
  const original = await readDocument(bytes)
  expect(validateDocumentStructure(original)).toEqual([])
  const table = tableOf(original)
  expect(table.tableRows[0]?.tableCells[0]?.rowSpan).toBe(3)
  expect(table.tableRows[1]?.tableCells[0]?.rowSpan).toBe(0)
  expect(table.tableRows[2]?.tableCells[0]?.rowSpan).toBe(0)

  const written = await writeDocument(original)
  const xml = writtenXml(written)
  expect(xml).toContain('<w:vMerge w:val="restart"/>')
  expect(xml).toContain('<w:vMerge/>')
  const reopened = await readDocument(written)
  expect(validateDocumentStructure(reopened)).toEqual([])
  const rewritten = tableOf(reopened)
  expect(rewritten.tableRows[0]?.tableCells[0]?.rowSpan).toBe(3)
  expect(rewritten.tableRows[1]?.tableCells[0]?.rowSpan).toBe(0)
  expect(rewritten.tableRows[2]?.tableCells[0]?.rowSpan).toBe(0)
})

it('maps cell borders, shading, vertical alignment, and margins', async () => {
  const bytes = documentPackage(
    '<w:tbl><w:tr><w:tc><w:tcPr><w:tcBorders>'
    + '<w:top w:val="single" w:sz="12" w:color="FF0000"/>'
    + '<w:left w:val="dashed" w:sz="6" w:color="00FF00"/>'
    + '<w:bottom w:val="none"/>'
    + '<w:right w:val="double" w:sz="24" w:color="0000FF"/>'
    + '</w:tcBorders>'
    + '<w:shd w:val="clear" w:color="auto" w:fill="ABCDEF"/>'
    + '<w:tcMar><w:top w:w="120" w:type="dxa"/><w:start w:w="240" w:type="dxa"/>'
    + '<w:bottom w:w="60" w:type="dxa"/><w:end w:w="300" w:type="dxa"/></w:tcMar>'
    + '<w:vAlign w:val="center"/>'
    + '</w:tcPr><w:p><w:r><w:t>formatted</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  )
  const original = await readDocument(bytes)
  const cell = tableOf(original).tableRows[0]?.tableCells[0]
  expect(cell?.borderTop).toEqual({
    color: { rgb: '#FF0000' },
    width: { v: 2 },
    dashStyle: DashStyleType.SOLID,
  })
  expect(cell?.borderLeft).toEqual({
    color: { rgb: '#00FF00' },
    width: { v: 1 },
    dashStyle: DashStyleType.DASH,
  })
  expect(cell?.borderBottom).toBeUndefined()
  expect(cell?.borderRight).toEqual({
    color: { rgb: '#0000FF' },
    width: { v: 4 },
    dashStyle: DashStyleType.SOLID,
  })
  expect(cell?.backgroundColor).toEqual({ rgb: '#ABCDEF' })
  expect(cell?.vAlign).toBe(VerticalAlignmentType.CENTER)
  expect(cell?.margin).toEqual({
    start: { v: 16 },
    end: { v: 20 },
    top: { v: 8 },
    bottom: { v: 4 },
  })

  const written = await writeDocument(original)
  const xml = writtenXml(written)
  expect(xml).toContain('<w:tcBorders>')
  expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="ABCDEF"/>')
  expect(xml).toContain('<w:tcMar>')
  expect(xml).toContain('<w:vAlign w:val="center"/>')
  const reopened = await readDocument(written)
  const rewritten = tableOf(reopened).tableRows[0]?.tableCells[0]
  expect(rewritten?.borderTop).toEqual(cell?.borderTop)
  expect(rewritten?.borderLeft).toEqual(cell?.borderLeft)
  expect(rewritten?.borderRight).toEqual(cell?.borderRight)
  expect(rewritten?.backgroundColor).toEqual(cell?.backgroundColor)
  expect(rewritten?.vAlign).toBe(cell?.vAlign)
  expect(rewritten?.margin).toEqual(cell?.margin)
})

it('maps table layout and repeating header row properties', async () => {
  const bytes = documentPackage(
    '<w:tbl><w:tblPr>'
    + '<w:tblW w:w="6000" w:type="dxa"/><w:jc w:val="center"/>'
    + '<w:tblInd w:w="300" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
    + '</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
    + '<w:tr><w:trPr><w:tblHeader/><w:cantSplit/>'
    + '<w:trHeight w:val="420" w:hRule="atLeast"/></w:trPr>'
    + '<w:tc><w:p><w:r><w:t>head</w:t></w:r></w:p></w:tc>'
    + '<w:tc><w:p><w:r><w:t>row</w:t></w:r></w:p></w:tc></w:tr>'
    + '<w:tr><w:tc><w:p><w:r><w:t>body</w:t></w:r></w:p></w:tc>'
    + '<w:tc><w:p><w:r><w:t>row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  )
  const original = await readDocument(bytes)
  const table = tableOf(original)
  expect(table.size).toEqual({ type: TableSizeType.SPECIFIED, width: { v: 400 } })
  expect(table.align).toBe(TableAlignmentType.CENTER)
  expect(table.indent).toEqual({ v: 20 })
  expect(table.layout).toBe(TableLayoutType.FIXED)
  const header = table.tableRows[0]
  expect(header?.trHeight).toEqual({ val: { v: 28 }, hRule: TableRowHeightRule.AT_LEAST })
  expect(header?.cantSplit).toBe(1)
  expect(header?.repeatHeaderRow).toBe(1)
  expect(header?.isFirstRow).toBe(1)
  expect(table.tableRows[1]?.isFirstRow).toBeUndefined()

  const written = await writeDocument(original)
  const xml = writtenXml(written)
  expect(xml).toContain('<w:tblW w:w="6000" w:type="dxa"/>')
  expect(xml).toContain('<w:jc w:val="center"/>')
  expect(xml).toContain('<w:tblInd w:w="300" w:type="dxa"/>')
  expect(xml).toContain('<w:tblLayout w:type="fixed"/>')
  expect(xml).toContain('<w:tblHeader/>')
  expect(xml).toContain('<w:cantSplit/>')
  const reopened = await readDocument(written)
  const rewritten = tableOf(reopened)
  expect(rewritten.size).toEqual(table.size)
  expect(rewritten.align).toBe(table.align)
  expect(rewritten.indent).toEqual(table.indent)
  expect(rewritten.layout).toBe(table.layout)
  expect(rewritten.tableRows[0]?.trHeight).toEqual(header?.trHeight)
  expect(rewritten.tableRows[0]?.cantSplit).toBe(header?.cantSplit)
  expect(rewritten.tableRows[0]?.repeatHeaderRow).toBe(header?.repeatHeaderRow)
  expect(rewritten.tableRows[0]?.isFirstRow).toBe(header?.isFirstRow)
})

it('maps percentage width, end alignment, automatic layout, and exact height', async () => {
  const bytes = documentPackage(
    '<w:tbl><w:tblPr><w:tblW w:w="2500" w:type="pct"/><w:jc w:val="end"/>'
    + '<w:tblLayout w:type="autofit"/></w:tblPr>'
    + '<w:tr><w:trPr><w:trHeight w:val="300" w:hRule="exact"/></w:trPr>'
    + '<w:tc><w:tcPr><w:tcBorders><w:top w:val="dotted"/></w:tcBorders>'
    + '<w:vAlign w:val="top"/></w:tcPr><w:p/></w:tc>'
    + '<w:tc><w:tcPr><w:vAlign w:val="bottom"/></w:tcPr><w:p/></w:tc>'
    + '</w:tr></w:tbl>',
  )
  const original = await readDocument(bytes)
  const table = tableOf(original)
  expect(table.size).toEqual({
    type: TableSizeType.SPECIFIED,
    width: { v: 50, u: NumberUnitType.PERCENT },
  })
  expect(table.align).toBe(TableAlignmentType.END)
  expect(table.layout).toBe(TableLayoutType.AUTO_FIT)
  expect(table.tableRows[0]?.trHeight).toEqual({ val: { v: 20 }, hRule: TableRowHeightRule.EXACT })
  expect(table.tableRows[0]?.tableCells[0]?.borderTop).toEqual({
    color: { rgb: '#000000' },
    width: { v: 4 / 6 },
    dashStyle: DashStyleType.DOT,
  })
  expect(table.tableRows[0]?.tableCells[0]?.vAlign).toBe(VerticalAlignmentType.TOP)
  expect(table.tableRows[0]?.tableCells[1]?.vAlign).toBe(VerticalAlignmentType.BOTTOM)

  const written = await writeDocument(original)
  const xml = writtenXml(written)
  expect(xml).toContain('<w:tblW w:w="2500" w:type="pct"/>')
  expect(xml).toContain('<w:jc w:val="end"/>')
  expect(xml).toContain('<w:tblLayout w:type="autofit"/>')
  expect(xml).toContain('w:hRule="exact"')
  const reopened = await readDocument(written)
  const rewritten = tableOf(reopened)
  expect(rewritten.size).toEqual(table.size)
  expect(rewritten.align).toBe(table.align)
  expect(rewritten.layout).toBe(table.layout)
  expect(rewritten.tableRows[0]?.trHeight).toEqual(table.tableRows[0]?.trHeight)
  expect(rewritten.tableRows[0]?.tableCells[0]?.borderTop).toEqual(table.tableRows[0]?.tableCells[0]?.borderTop)
  expect(rewritten.tableRows[0]?.tableCells[0]?.vAlign).toBe(VerticalAlignmentType.TOP)
  expect(rewritten.tableRows[0]?.tableCells[1]?.vAlign).toBe(VerticalAlignmentType.BOTTOM)
})

it('maps a nested table into its own range and writes it back', async () => {
  const bytes = documentPackage(
    '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>outer</w:t></w:r></w:p>'
    + '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    + '</w:tc></w:tr></w:tbl>',
  )
  const original = await readDocument(bytes)
  expect(validateDocumentStructure(original)).toEqual([])
  expect(original.body?.tables).toHaveLength(2)
  expect(original.body?.tables?.[0]?.tableId).toBe('table-1')
  expect(original.body?.tables?.[1]?.tableId).toBe('table-2')
  expect(original.body?.dataStream).toContain('inner')
  expect(tableOf(original, 1).tableRows).toHaveLength(1)

  const written = await writeDocument(original)
  const xml = writtenXml(written)
  expect(xml.match(/<w:tbl>/gu)).toHaveLength(2)
  expect(xml).toContain('inner')
  const reopened = await readDocument(written)
  expect(validateDocumentStructure(reopened)).toEqual([])
  expect(reopened.body?.tables).toHaveLength(2)
  expect(reopened.body?.dataStream).toContain('inner')
  expect(tableOf(reopened, 1).tableRows).toHaveLength(1)
})
