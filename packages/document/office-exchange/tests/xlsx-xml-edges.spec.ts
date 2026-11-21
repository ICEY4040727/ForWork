/** Scanner and parser edges reached through unusual but valid worksheet XML. */
import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { OfficeExchangeError } from '../src/errors.ts'
import { open } from '../src/package/ooxml-package.ts'
import { readWorkbookWithSource, writeWorkbookFromSource } from '../src/provider/xlsx.ts'
import type { OfficeSourceState } from '../src/source/source-state.ts'

const SPREADSHEET = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
const SHEET = 'xl/worksheets/sheet1.xml'

function buildWorkbook(sheet: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?>'
      + `<Types xmlns="${CT}">`
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      + '</Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0"?>'
      + `<Relationships xmlns="${PKG_REL}">`
      + `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`
      + '</Relationships>',
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?>'
      + `<workbook xmlns="${SPREADSHEET}" xmlns:r="${REL}">`
      + '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?>'
      + `<Relationships xmlns="${PKG_REL}">`
      + `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`
      + '</Relationships>',
    ),
    [SHEET]: strToU8(sheet),
  })
}

async function refusalOf(result: Promise<unknown>): Promise<OfficeExchangeError> {
  try {
    await result
  } catch (error) {
    if (error instanceof OfficeExchangeError) return error
    throw error
  }
  throw new Error('Expected the writer to refuse')
}

function byteOffsetOf(text: string, index: number): number {
  return strToU8(text.slice(0, index)).byteLength
}

describe('worksheet XML scanning', () => {
  it('accepts comments, CDATA, declarations, processing instructions, prefixes, and entities', async () => {
    const sheet = '<?xml version="1.0" encoding="UTF-8"?>'
      + '<?processing instruction?>'
      + '<!DOCTYPE worksheet>'
      + `<xsl:worksheet xmlns:xsl="${SPREADSHEET}" xmlns:r="${REL}" xmlns:foo="urn:example:foo">`
      + '<!-- a comment with <angle> brackets -->'
      + '<xsl:dimension ref="A1:F3"/>'
      + '<xsl:sheetData>'
      + '<xsl:row r="1" spans="1:6">'
      + '<xsl:c r="A1"><xsl:v>1</xsl:v></xsl:c>'
      + '<xsl:c r="B1"><![CDATA[<v>2</v>]]><xsl:v>2</xsl:v></xsl:c>'
      + '<xsl:c r="C1"><xsl:is><xsl:t xml:space="preserve">&#x41;&amp;&lt;&gt;&quot;&apos;&#66;</xsl:t></xsl:is></xsl:c>'
      + '<xsl:c r="D1"><xsl:f t="shared" si="0">A1+1</xsl:f><xsl:v>2</xsl:v></xsl:c>'
      + '<xsl:c r="E1"><xsl:v/></xsl:c>'
      + '<xsl:c r="F1"><foo:bar/></xsl:c>'
      + '</xsl:row>'
      + '<xsl:row r="2"><xsl:c r="A2"><xsl:is><xsl:r><xsl:t>Rich</xsl:t></xsl:r></xsl:is></xsl:c></xsl:row>'
      + '<xsl:row r="3"/>'
      + '</xsl:sheetData>'
      + '<xsl:mergeCells count="1"><xsl:mergeCell ref="A1:B1"/></xsl:mergeCells>'
      + '<foo:extension/>'
      + '</xsl:worksheet>'
    const source = buildWorkbook(sheet)
    const data = await readWorkbookWithSource(source, 'xlsx')
    const result = await writeWorkbookFromSource(data, source)

    expect(result.disposition.spliced.names).toContain(SHEET)
  })

  it('skips declaration and processing-instruction text that resembles worksheet markup', async () => {
    const cell = '<c r="A1"><v>1</v></c>'
    const sheet = '<?xml version="1.0"?>'
      + '<?pi value="<c r=\'Z8\'/>"?>'
      + '<!DOCTYPE worksheet PUBLIC "urn:pub>id" "urn:sys<c r=\'Z9\'/>x">'
      + `<worksheet xmlns="${SPREADSHEET}"><sheetData><row r="1">${cell}</row></sheetData></worksheet>`
    const source = buildWorkbook(sheet)
    const data = await readWorkbookWithSource(source, 'xlsx')
    const state = (data as { officeExchange?: { source?: OfficeSourceState } }).officeExchange?.source
    if (state === undefined || state.anchorMode !== 'per-unit') throw new Error('fixture has no per-unit source state')

    expect(state.anchors).toHaveLength(1)
    const cellIndex = sheet.indexOf(cell)
    expect(state.anchors[0]).toMatchObject({
      start: byteOffsetOf(sheet, cellIndex),
      end: byteOffsetOf(sheet, cellIndex + cell.length),
    })
    const result = await writeWorkbookFromSource(data, source)
    expect(result.bytes).toEqual(source)
  })

  it('refuses a worksheet that repeats its dimension', async () => {
    const sheet = '<?xml version="1.0"?>'
      + `<worksheet xmlns="${SPREADSHEET}"><dimension ref="A1"/><dimension ref="A1:B2"/>`
      + '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'
    const source = buildWorkbook(sheet)
    const data = await readWorkbookWithSource(source, 'xlsx')
    const error = await refusalOf(writeWorkbookFromSource(data, source))

    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('repeats its dimension')
  })

  it('refuses alternate content the scanner cannot resolve', async () => {
    const clean = buildWorkbook('<?xml version="1.0"?>'
      + `<worksheet xmlns="${SPREADSHEET}">`
      + '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>')
    const sheet = '<?xml version="1.0"?>'
      + `<worksheet xmlns="${SPREADSHEET}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">`
      + '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
      + '<mc:AlternateContent><mc:Choice Requires="x14"><extLst xmlns:x14="urn:x14"/></mc:Choice></mc:AlternateContent>'
      + '</worksheet>'
    const source = buildWorkbook(sheet)
    const data = await readWorkbookWithSource(clean, 'xlsx')
    const state = (data as { officeExchange?: { source?: OfficeSourceState } }).officeExchange?.source
    if (state === undefined) throw new Error('fixture has no source state')
    const archive = open(source, { maxEntries: 100 })
    ;(data as { officeExchange?: unknown }).officeExchange = {
      source: {
        ...state,
        packageDigest: archive.source.digest,
        byteLength: archive.source.bytes.byteLength,
      },
    }
    const error = await refusalOf(writeWorkbookFromSource(data, source))

    expect(error.code).toBe('conversion-failed')
    expect(error.message).toContain('ambiguous-span')
  })

  it('accepts comments, CDATA, processing instructions, and metadata inside cells', async () => {
    const sheet = '<?xml version="1.0"?>'
      + `<worksheet xmlns="${SPREADSHEET}"><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><sheetData><row r="1">`
      + '<c r="A1"><!-- comment --><v>1</v></c>'
      + '<c r="B1"><![CDATA[ignored]]><v>2</v></c>'
      + '<c r="C1"><?pi value?><v>3</v></c>'
      + '<c r="D1"><is><r><t>a</t></r><t>b</t></is></c>'
      + '<c r="E1" cm="1"><v>5</v></c>'
      + '<c r="F1"><is/></c>'
      + '</row></sheetData></worksheet>'
    const source = buildWorkbook(sheet)
    const data = await readWorkbookWithSource(source, 'xlsx')
    const result = await writeWorkbookFromSource(data, source)
    expect(result.bytes.byteLength).toBeGreaterThan(0)
  })
})
