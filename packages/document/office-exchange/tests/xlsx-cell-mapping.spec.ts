/** Cell value, formula, date, boolean, error, and hyperlink mapping. */
import ExcelJS from 'exceljs'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { DateSystem } from '@univerjs/core'
import type { ICellData, IWorkbookData } from '@univerjs/core'
import { readWorkbook, writeWorkbook } from '../src/provider/xlsx.ts'

async function baseWorkbook(): Promise<IWorkbookData> {
  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('Data')
  return readWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
}

function cell(data: IWorkbookData, row: number, column: number): ICellData {
  const sheet = data.sheets[data.sheetOrder[0] ?? '']
  if (sheet?.cellData?.[row]?.[column] === undefined) throw new Error(`no cell ${String(row)}:${String(column)}`)
  return sheet.cellData[row][column]
}

describe('reading cell values and hyperlinks', () => {
  it('maps booleans, errors, rich text, dates, and formulas with cached results', async () => {
    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('Data')
    sheet.getCell('A1').value = true
    sheet.getCell('A2').value = 42
    sheet.getCell('A3').value = 'text'
    sheet.getCell('A4').value = { richText: [{ text: 'a' }, { text: 'b' }] }
    sheet.getCell('A5').value = { error: '#DIV/0!' }
    sheet.getCell('A6').value = new Date('2024-01-02T00:00:00.000Z')
    sheet.getCell('A7').value = { formula: 'A1+1', result: 2 }
    sheet.getCell('A8').value = { formula: 'A1+2' }
    sheet.getCell('A9').value = { text: 'Link', hyperlink: 'https://example.invalid/a', tooltip: 'tip' }
    sheet.getCell('A10').value = { text: 'Link', hyperlink: 'https://example.invalid/b' }
    const textlessHyperlink = { text: '', hyperlink: 'https://example.invalid/c' }
    Reflect.deleteProperty(textlessHyperlink, 'text')
    sheet.getCell('A11').value = textlessHyperlink

    const data = await readWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))

    expect(cell(data, 0, 0)).toMatchObject({ v: true, t: 3 })
    expect(cell(data, 1, 0)).toMatchObject({ v: 42, t: 2 })
    expect(cell(data, 2, 0)).toMatchObject({ v: 'text', t: 1 })
    expect(cell(data, 3, 0)).toMatchObject({ v: 'ab', t: 1 })
    expect(cell(data, 4, 0)).toMatchObject({ v: '#DIV/0!', t: 1 })
    expect(cell(data, 5, 0)).toMatchObject({ t: 2 })
    expect(cell(data, 6, 0)).toMatchObject({ f: '=A1+1', v: 2, t: 2 })
    expect(cell(data, 7, 0)).toMatchObject({ f: '=A1+2' })
    expect(cell(data, 8, 0).p?.body?.dataStream).toBe('Link\r\n')
    expect(cell(data, 8, 0).p?.body?.customRanges?.[0]?.properties).toEqual({
      url: 'https://example.invalid/a',
    })
    expect(cell(data, 9, 0).p?.body?.customRanges?.[0]?.properties).toEqual({
      url: 'https://example.invalid/b',
    })
  })

  it('round-trips dates through both date systems', async () => {
    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('Data')
    sheet.getCell('A1').value = new Date('2024-01-02T00:00:00.000Z')
    const data = await readWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
    data.dateSystem = DateSystem.Date1904
    const reread = await readWorkbook(new Uint8Array(await writeWorkbook(data)))

    expect(cell(reread, 0, 0).t).toBe(2)
  })
})

function documentOf(stream: string, url?: string, tooltip?: string, rangeType = 0): NonNullable<ICellData['p']> {
  const customRanges = url === undefined
    ? []
    : [{
      startIndex: 0,
      endIndex: Math.max(0, stream.length - 1),
      rangeId: 'range',
      rangeType,
      properties: tooltip === undefined ? { url } : { url, tooltip },
    }]
  return {
    id: 'cell-doc',
    documentStyle: {},
    body: {
      dataStream: stream,
      ...(customRanges.length === 0 ? {} : { customRanges }),
    },
  }
}

describe('writing cell values and hyperlinks', () => {
  it('writes hyperlink documents with and without text and tooltips', async () => {
    const data = await baseWorkbook()
    const sheet = data.sheets[data.sheetOrder[0] ?? '']
    if (sheet === undefined) throw new Error('fixture has no sheet')
    sheet.cellData = {
      0: {
        0: { p: documentOf('Text\r\n', 'https://example.invalid/a') },
        1: { p: documentOf('', 'https://example.invalid/b') },
        2: { p: documentOf('Plain', 'https://example.invalid/c') },
        3: { p: documentOf('Tip\r\n', 'https://example.invalid/d', 'shown') },
        4: { p: documentOf('EmptyTip\r\n', 'https://example.invalid/e', '') },
        5: { p: documentOf('NoRange\r\n') },
        6: { p: documentOf('OtherRange\r\n', 'https://example.invalid/f', undefined, 1) },
      },
    }
    const bytes = new Uint8Array(await writeWorkbook(data))
    const sheetXml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array(0))
    const reread = await readWorkbook(bytes)

    expect(sheetXml).toContain('tooltip="shown"')
    expect(sheetXml).not.toContain('tooltip=""')
    expect(cell(reread, 0, 0)).toMatchObject({
      v: 'Text',
      p: { body: { customRanges: [{ properties: { url: 'https://example.invalid/a' } }] } },
    })
    expect(cell(reread, 0, 1).v).toBe('https://example.invalid/b')
    expect(cell(reread, 0, 2).v).toBe('Plain')
    expect(cell(reread, 0, 3).p?.body?.customRanges?.[0]?.properties).toEqual({
      url: 'https://example.invalid/d',
    })
    expect(cell(reread, 0, 4).p?.body?.customRanges?.[0]?.properties).toEqual({
      url: 'https://example.invalid/e',
    })
    expect(cell(reread, 0, 5).v).toBe('NoRange')
    expect(cell(reread, 0, 6).v).toBe('OtherRange')
  })

  it('ignores a hyperlink document whose URL is missing or empty', async () => {
    const data = await baseWorkbook()
    const sheet = data.sheets[data.sheetOrder[0] ?? '']
    if (sheet === undefined) throw new Error('fixture has no sheet')
    sheet.cellData = {
      0: {
        0: { p: documentOf('Missing', undefined) },
        1: { p: documentOf('Empty\r\n', '') },
      },
    }
    const reread = await readWorkbook(new Uint8Array(await writeWorkbook(data)))

    expect(cell(reread, 0, 0).v).toBe('Missing')
    expect(cell(reread, 0, 1).v).toBe('Empty')
  })
})
