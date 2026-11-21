/** XLSX conditional-formatting, data-validation, and drawing mapping. */
import ExcelJS, { type Anchor } from 'exceljs'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import type {
  ICellData,
  IWorksheetData,
  IWorkbookData,
} from '@univerjs/core'
import { DateSystem, ImageSourceType, LocaleType } from '@univerjs/core'
import { readWorkbook, writeWorkbook } from '../src/provider/xlsx.ts'

const SPREADSHEET = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${SPREADSHEET}">
  <numFmts count="0"/>
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <dxfs count="1"><dxf><font><b/><color rgb="FF112233"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFAABBCC"/></patternFill></fill></dxf></dxfs>
</styleSheet>`

interface PackageFixture {
  readonly sheetData?: string
  readonly dimension?: string
  readonly extra?: string
  readonly styles?: string
  readonly date1904?: boolean
}

function buildPackage(fixture: PackageFixture): Uint8Array {
  const sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<worksheet xmlns="${SPREADSHEET}" xmlns:r="${REL}">`
    + (fixture.dimension === undefined ? '' : `<dimension ref="${fixture.dimension}"/>`)
    + `<sheetData>${fixture.sheetData ?? ''}</sheetData>`
    + (fixture.extra ?? '')
    + '</worksheet>'
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?>'
      + `<Types xmlns="${CT}">`
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
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
      + (fixture.date1904 === true ? '<workbookPr date1904="1"/>' : '')
      + '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?>'
      + `<Relationships xmlns="${PKG_REL}">`
      + `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`
      + `<Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/>`
      + '</Relationships>',
    ),
    'xl/worksheets/sheet1.xml': strToU8(sheet),
    'xl/styles.xml': strToU8(fixture.styles ?? STYLES),
  })
}

interface ConditionalRuleView {
  readonly type: string
  readonly subType?: string
  readonly operator?: string
  readonly value?: unknown
  readonly style?: unknown
  readonly isBottom?: boolean
  readonly isPercent?: boolean
  readonly config?: unknown
  readonly isShowValue?: boolean
}

interface ConditionalEntryView {
  readonly cfId: string
  readonly ranges: readonly unknown[]
  readonly stopIfTrue: boolean
  readonly rule: ConditionalRuleView
}

interface ValidationView {
  readonly uid: string
  readonly type: string
  readonly operator?: string
  readonly formula1?: string
  readonly formula2?: string
  readonly allowBlank?: boolean
  readonly showErrorMessage?: boolean
  readonly errorStyle?: number
  readonly error?: string
  readonly errorTitle?: string
  readonly showInputMessage?: boolean
  readonly prompt?: string
  readonly promptTitle?: string
  readonly ranges: readonly {
    readonly startRow: number
    readonly startColumn: number
    readonly endRow: number
    readonly endColumn: number
  }[]
}

function resourceData(data: IWorkbookData, name: string): string | undefined {
  return data.resources?.find(resource => resource.name === name)?.data
}

function conditionalFormattingOf(data: IWorkbookData): Record<string, ConditionalEntryView[]> {
  const text = resourceData(data, 'SHEET_CONDITIONAL_FORMATTING_PLUGIN')
  if (text === undefined) throw new Error('workbook has no conditional formatting resource')
  return JSON.parse(text) as Record<string, ConditionalEntryView[]>
}

function validationsOf(data: IWorkbookData): Record<string, ValidationView[]> {
  const text = resourceData(data, 'SHEET_DATA_VALIDATION_PLUGIN')
  if (text === undefined) throw new Error('workbook has no data validation resource')
  return JSON.parse(text) as Record<string, ValidationView[]>
}

const CONDITIONAL_FORMATTING_XML = `
<conditionalFormatting sqref="A1:A3">
  <cfRule type="expression" dxfId="0" priority="5"><formula>A1&gt;0</formula></cfRule>
  <cfRule type="cellIs" operator="between" priority="2"><formula>1</formula><formula>3</formula></cfRule>
  <cfRule type="cellIs" operator="greaterThan" priority="3"><formula>B1</formula></cfRule>
  <cfRule type="cellIs" operator="between" priority="4"><formula>1</formula></cfRule>
  <cfRule type="cellIs" priority="1"><formula>abc</formula></cfRule>
  <cfRule type="containsText" operator="containsText" priority="6"><formula>NOT(ISERROR(SEARCH("foo",A1)))</formula></cfRule>
  <cfRule type="containsText" priority="7"/>
  <cfRule type="containsText" operator="containsText" priority="8"><formula>AND(A1,B1)</formula></cfRule>
  <cfRule type="beginsWith" operator="beginsWith" priority="9"><formula>LEFT(A1,LEN("bar"))="bar"</formula></cfRule>
  <cfRule type="beginsWith" operator="beginsWith" priority="10"/>
  <cfRule type="endsWith" operator="endsWith" priority="11"><formula>RIGHT(A1,LEN("baz"))="baz"</formula></cfRule>
  <cfRule type="notContainsText" operator="notContains" priority="12"><formula>ISERROR(SEARCH("q",A1))</formula></cfRule>
  <cfRule type="uniqueValues" priority="13"/>
  <cfRule type="duplicateValues" priority="14"/>
  <cfRule type="top10" priority="15"/>
  <cfRule type="top10" rank="3" percent="1" bottom="1" priority="16"/>
  <cfRule type="aboveAverage" priority="17"/>
  <cfRule type="aboveAverage" aboveAverage="0" priority="18"/>
  <cfRule type="timePeriod" timePeriod="lastWeek" priority="19"><formula>x</formula></cfRule>
  <cfRule type="timePeriod" timePeriod="weird" priority="20"><formula>z</formula></cfRule>
  <cfRule type="timePeriod" timePeriod="weird" priority="21"/>
  <cfRule type="timePeriod" priority="22"><formula>w</formula></cfRule>
  <cfRule type="colorScale" priority="23"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule>
  <cfRule type="colorScale" priority="24"><colorScale><cfvo type="autoMin"/><cfvo type="autoMax"/><cfvo type="num"/><cfvo type="percent" val="25"/><cfvo type="formula" val="5"/><cfvo type="formula"/><cfvo type="bogus"/></colorScale></cfRule>
  <cfRule type="dataBar" priority="25"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>
  <cfRule type="dataBar" priority="26"><dataBar/></cfRule>
  <cfRule type="iconSet" priority="27"><iconSet iconSet="4Rating" showValue="0"><cfvo type="percent" val="0"/><cfvo type="percent" val="25"/><cfvo type="percent" val="50"/><cfvo type="percent" val="75"/></iconSet></cfRule>
  <cfRule type="iconSet" priority="28"><iconSet iconSet="Bogus"><cfvo type="min"/></iconSet></cfRule>
  <cfRule type="iconSet" priority="29"><iconSet/></cfRule>
  <cfRule type="expression" priority="0"><formula>B1&gt;1</formula></cfRule>
  <cfRule type="expression"><formula>=D1&gt;0</formula></cfRule>
  <cfRule type="expression"/>
  <cfRule type="cellIs" operator="between"/>
  <cfRule type="cellIs" operator="between"><formula>B1</formula><formula>C1</formula></cfRule>
  <cfRule type="cellIs" operator="notBetween"><formula>B1</formula><formula>C1</formula></cfRule>
  <cfRule type="cellIs" operator="equal"><formula>7</formula></cfRule>
  <cfRule type="beginsWith" operator="beginsWith"><formula>AND(D1,E1)</formula></cfRule>
  <cfRule type="bogusRule" priority="30"/>
</conditionalFormatting>
<conditionalFormatting sqref="D1:D2 E1:E2"><cfRule type="cellIs" operator="greaterThan"><formula>B1</formula></cfRule></conditionalFormatting>
<conditionalFormatting sqref="bogus"><cfRule type="expression" priority="31"><formula>1</formula></cfRule></conditionalFormatting>
<conditionalFormatting sqref="C1">
  <cfRule type="containsBlanks" priority="32"><formula>LEN(TRIM(C1))=0</formula></cfRule>
  <cfRule type="notContainsBlanks" priority="33"/>
  <cfRule type="containsErrors" priority="34"><formula>ISERROR(C1)</formula></cfRule>
  <cfRule type="notContainsErrors" priority="35"><formula>NOT(ISERROR(C1))</formula></cfRule>
</conditionalFormatting>`

describe('reading conditional formatting', () => {
  it('maps every supported rule type in priority order and drops unmapped rules', async () => {
    const workbook = await readWorkbook(buildPackage({
      dimension: 'A1:C3',
      sheetData: '<row r="1"><c r="A1"><v>5</v></c></row>',
      extra: CONDITIONAL_FORMATTING_XML,
    }))
    const entries = conditionalFormattingOf(workbook)['sheet-1']
    if (entries === undefined) throw new Error('workbook has no rules for sheet-1')
    const rules = entries.map(entry => entry.rule)

    expect(rules.length).toBeGreaterThan(25)
    expect(rules.find(rule => rule.subType === 'formula' && rule.value === '=A1>0')?.style)
      .toMatchObject({ bl: 1, cl: { rgb: '#112233' }, bg: { rgb: '#AABBCC' } })
    expect(rules.find(rule => rule.subType === 'number' && Array.isArray(rule.value)))
      .toMatchObject({ operator: 'between', value: [1, 3] })
    expect(rules.find(rule => rule.subType === 'formula' && rule.value === '=A1=1')).toBeDefined()
    expect(rules.find(rule => rule.subType === 'formula' && rule.value === '=A1>B1')).toBeDefined()
    expect(rules.find(rule => rule.subType === 'text' && rule.value === 'foo'))
      .toMatchObject({ operator: 'containsText' })
    expect(rules.find(rule => rule.subType === 'text' && rule.operator === 'containsText' && rule.value === ''))
      .toBeDefined()
    expect(rules.find(rule => rule.subType === 'formula' && rule.value === '=AND(A1,B1)')).toBeDefined()
    expect(rules.find(rule => rule.subType === 'text' && rule.value === 'bar')).toBeDefined()
    expect(rules.find(rule => rule.subType === 'text' && rule.value === 'baz')).toBeDefined()
    expect(rules.find(rule => rule.subType === 'text' && rule.operator === 'notContainsText')).toBeDefined()
    expect(rules.find(rule => rule.subType === 'uniqueValues')).toBeDefined()
    expect(rules.find(rule => rule.subType === 'duplicateValues')).toBeDefined()
    expect(rules.find(rule => rule.subType === 'rank' && rule.value === 10))
      .toMatchObject({ isBottom: false, isPercent: false })
    expect(rules.find(rule => rule.subType === 'rank' && rule.value === 3))
      .toMatchObject({ isBottom: true, isPercent: true })
    expect(rules.find(rule => rule.subType === 'average' && rule.operator === 'greaterThan')).toBeDefined()
    expect(rules.find(rule => rule.subType === 'average' && rule.operator === 'lessThan')).toBeDefined()
    expect(rules.find(rule => rule.subType === 'timePeriod' && rule.operator === 'lastWeek')).toBeDefined()
    expect(rules.find(rule => rule.subType === 'formula' && rule.value === '=z')).toBeDefined()
    expect(rules.find(rule => rule.subType === 'formula' && rule.value === '=w')).toBeDefined()
    const colorScales = rules.filter(rule => rule.type === 'colorScale')
    expect(colorScales).toHaveLength(2)
    expect(colorScales[0]?.config).toEqual([
      { index: 0, color: '#FF0000', value: { type: 'min' } },
      { index: 1, color: '#00FF00', value: { type: 'percentile', value: 50 } },
      { index: 2, color: '#FF0000', value: { type: 'max' } },
    ])
    expect(colorScales[1]?.config).toEqual([
      { index: 0, color: '#000000', value: { type: 'min' } },
      { index: 1, color: '#000000', value: { type: 'max' } },
      { index: 2, color: '#000000', value: { type: 'num', value: 0 } },
      { index: 3, color: '#000000', value: { type: 'percent', value: 25 } },
      { index: 4, color: '#000000', value: { type: 'formula', value: 5 } },
      { index: 5, color: '#000000', value: { type: 'formula', value: 0 } },
      { index: 6, color: '#000000', value: { type: 'min' } },
    ])
    const dataBars = rules.filter(rule => rule.type === 'dataBar')
    expect(dataBars).toHaveLength(2)
    expect(dataBars[0]).toMatchObject({
      isShowValue: true,
      config: {
        min: { type: 'min' },
        max: { type: 'max' },
        isGradient: true,
        positiveColor: '#638EC6',
      },
    })
    expect(dataBars[1]?.config).toMatchObject({ min: { type: 'min' }, max: { type: 'min' } })
    const iconSets = rules.filter(rule => rule.type === 'iconSet')
    expect(iconSets[0]).toMatchObject({ isShowValue: false })
    expect(iconSets[1]?.config).toContainEqual({
      operator: 'greaterThanOrEqual',
      value: { type: 'min' },
      iconType: 'EMPTY_ICON_TYPE',
      iconId: '0',
    })
    expect(iconSets[2]?.config).toEqual([])
    const blanks = rules.filter(rule => rule.operator === 'notContainsBlanks')
    expect(blanks).toHaveLength(1)
    expect(rules.some(rule => rule.subType === 'formula' && String(rule.value).includes('LEN(TRIM('))).toBe(true)
    expect(rules.some(rule => rule.subType === 'formula' && String(rule.value).includes('ISERROR(C1)'))).toBe(true)
    expect(rules[0]?.subType).toBe('formula')
  })
})

const DATA_VALIDATION_XML = `
<dataValidations count="12">
  <dataValidation type="list" allowBlank="1" showErrorMessage="1" errorStyle="stop" error="bad" errorTitle="Bad" showInputMessage="1" prompt="pick" promptTitle="Pick" sqref="A1"><formula1>"a,b"</formula1></dataValidation>
  <dataValidation type="list" errorStyle="information" sqref="A2"><formula1>A1:A3</formula1></dataValidation>
  <dataValidation type="list" errorStyle="warning" sqref="A3"><formula1>5</formula1></dataValidation>
  <dataValidation type="date" operator="between" sqref="B1"><formula1>45000</formula1><formula2>DATE(2024,1,1)</formula2></dataValidation>
  <dataValidation type="whole" operator="greaterThan" sqref="C1"><formula1>5</formula1></dataValidation>
  <dataValidation type="decimal" operator="lessThan" sqref="D1"><formula1>1.5</formula1><formula2>2.5</formula2></dataValidation>
  <dataValidation type="textLength" operator="equal" sqref="E1"><formula1>abc</formula1></dataValidation>
  <dataValidation type="custom" sqref="F1"><formula1>A1&gt;0</formula1></dataValidation>
  <dataValidation type="any" sqref="G1"/>
  <dataValidation type="any" sqref="H1"><formula1>foo</formula1></dataValidation>
  <dataValidation type="any" sqref="I1"><formula1>$A$1</formula1></dataValidation>
  <dataValidation type="any" sqref="J1"><formula1>A1</formula1></dataValidation>
  <dataValidation type="list" sqref="bogus"><formula1>"x"</formula1></dataValidation>
</dataValidations>`

describe('reading data validations', () => {
  it('maps criteria, options, and grouped ranges', async () => {
    const workbook = await readWorkbook(buildPackage({
      dimension: 'A1:J1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      extra: DATA_VALIDATION_XML,
    }))
    const rules = validationsOf(workbook)['sheet-1']
    if (rules === undefined) throw new Error('workbook has no validations for sheet-1')

    expect(rules).toHaveLength(12)
    expect(rules[0]).toMatchObject({
      type: 'list',
      formula1: '["a","b"]',
      allowBlank: true,
      showErrorMessage: true,
      errorStyle: 1,
      error: 'bad',
      errorTitle: 'Bad',
      showInputMessage: true,
      prompt: 'pick',
      promptTitle: 'Pick',
    })
    expect(rules[1]).toMatchObject({ type: 'list', formula1: '=A1:A3', errorStyle: 0 })
    expect(rules[2]).toMatchObject({ type: 'list', formula1: '=5', errorStyle: 2 })
    expect(rules[3]).toMatchObject({ type: 'date', operator: 'between', formula1: '45000', formula2: 'NaN' })
    expect(rules[4]).toMatchObject({ type: 'whole', operator: 'greaterThan', formula1: '5' })
    expect(rules[5]).toMatchObject({ type: 'decimal', formula1: '1.5', formula2: '2.5' })
    expect(rules[6]).toMatchObject({ type: 'textLength', formula1: 'NaN' })
    expect(rules[7]).toMatchObject({ type: 'custom', formula1: '=A1>0' })
    expect(rules[8]).toMatchObject({ type: 'any' })
    expect(rules[9]).toMatchObject({ type: 'any', formula1: 'foo' })
    expect(rules[10]).toMatchObject({ type: 'any', formula1: '=$A$1' })
    expect(rules[11]).toMatchObject({ type: 'any', formula1: '=A1' })
    expect(rules.every(rule => rule.ranges.length > 0)).toBe(true)
  })

  it('converts date criteria through the 1904 serial system', async () => {
    const workbook = await readWorkbook(buildPackage({
      dimension: 'A1',
      sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>',
      date1904: true,
      extra: '<dataValidations count="1">'
        + '<dataValidation type="date" sqref="A1"><formula1>1</formula1></dataValidation>'
        + '</dataValidations>',
    }))

    expect(validationsOf(workbook)['sheet-1']?.[0]?.formula1).toBe('-1461')
  })
})

interface WorkbookResources {
  readonly conditionalFormatting?: Record<string, readonly Record<string, unknown>[]>
  readonly dataValidation?: Record<string, readonly Record<string, unknown>[]>
  readonly drawings?: Record<string, {
    readonly data: Record<string, Record<string, unknown>>
    readonly order: readonly string[]
  }>
  readonly styles?: Record<string, Record<string, unknown>>
  readonly backgroundImages?: Record<string, { readonly source: string; readonly imageSourceType: string }>
}

function snapshotWith(resources: WorkbookResources): IWorkbookData {
  const sheetResources = resources
  const entries = [
    ...(sheetResources.conditionalFormatting === undefined
      ? []
      : [{
        name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN',
        data: JSON.stringify(sheetResources.conditionalFormatting),
      }]),
    ...(sheetResources.dataValidation === undefined
      ? []
      : [{
        name: 'SHEET_DATA_VALIDATION_PLUGIN',
        data: JSON.stringify(sheetResources.dataValidation),
      }]),
    ...(sheetResources.drawings === undefined
      ? []
      : [{
        name: 'SHEET_DRAWING_PLUGIN',
        data: JSON.stringify(sheetResources.drawings),
      }]),
  ]
  const cells: Record<number, Record<number, ICellData>> = { 0: { 0: { v: 'x' } } }
  const sheet: Partial<IWorksheetData> = {
    id: 'sheet-1',
    name: 'Data',
    cellData: cells,
    ...(sheetResources.backgroundImages?.['sheet-1'] === undefined
      ? {}
      : {
        backgroundImage: {
          source: sheetResources.backgroundImages['sheet-1'].source,
          imageSourceType: sheetResources.backgroundImages['sheet-1'].imageSourceType as ImageSourceType,
        },
      }),
  }
  return {
    id: 'resource-workbook',
    name: 'Resource workbook',
    appVersion: '1.4.0',
    locale: LocaleType.EN_US,
    dateSystem: DateSystem.Date1900,
    styles: sheetResources.styles ?? {},
    sheetOrder: ['sheet-1'],
    sheets: { 'sheet-1': sheet },
    resources: entries,
  }
}

function highlightRule(overrides: Record<string, unknown>): Record<string, unknown> {
  return {
    cfId: 'cf',
    ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }],
    stopIfTrue: false,
    rule: {
      type: 'highlightCell',
      subType: 'formula',
      value: '=A1>0',
      style: {},
      ...overrides,
    },
  }
}

async function writtenSheetXml(workbook: IWorkbookData): Promise<string> {
  const bytes = await writeWorkbook(workbook)
  return strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array(0))
}

describe('writing conditional formatting', () => {
  it('writes every highlight rule variant', async () => {
    const rules = [
      highlightRule({}),
      highlightRule({ subType: 'number', operator: 'between', value: [1, 3] }),
      highlightRule({ subType: 'number', operator: 'greaterThan', value: 5 }),
      highlightRule({ subType: 'text', operator: 'equal', value: 'a"b' }),
      highlightRule({ subType: 'text', operator: 'notEqual', value: 'x' }),
      highlightRule({ subType: 'text', operator: 'beginsWith', value: 'a' }),
      highlightRule({ subType: 'text', operator: 'endsWith', value: 'a' }),
      highlightRule({ subType: 'text', operator: 'notContainsText', value: 'a' }),
      highlightRule({ subType: 'text', operator: 'containsBlanks', value: 'a' }),
      highlightRule({ subType: 'text', operator: 'notContainsBlanks', value: 'a' }),
      highlightRule({ subType: 'text', operator: 'containsErrors', value: 'a' }),
      highlightRule({ subType: 'text', operator: 'notContainsErrors', value: 'a' }),
      highlightRule({ subType: 'text', operator: 'containsText', value: 'a' }),
      highlightRule({ subType: 'text', operator: 'unknownText', value: 'a' }),
      highlightRule({ subType: 'rank', isBottom: true, isPercent: true, value: 3 }),
      highlightRule({ subType: 'average', operator: 'lessThan' }),
      highlightRule({ subType: 'average', operator: 'unknown' }),
      highlightRule({ subType: 'uniqueValues' }),
      highlightRule({ subType: 'duplicateValues' }),
      highlightRule({ subType: 'timePeriod', operator: 'lastWeek' }),
    ]
    rules.push({
      cfId: 'cf-unique-empty',
      ranges: [],
      stopIfTrue: false,
      rule: { type: 'highlightCell', subType: 'uniqueValues', style: {} },
    })
    const xml = await writtenSheetXml(snapshotWith({
      conditionalFormatting: { 'sheet-1': rules },
    }))

    expect(xml).toContain('type="expression"')
    expect(xml).toContain('type="cellIs"')
    expect(xml).toContain('type="containsText"')
    expect(xml).toContain('type="top10"')
    expect(xml).toContain('timePeriod="lastWeek"')
  })

  it('writes color scales, data bars, and icon sets', async () => {
    const rules = [
      {
        cfId: 'cf-scale',
        ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }],
        stopIfTrue: false,
        rule: {
          type: 'colorScale',
          config: [
            { index: 0, color: '#112233', value: { type: 'min' } },
            { index: 1, color: 'not-a-color', value: { type: 'num', value: 5 } },
            { index: 2, color: '#445566', value: { type: 'max' } },
            { index: 3, color: '#445566', value: { type: 'percentile', value: '7' } },
            { index: 4, color: '#445566', value: { type: 'formula', value: undefined } },
            { index: 5, color: '#445566', value: { type: 'num', value: undefined } },
            { index: 6, color: '#445566', value: { type: 'formula', value: 'not-a-number' } },
          ],
        },
      },
      {
        cfId: 'cf-bar',
        ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }],
        stopIfTrue: false,
        rule: {
          type: 'dataBar',
          isShowValue: false,
          config: {
            min: { type: 'min' },
            max: { type: 'max' },
            isGradient: false,
            positiveColor: '#112233',
            nativeColor: '#112233',
          },
        },
      },
      {
        cfId: 'cf-bar-default',
        ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }],
        stopIfTrue: false,
        rule: {
          type: 'dataBar',
          isShowValue: true,
          config: {
            min: { type: 'num', value: 1 },
            max: { type: 'num', value: 2 },
            isGradient: true,
            positiveColor: 'wrong',
            nativeColor: 'wrong',
          },
        },
      },
      {
        cfId: 'cf-icons',
        ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }],
        stopIfTrue: false,
        rule: {
          type: 'iconSet',
          isShowValue: false,
          config: [
            { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 0 }, iconType: '3Arrows', iconId: '0' },
          ],
        },
      },
      {
        cfId: 'cf-icons-empty',
        ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }],
        stopIfTrue: false,
        rule: {
          type: 'iconSet',
          isShowValue: true,
          config: [
            { operator: 'greaterThanOrEqual', value: { type: 'min' }, iconType: 'EMPTY_ICON_TYPE', iconId: '0' },
          ],
        },
      },
      {
        cfId: 'cf-icons-none',
        ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }],
        stopIfTrue: false,
        rule: { type: 'iconSet', isShowValue: true, config: [] },
      },
    ]
    const xml = await writtenSheetXml(snapshotWith({
      conditionalFormatting: { 'sheet-1': rules },
    }))

    expect(xml).toContain('type="colorScale"')
    expect(xml).toContain('type="dataBar"')
    expect(xml).toContain('type="iconSet"')
  })

  it('refuses an unrepresentable time-period operator', async () => {
    await expect(writeWorkbook(snapshotWith({
      conditionalFormatting: {
        'sheet-1': [highlightRule({ subType: 'timePeriod', operator: 'notAPeriod' })],
      },
    }))).rejects.toMatchObject({ code: 'conversion-failed' })
  })
})

function validationRule(overrides: Record<string, unknown>): Record<string, unknown> {
  return {
    uid: 'dv',
    type: 'whole',
    operator: 'greaterThan',
    formula1: '5',
    ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }],
    ...overrides,
  }
}

describe('writing data validations', () => {
  it('writes list, date, numeric, custom, and untyped rules with options', async () => {
    const templates = [
      validationRule({
        type: 'list',
        formula1: '["a","b"]',
        allowBlank: true,
        showErrorMessage: true,
        errorStyle: 1,
        error: 'bad',
        errorTitle: 'Bad',
        showInputMessage: true,
        prompt: 'pick',
        promptTitle: 'Pick',
      }),
      validationRule({ type: 'list', formula1: 'not-json' }),
      validationRule({ type: 'list', formula1: '[1,2' }),
      validationRule({ type: 'list', formula1: '["a",2]', operator: 'ignored' }),
      validationRule({ type: 'date', formula1: '45000', formula2: '=A1' }),
      validationRule({ type: 'date', formula1: '=A1' }),
      validationRule({ type: 'whole', formula1: 'abc' }),
      validationRule({ type: 'textLength', formula1: '=[A1]' }),
      validationRule({ type: 'decimal', formula1: '1.5', formula2: '2.5' }),
      validationRule({ type: 'custom', formula1: '=A1>0', operator: 'ignored' }),
      validationRule({ type: 'any', formula1: '=A1', operator: 'ignored' }),
      validationRule({ type: 'whole', operator: undefined, formula1: undefined, formula2: undefined, errorStyle: 0 }),
      validationRule({ type: 'whole', errorStyle: 2 }),
    ]
    const rules = templates.map((template, index) => ({
      ...template,
      ranges: index === 0
        ? [{ startRow: 0, startColumn: index, endRow: 1, endColumn: index + 2 }]
        : [{ startRow: 0, startColumn: index, endRow: 0, endColumn: index }],
    }))
    const xml = await writtenSheetXml(snapshotWith({
      dataValidation: { 'sheet-1': rules },
    }))

    expect(xml).toContain('<dataValidations')
    expect(xml).toContain('type="list"')
    expect(xml).toContain('type="date"')
    expect(xml).toContain('type="custom"')
  })

  it('writes date criteria through the 1904 serial system', async () => {
    const snapshot = snapshotWith({
      dataValidation: {
        'sheet-1': [validationRule({ type: 'date', formula1: '1', formula2: undefined })],
      },
    })
    snapshot.dateSystem = DateSystem.Date1904
    const xml = await writtenSheetXml(snapshot)

    expect(xml).toContain('type="date"')
  })

  it('refuses data-validation types ExcelJS cannot write', async () => {
    for (const type of ['listMultiple', 'checkbox', 'none']) {
      await expect(writeWorkbook(snapshotWith({
        dataValidation: { 'sheet-1': [validationRule({ type })] },
      }))).rejects.toMatchObject({ code: 'conversion-failed' })
    }
  })
})

const PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const JPG_BASE64 = '/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/wAALCAABAAEBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AKp//2Q=='

function drawingResource(overrides: Record<string, unknown>): Record<string, unknown> {
  return {
    unitId: 'office-exchange-workbook',
    subUnitId: 'sheet-1',
    drawingId: 'drawing-1',
    drawingType: 0,
    imageSourceType: 'BASE64',
    source: `data:image/png;base64,${PNG_BASE64}`,
    sheetTransform: {
      from: { column: 0, columnOffset: 0, row: 0, rowOffset: 0 },
      to: { column: 1, columnOffset: 0, row: 1, rowOffset: 0 },
    },
    axisAlignSheetTransform: {
      from: { column: 0, columnOffset: 0, row: 0, rowOffset: 0 },
      to: { column: 1, columnOffset: 0, row: 1, rowOffset: 0 },
    },
    ...overrides,
  }
}

describe('writing drawings and background images', () => {
  it('writes both anchor placements with hyperlinks and skips missing entries', async () => {
    const drawings = {
      'sheet-1': {
        data: {
          'drawing-1': drawingResource({
            xlsx: {
              extension: 'png',
              editAs: 'oneCell',
              ext: { width: 20, height: 30 },
              hyperlink: { hyperlink: 'https://example.invalid/a', tooltip: 'tip' },
            },
          }),
          'drawing-2': drawingResource({
            drawingId: 'drawing-2',
            source: `data:image/jpeg;base64,${JPG_BASE64}`,
            xlsx: { hyperlink: { hyperlink: 'https://example.invalid/b' } },
          }),
          'drawing-3': drawingResource({
            drawingId: 'drawing-3',
            xlsx: undefined,
          }),
          'drawing-4': drawingResource({
            drawingId: 'drawing-4',
            xlsx: { ext: { width: 10, height: 20 } },
          }),
        },
        order: ['drawing-1', 'drawing-missing', 'drawing-2', 'drawing-3', 'drawing-4'],
      },
    }
    const xml = await writtenSheetXml(snapshotWith({ drawings }))

    expect(xml).toContain('<drawing')
  })

  it('refuses a non-image drawing', async () => {
    await expect(writeWorkbook(snapshotWith({
      drawings: {
        'sheet-1': {
          data: { 'drawing-1': drawingResource({ drawingType: 1 }) },
          order: ['drawing-1'],
        },
      },
    }))).rejects.toMatchObject({ code: 'conversion-failed' })
  })

  it('refuses an image source that is not a base64 data URL', async () => {
    await expect(writeWorkbook(snapshotWith({
      drawings: {
        'sheet-1': {
          data: { 'drawing-1': drawingResource({ source: 'https://example.invalid/a.png' }) },
          order: ['drawing-1'],
        },
      },
    }))).rejects.toMatchObject({ code: 'conversion-failed' })
    await expect(writeWorkbook(snapshotWith({
      backgroundImages: { 'sheet-1': { source: 'not-a-data-url', imageSourceType: 'BASE64' } },
    }))).rejects.toMatchObject({ code: 'conversion-failed' })
  })

  it('writes a background image', async () => {
    const xml = await writtenSheetXml(snapshotWith({
      backgroundImages: {
        'sheet-1': {
          source: `data:image/jpeg;base64,${JPG_BASE64}`,
          imageSourceType: 'BASE64',
        },
      },
    }))

    expect(xml).toContain('<picture')
  })
})

function anchorOf(col: number, row: number): Anchor {
  return { col, row, nativeCol: col, nativeRow: row, nativeColOff: 0, nativeRowOff: 0 } as Anchor
}

describe('reading drawings and background images', () => {
  it('maps anchored images, extensions, hyperlinks, and the sheet background', async () => {
    const workbook = new ExcelJS.Workbook()
    const worksheet = workbook.addWorksheet('Images')
    worksheet.getRow(2).height = 30
    worksheet.getColumn(2).width = 12
    const pngId = workbook.addImage({
      base64: `data:image/png;base64,${PNG_BASE64}`,
      extension: 'png',
    })
    const jpgId = workbook.addImage({
      base64: `data:image/jpeg;base64,${JPG_BASE64}`,
      extension: 'jpeg',
    })
    const jpgExtensionId = workbook.addImage({
      base64: `data:image/jpeg;base64,${JPG_BASE64}`,
      extension: 'jpg' as ExcelJS.Image['extension'],
    })
    worksheet.addImage(pngId, {
      tl: anchorOf(0, 0),
      br: anchorOf(2, 2),
      editAs: 'twoCell',
      hyperlinks: { hyperlink: 'https://example.invalid/a', tooltip: 'tip' },
    })
    worksheet.addImage(jpgId, {
      tl: { col: 1.5, row: 1.25 },
      ext: { width: 500, height: 200 },
    })
    worksheet.addImage(pngId, {
      tl: { col: 3, row: 3 },
      ext: { width: 30, height: 40 },
      editAs: 'oneCell',
      hyperlinks: { hyperlink: 'https://example.invalid/b' },
    })
    worksheet.addImage(jpgExtensionId, {
      tl: anchorOf(5, 5),
      br: anchorOf(6, 6),
    })
    worksheet.addBackgroundImage(pngId)

    const bytes = new Uint8Array(await workbook.xlsx.writeBuffer())
    const entries = unzipSync(bytes)
    const sheetXml = strFromU8(entries['xl/worksheets/sheet1.xml'] ?? new Uint8Array(0))
    const stripped = zipSync({
      ...entries,
      'xl/worksheets/sheet1.xml': strToU8(sheetXml.replace(/<sheetFormatPr[^>]*\/>/u, '')),
    })
    const data = await readWorkbook(stripped)
    const text = resourceData(data, 'SHEET_DRAWING_PLUGIN')
    if (text === undefined) throw new Error('workbook has no drawing resource')
    const drawings = JSON.parse(text) as Record<string, {
      readonly data: Record<string, {
        readonly source: string
        readonly sheetTransform: { readonly to: { readonly column: number; readonly row: number } }
        readonly xlsx?: {
          readonly extension?: string
          readonly editAs?: string
          readonly ext?: { readonly width: number; readonly height: number }
          readonly hyperlink?: { readonly hyperlink: string; readonly tooltip?: string }
        }
      }>
      readonly order: readonly string[]
    }>
    const sheet = drawings['sheet-1']
    if (sheet === undefined) throw new Error('sheet has no drawings')
    expect(sheet.order).toHaveLength(4)
    expect(sheet.data['sheet-drawing-1-1']?.xlsx).toMatchObject({
      extension: 'png',
      editAs: 'twoCell',
      hyperlink: { hyperlink: 'https://example.invalid/a', tooltip: 'tip' },
    })
    expect(sheet.data['sheet-drawing-1-2']?.xlsx).toMatchObject({
      extension: 'jpeg',
      ext: { width: 500, height: 200 },
    })
    expect(sheet.data['sheet-drawing-1-2']?.sheetTransform.to.column).toBeGreaterThan(1)
    expect(sheet.data['sheet-drawing-1-3']?.xlsx?.hyperlink).toEqual({ hyperlink: 'https://example.invalid/b' })
    expect(sheet.data['sheet-drawing-1-4']?.xlsx).toMatchObject({ extension: 'jpg' })
    expect(data.sheets['sheet-1']?.backgroundImage).toMatchObject({
      imageSourceType: 'BASE64',
    })

    const reread = await readWorkbook(new Uint8Array(await writeWorkbook(data)))
    expect(resourceData(reread, 'SHEET_DRAWING_PLUGIN')).toBeDefined()
  })
})

describe('malformed resources', () => {
  it('reports malformed plugin resource JSON', async () => {
    const snapshot = snapshotWith({})
    snapshot.resources = [{ name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN', data: 'not-json' }]

    const failure = await writeWorkbook(snapshot).then(() => undefined, (error: unknown) => error)

    expect(failure).toMatchObject({ code: 'conversion-failed' })
    expect(failure instanceof Error ? failure.message : '').toContain('SHEET_CONDITIONAL_FORMATTING_PLUGIN')
  })
})
