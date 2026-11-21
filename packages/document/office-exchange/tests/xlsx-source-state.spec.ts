/** Workbook source-state derivation from the OPC package and ExcelJS model. */
import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import type { IWorkbookData } from '@univerjs/core'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from '../src/config.ts'
import { readWorkbookWithSource } from '../src/provider/xlsx.ts'

const SPREADSHEET = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const STRICT_REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships'
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

interface SourceStateView {
  readonly anchorMode?: string
  readonly anchors?: readonly unknown[]
  readonly sheets?: readonly unknown[]
  readonly sourceParts?: readonly string[]
}

function sourceOf(data: IWorkbookData): SourceStateView | undefined {
  return (data as IWorkbookData & { officeExchange?: { source?: SourceStateView } })
    .officeExchange?.source
}

interface PackageOptions {
  readonly workbook?: string
  readonly workbookRels?: string
  readonly sheet?: string
  readonly entries?: Readonly<Record<string, Uint8Array>>
}

function buildPackage(options: PackageOptions = {}): Uint8Array {
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
    'xl/workbook.xml': strToU8(options.workbook
      ?? '<?xml version="1.0"?>'
      + `<workbook xmlns="${SPREADSHEET}" xmlns:r="${REL}">`
      + '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>'),
    'xl/_rels/workbook.xml.rels': strToU8(options.workbookRels
      ?? '<?xml version="1.0"?>'
      + `<Relationships xmlns="${PKG_REL}">`
      + `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`
      + '</Relationships>'),
    'xl/worksheets/sheet1.xml': strToU8(options.sheet
      ?? '<?xml version="1.0"?>'
      + `<worksheet xmlns="${SPREADSHEET}" xmlns:r="${REL}">`
      + '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'),
    ...options.entries,
  })
}

function readCentralUint32(bytes: Uint8Array, offset: number): number {
  return (bytes[offset]! | (bytes[offset + 1]! << 8) | (bytes[offset + 2]! << 16) | (bytes[offset + 3]! << 24)) >>> 0
}

function writeCentralUint32(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = value & 0xff
  bytes[offset + 1] = (value >>> 8) & 0xff
  bytes[offset + 2] = (value >>> 16) & 0xff
  bytes[offset + 3] = (value >>> 24) & 0xff
}

function corruptEntryCrc(source: Uint8Array, name: string): Uint8Array {
  const bytes = source.slice()
  let eocd = bytes.length - 22
  while (eocd >= 0 && readCentralUint32(bytes, eocd) !== 0x0605_4b50) eocd -= 1
  if (eocd < 0) throw new Error('package has no end-of-central-directory record')
  let cursor = readCentralUint32(bytes, eocd + 16)
  const count = bytes[eocd + 10]! | (bytes[eocd + 11]! << 8)
  for (let index = 0; index < count; index += 1) {
    if (readCentralUint32(bytes, cursor) !== 0x0201_4b50) throw new Error('malformed central directory')
    const nameLength = bytes[cursor + 28]! | (bytes[cursor + 29]! << 8)
    const extraLength = bytes[cursor + 30]! | (bytes[cursor + 31]! << 8)
    const commentLength = bytes[cursor + 32]! | (bytes[cursor + 33]! << 8)
    const entryName = new TextDecoder().decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength))
    if (entryName === name) {
      const localOffset = readCentralUint32(bytes, cursor + 42)
      writeCentralUint32(bytes, localOffset + 14, readCentralUint32(bytes, localOffset + 14) ^ 0xff)
      writeCentralUint32(bytes, cursor + 16, readCentralUint32(bytes, cursor + 16) ^ 0xff)
      return bytes
    }
    cursor += 46 + nameLength + extraLength + commentLength
  }
  throw new Error(`package has no entry ${name}`)
}

describe('readWorkbookWithSource', () => {
  it('records per-unit anchors for uniquely addressable cells', async () => {
    const data = await readWorkbookWithSource(buildPackage(), 'xlsx')

    expect(sourceOf(data)).toMatchObject({
      anchorMode: 'per-unit',
      sourceParts: ['xl/worksheets/sheet1.xml'],
    })
    expect(sourceOf(data)?.anchors).toHaveLength(1)
  })

  it('falls back to sheet summaries when two cells claim one reference', async () => {
    const data = await readWorkbookWithSource(buildPackage({
      sheet: '<?xml version="1.0"?>'
        + `<worksheet xmlns="${SPREADSHEET}" xmlns:r="${REL}">`
        + '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="A1"><v>2</v></c></row></sheetData>'
        + '</worksheet>',
    }), 'xlsx')

    expect(sourceOf(data)).toMatchObject({ anchorMode: 'sheet-summary' })
    expect(sourceOf(data)?.sheets).toHaveLength(1)
  })

  it('omits the sidecar when worksheet scanning fails', async () => {
    const data = await readWorkbookWithSource(buildPackage({
      sheet: '<?xml version="1.0"?>'
        + `<worksheet xmlns="${SPREADSHEET}" xmlns:r="${REL}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">`
        + '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
        + '<mc:AlternateContent><mc:Choice Requires="x14ac"><extLst/></mc:Choice></mc:AlternateContent>'
        + '</worksheet>',
    }), 'xlsx')

    expect(sourceOf(data)).toBeUndefined()
  })

  it('omits the sidecar when the anchor payload exceeds its budget', async () => {
    const data = await readWorkbookWithSource(buildPackage(), 'xlsx', {
      ...DEFAULT_OFFICE_EXCHANGE_CONFIG,
      maxSourceAnchorBytes: 1,
    })

    expect(sourceOf(data)).toBeUndefined()
  })

  it('ignores directory entries when indexing part lengths', async () => {
    const data = await readWorkbookWithSource(buildPackage({
      entries: { 'xl/': new Uint8Array(0) },
    }), 'xlsx')

    expect(sourceOf(data)).toBeDefined()
  })

  it('omits the sidecar when the sheets element lists no sheet entries', async () => {
    const data = await readWorkbookWithSource(buildPackage({
      workbook: '<?xml version="1.0"?>'
        + `<workbook xmlns="${SPREADSHEET}" xmlns:r="${REL}"><sheets><other/></sheets></workbook>`,
    }), 'xlsx')

    expect(sourceOf(data)).toBeUndefined()
  })

  it('omits the sidecar when a worksheet part cannot be read', async () => {
    const data = await readWorkbookWithSource(corruptEntryCrc(buildPackage(), 'xl/worksheets/sheet1.xml'), 'xlsx')

    expect(sourceOf(data)).toBeUndefined()
  })

  it('omits the sidecar when the workbook or relationships part cannot be read', async () => {
    const missingWorkbook = await readWorkbookWithSource(corruptEntryCrc(buildPackage(), 'xl/workbook.xml'), 'xlsx')
    const missingRelationships = await readWorkbookWithSource(
      corruptEntryCrc(buildPackage(), 'xl/_rels/workbook.xml.rels'),
      'xlsx',
    )

    expect(sourceOf(missingWorkbook)).toBeUndefined()
    expect(sourceOf(missingRelationships)).toBeUndefined()
  })

  it('omits the sidecar when the archive exceeds the source entry bound', async () => {
    const extras: Record<string, Uint8Array> = {}
    for (let index = 0; index < 10_001; index += 1) extras[`extra/${String(index)}.txt`] = strToU8('x')
    const data = await readWorkbookWithSource(buildPackage({ entries: extras }), 'xlsx')

    expect(sourceOf(data)).toBeUndefined()
  })

  it('omits the sidecar when the workbook declares no sheets', async () => {
    const data = await readWorkbookWithSource(buildPackage({
      workbook: '<?xml version="1.0"?>'
        + `<workbook xmlns="${SPREADSHEET}" xmlns:r="${REL}"><sheets/></workbook>`,
    }), 'xlsx')

    expect(sourceOf(data)).toBeUndefined()
  })

  it('omits the sidecar when a chartsheet makes the model sheet count differ', async () => {
    const data = await readWorkbookWithSource(buildPackage({
      workbookRels: '<?xml version="1.0"?>'
        + `<Relationships xmlns="${PKG_REL}">`
        + `<Relationship Id="rId1" Type="${REL}/chartsheet" Target="worksheets/sheet1.xml"/>`
        + '</Relationships>',
    }), 'xlsx')

    expect(sourceOf(data)).toBeUndefined()
  })

  it('omits the sidecar when the sheets element is not an object', async () => {
    const data = await readWorkbookWithSource(buildPackage({
      workbook: '<?xml version="1.0"?>'
        + `<workbook xmlns="${SPREADSHEET}" xmlns:r="${REL}"><sheets>text</sheets></workbook>`,
    }), 'xlsx')

    expect(sourceOf(data)).toBeUndefined()
  })

  it('omits the sidecar when a relationship entry has no target', async () => {
    const data = await readWorkbookWithSource(buildPackage({
      workbookRels: '<?xml version="1.0"?>'
        + `<Relationships xmlns="${PKG_REL}">`
        + `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`
        + '<Relationship Id="rId9" Type="http://example.invalid/unknown"/>'
        + '</Relationships>',
    }), 'xlsx')

    expect(sourceOf(data)).toBeUndefined()
  })

  it('resolves a strict worksheet relationship with an absolute part target', async () => {
    const data = await readWorkbookWithSource(buildPackage({
      workbookRels: '<?xml version="1.0"?>'
        + `<Relationships xmlns="${PKG_REL}">`
        + `<Relationship Id="rId1" Type="${STRICT_REL}/worksheet" Target="/xl/worksheets/sheet1.xml"/>`
        + '</Relationships>',
    }), 'xlsx')

    expect(sourceOf(data)).toMatchObject({ sourceParts: ['xl/worksheets/sheet1.xml'] })
  })

  it('records an empty sheet without cells', async () => {
    const data = await readWorkbookWithSource(buildPackage({
      sheet: '<?xml version="1.0"?>'
        + `<worksheet xmlns="${SPREADSHEET}" xmlns:r="${REL}"><sheetData/></worksheet>`,
    }), 'xlsx')

    expect(sourceOf(data)).toMatchObject({ anchorMode: 'per-unit', sourceParts: ['xl/worksheets/sheet1.xml'] })
  })
})
