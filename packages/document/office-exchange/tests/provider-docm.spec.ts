/**
 * Macro-enabled provider writes: a non-intersecting edit must carry the VBA
 * project and the macro-enabled main content type unchanged.
 *
 * The tests drive `writeDocumentFromSource` and `writeWorkbookFromSource`
 * directly, so the provider splice path is exercised without the Remote
 * whole-file save gate.
 */
import { readFile } from 'node:fs/promises'
import { Context } from '@forwork/cordis'
import type { IDocumentData, IWorkbookData } from '@univerjs/core'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { open } from '../src/package/ooxml-package.ts'
import LocalOfficeExchange from '../src/provider/index.ts'
import { readDocumentWithSource, writeDocumentFromSource } from '../src/provider/docx.ts'
import { readWorkbookWithSource, writeWorkbookFromSource } from '../src/provider/xlsx.ts'

const DOCM_FIXTURE = new URL('./fixtures/macro.docm', import.meta.url)
const XLSM_FIXTURE = new URL('./fixtures/macro.xlsm', import.meta.url)
const CONTENT_TYPES_PART = '[Content_Types].xml'
const VBA_CONTENT_TYPE = 'application/vnd.ms-office.vbaProject'
const DOCM_MAIN_CONTENT_TYPE = 'application/vnd.ms-word.document.macroEnabled.main+xml'
const DOCM_MACRO_FREE_CONTENT_TYPE
  = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
const XLSM_MAIN_CONTENT_TYPE = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
const XLSM_MACRO_FREE_CONTENT_TYPE
  = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'

interface SourceCarrier {
  /** Provider-owned sidecar, present when the read resolved every unit. */
  officeExchange?: {
    /** Source description recorded at read time. */
    readonly source?: {
      /** Format recorded for the source package. */
      readonly format?: string
    }
  }
}

function entryOf(entries: Readonly<Record<string, Uint8Array>>, name: string): Uint8Array {
  const bytes = entries[name]
  if (bytes === undefined) throw new Error(`package has no ${name}`)
  return bytes
}

/**
 * The content type an OPC override assigns to one part.
 * @param contentTypes - `[Content_Types].xml` text.
 * @param partName - Package part name including its leading slash.
 * @returns The declared content type.
 */
function contentTypeOf(contentTypes: string, partName: string): string {
  const pattern = new RegExp(
    `<Override\\b[^>]*PartName="${partName.replaceAll('/', '\\/')}"[^>]*ContentType="([^"]+)"`,
    'u',
  )
  const match = pattern.exec(contentTypes)
  if (match === null) throw new Error(`content types declare no override for ${partName}`)
  return match[1] ?? ''
}

function documentBodyOf(data: IDocumentData): NonNullable<IDocumentData['body']> {
  if (data.body === undefined) throw new Error('document snapshot has no body')
  return data.body
}

describe('macro-enabled DOCX provider save', () => {
  it('carries the VBA project and macro-enabled content type through an edit elsewhere', async () => {
    const source = new Uint8Array(await readFile(DOCM_FIXTURE))
    const snapshot = await readDocumentWithSource(source, 'docm')
    expect((snapshot as SourceCarrier).officeExchange?.source?.format).toBe('docm')

    const body = documentBodyOf(snapshot)
    const text = body.dataStream ?? ''
    const start = text.indexOf('Quarterly')
    expect(start).toBeGreaterThanOrEqual(0)
    body.dataStream = `${text.slice(0, start)}Financial${text.slice(start + 'Quarterly'.length)}`

    const result = await writeDocumentFromSource(snapshot, source)
    expect(result.bytes).not.toEqual(source)

    const sourceEntries = unzipSync(source)
    const writtenEntries = unzipSync(result.bytes)
    const sourceTypes = strFromU8(entryOf(sourceEntries, CONTENT_TYPES_PART))
    const writtenTypes = strFromU8(entryOf(writtenEntries, CONTENT_TYPES_PART))

    expect(entryOf(writtenEntries, 'word/vbaProject.bin')).toEqual(entryOf(sourceEntries, 'word/vbaProject.bin'))
    expect(entryOf(writtenEntries, 'word/_rels/document.xml.rels'))
      .toEqual(entryOf(sourceEntries, 'word/_rels/document.xml.rels'))
    expect(strFromU8(entryOf(writtenEntries, 'word/document.xml'))).toContain('Financial Report')
    expect(strFromU8(entryOf(writtenEntries, 'word/document.xml'))).not.toContain('Quarterly Report')

    expect(contentTypeOf(sourceTypes, '/word/document.xml')).toBe(DOCM_MAIN_CONTENT_TYPE)
    expect(contentTypeOf(writtenTypes, '/word/document.xml')).toBe(DOCM_MAIN_CONTENT_TYPE)
    expect(contentTypeOf(writtenTypes, '/word/vbaProject.bin')).toBe(VBA_CONTENT_TYPE)
    expect(writtenTypes).not.toContain(DOCM_MACRO_FREE_CONTENT_TYPE)

    expect(entryOf(writtenEntries, 'word/document.xml')).toBeDefined()
    expect(writtenEntries['xl/workbook.xml']).toBeUndefined()
    expect(writtenTypes).not.toContain(XLSM_MAIN_CONTENT_TYPE)
    expect(writtenTypes).not.toContain(XLSM_MACRO_FREE_CONTENT_TYPE)

    const disposed = [
      ...result.disposition.spliced.names,
      ...result.disposition.added.names,
      ...result.disposition.removed.names,
    ]
    expect(disposed).not.toContain('word/vbaProject.bin')
  })
})

describe('macro-enabled XLSX provider save', () => {
  it('carries the VBA project and macro-enabled content type through an edit elsewhere', async () => {
    const source = new Uint8Array(await readFile(XLSM_FIXTURE))
    const snapshot = await readWorkbookWithSource(source, 'xlsm')
    expect((snapshot as SourceCarrier).officeExchange?.source?.format).toBe('xlsm')

    editWorkbookCell(snapshot, 'sheet-1', 1, 0, 'Sprocket')

    const result = await writeWorkbookFromSource(snapshot, source)
    expect(result.bytes).not.toEqual(source)

    const sourceEntries = unzipSync(source)
    const writtenEntries = unzipSync(result.bytes)
    const sourceTypes = strFromU8(entryOf(sourceEntries, CONTENT_TYPES_PART))
    const writtenTypes = strFromU8(entryOf(writtenEntries, CONTENT_TYPES_PART))

    expect(entryOf(writtenEntries, 'xl/vbaProject.bin')).toEqual(entryOf(sourceEntries, 'xl/vbaProject.bin'))
    expect(entryOf(writtenEntries, 'xl/_rels/workbook.xml.rels'))
      .toEqual(entryOf(sourceEntries, 'xl/_rels/workbook.xml.rels'))
    const writtenSheet = strFromU8(entryOf(writtenEntries, 'xl/worksheets/sheet1.xml'))
    expect(writtenSheet).toContain('Sprocket')
    expect(writtenSheet).not.toContain('Widget')

    expect(contentTypeOf(sourceTypes, '/xl/workbook.xml')).toBe(XLSM_MAIN_CONTENT_TYPE)
    expect(contentTypeOf(writtenTypes, '/xl/workbook.xml')).toBe(XLSM_MAIN_CONTENT_TYPE)
    expect(contentTypeOf(writtenTypes, '/xl/vbaProject.bin')).toBe(VBA_CONTENT_TYPE)
    expect(writtenTypes).not.toContain(XLSM_MACRO_FREE_CONTENT_TYPE)

    expect(entryOf(writtenEntries, 'xl/workbook.xml')).toBeDefined()
    expect(writtenEntries['word/document.xml']).toBeUndefined()
    expect(writtenTypes).not.toContain(DOCM_MAIN_CONTENT_TYPE)
    expect(writtenTypes).not.toContain(DOCM_MACRO_FREE_CONTENT_TYPE)

    const disposed = [
      ...result.disposition.spliced.names,
      ...result.disposition.added.names,
      ...result.disposition.removed.names,
    ]
    expect(disposed).not.toContain('xl/vbaProject.bin')
  })
})

describe('macro-enabled provider fallback', () => {
  it('refuses a DOCM save with no anchored source and produces no bytes', async () => {
    const source = new Uint8Array(await readFile(DOCM_FIXTURE))
    const ctx = new Context()
    await ctx.plugin(LocalOfficeExchange)
    const office = ctx.officeExchange

    const anchored = await readDocumentWithSource(source, 'docm')
    await expect(office.write({ kind: 'document', data: anchored }, 'docm')).rejects.toMatchObject({
      code: 'unsupported-feature',
    })

    const detached = await readDocumentWithSource(source, 'docm')
    delete (detached as SourceCarrier).officeExchange
    const sourcePackage = open(source, { maxEntries: 10_000 }).source
    await expect(office.write({ kind: 'document', data: detached }, 'docm', sourcePackage)).rejects.toMatchObject({
      code: 'unsupported-feature',
    })

    const plain = await office.write({ kind: 'document', data: detached }, 'docx')
    expect(plain.bytes.byteLength).toBeGreaterThan(0)
  })

  it('refuses an XLSM save with no anchored source and produces no bytes', async () => {
    const source = new Uint8Array(await readFile(XLSM_FIXTURE))
    const ctx = new Context()
    await ctx.plugin(LocalOfficeExchange)
    const office = ctx.officeExchange

    const anchored = await readWorkbookWithSource(source, 'xlsm')
    await expect(office.write({ kind: 'workbook', data: anchored }, 'xlsm')).rejects.toMatchObject({
      code: 'unsupported-feature',
    })

    const detached = await readWorkbookWithSource(source, 'xlsm')
    delete (detached as SourceCarrier).officeExchange
    const sourcePackage = open(source, { maxEntries: 10_000 }).source
    await expect(office.write({ kind: 'workbook', data: detached }, 'xlsm', sourcePackage)).rejects.toMatchObject({
      code: 'unsupported-feature',
    })

    const plain = await office.write({ kind: 'workbook', data: detached }, 'xlsx')
    expect(plain.bytes.byteLength).toBeGreaterThan(0)
  })
})

/**
 * Write one worksheet cell value, as an editor would, without touching the
 * workbook-level parts the VBA relationship lives in.
 * @param workbook - Workbook snapshot to edit in place.
 * @param sheetId - Sheet id recorded in the snapshot.
 * @param row - Zero-based row index.
 * @param column - Zero-based column index.
 * @param value - Replacement cell value.
 */
function editWorkbookCell(
  workbook: IWorkbookData,
  sheetId: string,
  row: number,
  column: number,
  value: string,
): void {
  const cell = workbook.sheets[sheetId]?.cellData?.[row]?.[column]
  if (cell === undefined) throw new Error(`workbook sheet ${sheetId} has no cell ${String(row)}:${String(column)}`)
  cell.v = value
}
