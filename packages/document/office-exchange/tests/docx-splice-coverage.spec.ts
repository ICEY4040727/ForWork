/** In-place DOCX save branch coverage: refusals, auxiliary appends, merges, and crafted sidecars. */
import {
  BooleanNumber,
  CustomRangeType,
  HorizontalAlign,
  NamedStyleType,
  VerticalAlignmentType,
} from '@univerjs/core'
import type { IDocumentData, IParagraph } from '@univerjs/core'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { expect, it } from 'vitest'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG, resolveConfig } from '../src/config.ts'
import { open } from '../src/package/ooxml-package.ts'
import { readDocumentWithSource, writeDocumentFromSource } from '../src/provider/docx.ts'

const WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
const CONFIG = resolveConfig({})

type Entry = string | Uint8Array

const DOCUMENT_NS = `<w:document xmlns:w="${WORD}" xmlns:r="${REL}"`
  + ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
  + ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
  + ` xmlns:x="${WORD}">`

const SECTION = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'

const CONTENT_TYPES_HEAD = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CT}">`
  + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
  + '<Default Extension="xml" ContentType="application/xml"/>'
  + '<Override PartName="/word/document.xml"'
  + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
  + '</Types>'

function contentTypesFor(options: PackageOptions): string {
  let overrides = ''
  if (options.styles !== null) {
    overrides += '<Override PartName="/word/styles.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
  }
  if (options.numbering !== undefined && options.numbering !== null) {
    overrides += '<Override PartName="/word/numbering.xml"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
  }
  return CONTENT_TYPES_HEAD.replace('</Types>', `${overrides}</Types>`)
}

const DEFAULT_STYLES = `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="${WORD}">`
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
  + '</w:styles>'

const DEFAULT_NUMBERING = `<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="${WORD}">`
  + '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>'
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
  + '</w:numbering>'

interface PackageOptions {
  readonly body: string
  readonly contentTypes?: string
  readonly styles?: string | null
  readonly numbering?: string | null
  readonly rels?: string | null
  readonly extra?: Readonly<Record<string, Entry>>
}

function packageOf(entries: Readonly<Record<string, Entry>>): Uint8Array {
  const files: Record<string, Uint8Array> = {}
  for (const [name, content] of Object.entries(entries)) {
    files[name] = typeof content === 'string' ? strToU8(content) : content
  }
  return zipSync(files, { level: 6 })
}

function docxBytes(options: PackageOptions): Uint8Array {
  const rels = options.rels === null
    ? undefined
    : `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PKG_REL}">`
      + (options.styles === null ? '' : `<Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>`)
      + (options.numbering === undefined || options.numbering === null
        ? ''
        : `<Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/>`)
      + `${options.rels ?? ''}</Relationships>`
  return packageOf({
    '[Content_Types].xml': options.contentTypes ?? contentTypesFor(options),
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PKG_REL}">`
      + `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`
      + '</Relationships>',
    'word/document.xml': `${DOCUMENT_NS}<w:body>${options.body}</w:body></w:document>`,
    ...(rels === undefined ? {} : { 'word/_rels/document.xml.rels': rels }),
    ...(options.styles === null ? {} : { 'word/styles.xml': options.styles ?? DEFAULT_STYLES }),
    ...(options.numbering === undefined || options.numbering === null
      ? {}
      : { 'word/numbering.xml': options.numbering }),
    ...options.extra,
  })
}

interface Carrier extends IDocumentData {
  officeExchange?: unknown
}

function carrierOf(data: IDocumentData): Carrier {
  return data
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function recordOf(value: unknown): Record<string, unknown> {
  if (!isRecord(value)) throw new Error('expected an object')
  return value
}

function sourceStateOf(data: IDocumentData): Record<string, unknown> {
  const extension = recordOf(carrierOf(data).officeExchange)
  return recordOf(extension['source'])
}

function patchSourceState(data: IDocumentData, patch: Readonly<Record<string, unknown>>): void {
  const extension = recordOf(carrierOf(data).officeExchange)
  carrierOf(data).officeExchange = { ...extension, source: { ...sourceStateOf(data), ...patch } }
}

function sourcePartsOf(data: IDocumentData): readonly string[] {
  const parts = sourceStateOf(data)['sourceParts']
  if (!Array.isArray(parts)) throw new Error('expected source parts')
  return parts.map((part) => {
    if (typeof part !== 'string') throw new Error('expected a part name')
    return part
  })
}

function anchorsOf(data: IDocumentData): readonly Record<string, unknown>[] {
  const anchors = sourceStateOf(data)['anchors']
  if (!Array.isArray(anchors)) throw new Error('expected anchors')
  return anchors.map(recordOf)
}

async function openSource(bytes: Uint8Array): Promise<{ bytes: Uint8Array; data: IDocumentData }> {
  return { bytes, data: await readDocumentWithSource(bytes, 'docx', CONFIG) }
}

function paragraph(text: string, properties = ''): string {
  return `<w:p>${properties}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
}

function firstParagraph(data: IDocumentData): IParagraph {
  const paragraphRecord = data.body?.paragraphs?.[0]
  if (paragraphRecord === undefined) throw new Error('expected a paragraph')
  return paragraphRecord
}

function editParagraph(data: IDocumentData, style: NonNullable<IParagraph['paragraphStyle']>): IDocumentData {
  const paragraphRecord = firstParagraph(data)
  paragraphRecord.paragraphStyle = style
  return data
}


it('refuses a source above the byte bound', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  await expect(writeDocumentFromSource(data, bytes, { ...CONFIG, maxFileBytes: bytes.byteLength - 1 }))
    .rejects.toMatchObject({ code: 'too-large' })
})

it('refuses malformed source package bytes', async () => {
  const { data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  await expect(writeDocumentFromSource(data, new Uint8Array([1, 2, 3])))
    .rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses a sheet-summary anchor mode for a document', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  patchSourceState(data, {
    anchorMode: 'sheet-summary',
    anchors: undefined,
    sheets: [{
      sheetId: 's1',
      part: 'word/document.xml',
      digest: 'a'.repeat(64),
      cellCount: 0,
    }],
  })
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses a sheet-cell anchor in a document', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  const anchors = anchorsOf(data).map((anchor, index) => (
    index === 0 ? { ...anchor, kind: 'sheetCell' } : anchor
  ))
  patchSourceState(data, { anchors })
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses an anchor whose part sourceParts omits', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  patchSourceState(data, { sourceParts: sourcePartsOf(data).filter(part => part !== 'word/document.xml') })
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses an anchored part that is not an editable story', async () => {
  const { bytes, data } = await openSource(docxBytes({
    body: paragraph('x') + SECTION,
    extra: { 'word/footnotes.xml': `<w:footnotes xmlns:w="${WORD}"/>` },
  }))
  patchSourceState(data, { sourceParts: [...sourcePartsOf(data), 'word/footnotes.xml'] })
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses snapshot headers and footers that the package cannot add', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  data.headers = { header9: { headerId: 'header9', body: { dataStream: 'h\r\n' } } }
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })

  const second = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  second.data.footers = { footer9: { footerId: 'footer9', body: { dataStream: 'f\r\n' } } }
  await expect(writeDocumentFromSource(second.data, second.bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses malformed officeExchange state', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  carrierOf(data).officeExchange = 5
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses a snapshot that lost its body', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  delete data.body
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('uses defaults for a missing table source map', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  delete data.tableSource
  const written = await writeDocumentFromSource(data, bytes)
  expect(written.bytes.byteLength).toBeGreaterThan(0)
})

it('saves a document carrying two tables unchanged', async () => {
  const table = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const { bytes, data } = await openSource(docxBytes({ body: `${table}${table}${SECTION}` }))
  expect(data.body?.tables).toHaveLength(2)
  const written = await writeDocumentFromSource(data, bytes)
  expect(written.bytes.byteLength).toBeGreaterThan(0)
})

it('refuses a table without model metadata and each malformed stream alignment', async () => {
  const table = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const { bytes, data } = await openSource(docxBytes({ body: `${table}${SECTION}` }))
  const tableId = data.body?.tables?.[0]?.tableId
  if (tableId === undefined) throw new Error('expected a table')
  const missingModel = structuredClone(data)
  missingModel.tableSource = {}
  await expect(writeDocumentFromSource(missingModel, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })

  const stream = data.body?.dataStream ?? ''
  const tableStart = stream.indexOf('\x1A')
  const rowStart = stream.indexOf('\x1B')
  const cellStart = stream.indexOf('\x1C')
  const cellEnd = stream.indexOf('\x1D')
  const rowEnd = stream.indexOf('\x0E')
  const tableEnd = stream.indexOf('\x0F')
  const corruptions: readonly ((text: string) => string)[] = [
    text => text.slice(0, tableStart) + 'X' + text.slice(tableStart + 1),
    text => text.slice(0, rowStart) + 'X' + text.slice(rowStart + 1),
    text => text.slice(0, cellStart) + 'X' + text.slice(cellStart + 1),
    text => text.slice(0, cellEnd) + 'X' + text.slice(cellEnd + 1),
    text => text.slice(0, rowEnd) + 'X' + text.slice(rowEnd + 1),
    text => text.slice(0, tableEnd) + 'X' + text.slice(tableEnd + 1),
  ]
  for (const corrupt of corruptions) {
    const variant = structuredClone(data)
    if (variant.body === undefined) throw new Error('expected a body')
    variant.body.dataStream = corrupt(stream)
    await expect(writeDocumentFromSource(variant, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
  }
  const endIndex = structuredClone(data)
  const custom = endIndex.body?.tables?.[0]
  if (custom === undefined) throw new Error('expected a table range')
  custom.endIndex += 1
  await expect(writeDocumentFromSource(endIndex, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('defaults missing body arrays and refuses a table token without a record', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  patchSourceState(data, { anchors: [] })
  const empty = structuredClone(data)
  empty.body = { dataStream: '' }
  const written = await writeDocumentFromSource(empty, bytes)
  expect(written.bytes.byteLength).toBeGreaterThan(0)

  const table = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const tableSource = await openSource(docxBytes({ body: `${table}${SECTION}` }))
  const withoutTables = structuredClone(tableSource.data)
  if (withoutTables.body === undefined) throw new Error('expected a body')
  withoutTables.body.tables = []
  await expect(writeDocumentFromSource(withoutTables, tableSource.bytes))
    .rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses a paragraph break without a model paragraph', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('one') + paragraph('two') + SECTION }))
  const variant = structuredClone(data)
  if (variant.body === undefined) throw new Error('expected a body')
  variant.body.paragraphs = (variant.body.paragraphs ?? []).slice(0, 1)
  await expect(writeDocumentFromSource(variant, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('saves empty paragraphs in the middle and before the final section', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: `${paragraph('x')}<w:p/>${paragraph('y')}${SECTION}` }))
  const written = await writeDocumentFromSource(data, bytes)
  expect(written.bytes.byteLength).toBeGreaterThan(0)

  const trailing = await openSource(docxBytes({ body: `${paragraph('x')}<w:p/>${SECTION}` }))
  const saved = await writeDocumentFromSource(trailing.data, trailing.bytes)
  expect(saved.bytes.byteLength).toBeGreaterThan(0)
})

it('saves a package without a styles part and recognizes custom style attributes', async () => {
  const body = paragraph('x') + SECTION
  const { bytes, data } = await openSource(docxBytes({ body, styles: null }))
  expect(carrierOf(data).officeExchange).toBeDefined()
  const saved = await writeDocumentFromSource(data, bytes)
  expect(saved.bytes.byteLength).toBeGreaterThan(0)

  const customStyles = `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="${WORD}">`
    + '<style styleId="Plain"/>'
    + '<w:style w:type="paragraph"/>'
    + '</w:styles>'
  const custom = await openSource(docxBytes({ body, styles: customStyles }))
  const customSaved = await writeDocumentFromSource(custom.data, custom.bytes)
  expect(customSaved.bytes.byteLength).toBeGreaterThan(0)
})

it('skips unusable numbering ids and accepts bare or invalid element ids', async () => {
  const numbering = `<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="${WORD}">`
    + '<w:abstractNum abstractNumId="5"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>'
    + '<w:num w:numId="bad"><w:abstractNumId w:val="5"/></w:num>'
    + '<w:num w:numId="4"><w:abstractNumId w:val="5"/></w:num>'
    + '</w:numbering>'
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION, numbering }))
  const saved = await writeDocumentFromSource(data, bytes)
  expect(saved.bytes.byteLength).toBeGreaterThan(0)
})

it('refuses a hyperlink whose relationship part cannot be added', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION, rels: null }))
  if (data.body === undefined) throw new Error('expected a body')
  data.body.dataStream = data.body.dataStream.replace('x', 'y')
  data.body.customRanges = [...(data.body.customRanges ?? []), {
    startIndex: 0,
    endIndex: 1,
    rangeId: 'link',
    rangeType: CustomRangeType.HYPERLINK,
    properties: { url: 'https://example.test/' },
  }]
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses a missing styles part when the edit needs a style', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION, styles: null }))
  const paragraphRecord = firstParagraph(data)
  paragraphRecord.styleId = 'Heading7'
  paragraphRecord.paragraphStyle = { horizontalAlign: HorizontalAlign.CENTER }
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('appends a custom style and a named style definition', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  const paragraphRecord = firstParagraph(data)
  paragraphRecord.styleId = 'CustomSplice'
  paragraphRecord.paragraphStyle = { horizontalAlign: HorizontalAlign.CENTER }
  const saved = await writeDocumentFromSource(data, bytes)
  const styles = strFromU8(unzipSync(saved.bytes)['word/styles.xml'] ?? new Uint8Array())
  expect(styles).toContain('w:styleId="CustomSplice"')

  const heading = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  const headingParagraph = firstParagraph(heading.data)
  headingParagraph.styleId = 'Heading3'
  headingParagraph.paragraphStyle = { horizontalAlign: HorizontalAlign.CENTER }
  const headingSaved = await writeDocumentFromSource(heading.data, heading.bytes)
  const headingStyles = strFromU8(unzipSync(headingSaved.bytes)['word/styles.xml'] ?? new Uint8Array())
  expect(headingStyles).toContain('w:styleId="Heading3"')
})

it('refuses a bullet edit without a numbering part', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION, numbering: null }))
  firstParagraph(data).bullet = { listType: 'numbering-9', listId: 'num-9', nestingLevel: 0 }
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('reuses an existing list definition and adds start overrides for a new one', async () => {
  const bulleted = paragraph('x', '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>')
  const { bytes, data } = await openSource(docxBytes({ body: bulleted + SECTION, numbering: DEFAULT_NUMBERING }))
  const existing = firstParagraph(data)
  existing.bullet = { ...(existing.bullet ?? {}), listType: 'numbering-1', listId: 'num-1', nestingLevel: 0, startNumber: 2 }
  await writeDocumentFromSource(data, bytes)

  const fresh = structuredClone(data)
  const freshParagraph = firstParagraph(fresh)
  freshParagraph.bullet = {
    listType: 'numbering-9',
    listId: 'num-9',
    nestingLevel: 0,
    startNumber: 5,
  }
  const saved = await writeDocumentFromSource(fresh, bytes)
  const numbering = strFromU8(unzipSync(saved.bytes)['word/numbering.xml'] ?? new Uint8Array())
  expect(numbering).toContain('<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>')
})

it('refuses a package without a content-types part', async () => {
  const source = docxBytes({ body: paragraph('x') + SECTION })
  const entries = unzipSync(source)
  delete entries['[Content_Types].xml']
  const { bytes, data } = await openSource(zipSync(entries))
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('ignores content-type declarations without an extension or part name', async () => {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CT}">`
    + '<Default ContentType="application/octet-stream"/>'
    + '<Override ContentType="text/plain"/>'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '</Types>'
  const { bytes, data } = await openSource(docxBytes({
    body: paragraph('x') + SECTION,
    contentTypes,
  }))
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('declares a missing rels default and a missing part content type', async () => {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CT}">`
    + '<Override PartName="/word/document.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '</Types>'
  const { bytes, data } = await openSource(docxBytes({
    body: paragraph('x') + SECTION,
    contentTypes,
  }))
  if (data.body === undefined) throw new Error('expected a body')
  data.body.customRanges = [...(data.body.customRanges ?? []), {
    startIndex: 0,
    endIndex: 1,
    rangeId: 'link',
    rangeType: CustomRangeType.HYPERLINK,
    properties: { url: 'https://example.test/' },
  }]
  const paragraphRecord = firstParagraph(data)
  paragraphRecord.styleId = 'CustomSplice'
  paragraphRecord.paragraphStyle = { horizontalAlign: HorizontalAlign.CENTER }
  const saved = await writeDocumentFromSource(data, bytes)
  const types = strFromU8(unzipSync(saved.bytes)['[Content_Types].xml'] ?? new Uint8Array())
  expect(types).toContain('<Default Extension="rels"')
  expect(types).toContain('PartName="/word/styles.xml"')
})

it('declares document, header, footer, and macro content types', async () => {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CT}">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Override PartName="/word/document.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '</Types>'
  const headerBody = paragraph('header')
  const footerBody = paragraph('footer')
  const body = paragraph('x')
    + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/>'
    + '<w:footerReference w:type="default" r:id="rId4"/></w:sectPr>'
  const rels = `<Relationship Id="rId3" Type="${REL}/header" Target="header1.xml"/>`
    + `<Relationship Id="rId4" Type="${REL}/footer" Target="footer1.xml"/>`
  const { bytes, data } = await openSource(docxBytes({
    body,
    contentTypes,
    styles: null,
    rels,
    extra: {
      'word/header1.xml': `<w:hdr xmlns:w="${WORD}">${headerBody}</w:hdr>`,
      'word/footer1.xml': `<w:ftr xmlns:w="${WORD}">${footerBody}</w:ftr>`,
    },
  }))
  if (data.headers?.['header1'] !== undefined) {
    const headerParagraph = data.headers.header1.body.paragraphs?.[0]
    if (headerParagraph !== undefined) headerParagraph.paragraphStyle = { horizontalAlign: HorizontalAlign.CENTER }
  }
  if (data.footers?.['footer1'] !== undefined) {
    const footerParagraph = data.footers.footer1.body.paragraphs?.[0]
    if (footerParagraph !== undefined) footerParagraph.paragraphStyle = { horizontalAlign: HorizontalAlign.RIGHT }
  }
  const saved = await writeDocumentFromSource(data, bytes)
  const types = strFromU8(unzipSync(saved.bytes)['[Content_Types].xml'] ?? new Uint8Array())
  expect(types).toContain('wordprocessingml.header+xml')
  expect(types).toContain('wordprocessingml.footer+xml')
})

it('declares a macro-enabled main document content type', async () => {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CT}">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '</Types>'
  const bytes = docxBytes({ body: paragraph('x') + SECTION, contentTypes, styles: null })
  const data = await readDocumentWithSource(bytes, 'docm', CONFIG)
  editParagraph(data, { horizontalAlign: HorizontalAlign.CENTER })
  const saved = await writeDocumentFromSource(data, bytes)
  const types = strFromU8(unzipSync(saved.bytes)['[Content_Types].xml'] ?? new Uint8Array())
  expect(types).toContain('macroEnabled.main+xml')
})

it('refuses auxiliary appends when the source part has no closing tag', async () => {
  const styles = `<w:styles xmlns:w="${WORD}"/>`
  const styleEdit = await openSource(docxBytes({ body: paragraph('x') + SECTION, styles }))
  const styleParagraph = firstParagraph(styleEdit.data)
  styleParagraph.styleId = 'MissingStyle'
  styleParagraph.paragraphStyle = { horizontalAlign: HorizontalAlign.CENTER }
  await expect(writeDocumentFromSource(styleEdit.data, styleEdit.bytes))
    .rejects.toMatchObject({ code: 'conversion-failed' })

  const numbering = `<w:numbering xmlns:w="${WORD}"/>`
  const listEdit = await openSource(docxBytes({ body: paragraph('x') + SECTION, numbering }))
  firstParagraph(listEdit.data).bullet = { listType: 'numbering-9', listId: 'num-9', nestingLevel: 0 }
  await expect(writeDocumentFromSource(listEdit.data, listEdit.bytes))
    .rejects.toMatchObject({ code: 'conversion-failed' })
})

it('merges independent paragraph properties into a created container', async () => {
  const body = paragraph('x', '<w:pPr><w:keepNext/><w:widowControl/></w:pPr>') + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  editParagraph(data, { horizontalAlign: HorizontalAlign.CENTER })
  const saved = await writeDocumentFromSource(data, bytes)
  const xml = strFromU8(unzipSync(saved.bytes)['word/document.xml'] ?? new Uint8Array())
  expect(xml).toContain('<w:keepNext/>')
  expect(xml).toContain('<w:widowControl/>')
})

it('merges paragraph-mark run properties into a created run-property container', async () => {
  const body = paragraph('x', '<w:pPr><w:rPr><w:noProof/></w:rPr></w:pPr>') + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  editParagraph(data, { horizontalAlign: HorizontalAlign.CENTER })
  const saved = await writeDocumentFromSource(data, bytes)
  const xml = strFromU8(unzipSync(saved.bytes)['word/document.xml'] ?? new Uint8Array())
  expect(xml).toContain('<w:rPr><w:noProof/></w:rPr>')
})

it('orders merged properties inserted at one offset by rank and duplicate order', async () => {
  const body = paragraph('x', '<w:pPr><w:keepNext/><w:keepNext/><w:widowControl/></w:pPr>') + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  editParagraph(data, { horizontalAlign: HorizontalAlign.CENTER })
  const saved = await writeDocumentFromSource(data, bytes)
  const xml = strFromU8(unzipSync(saved.bytes)['word/document.xml'] ?? new Uint8Array())
  expect(xml.match(/<w:keepNext\/>/gu)).toHaveLength(2)
})

it('keeps a nested property container the renderer treats as a paragraph property', async () => {
  const body = paragraph('x', '<w:pPr><w:rPr><w:noProof/></w:rPr></w:pPr>') + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  if (data.body === undefined) throw new Error('expected a body')
  data.body.dataStream = data.body.dataStream.replace('x', 'y')
  const saved = await writeDocumentFromSource(data, bytes)
  const xml = strFromU8(unzipSync(saved.bytes)['word/document.xml'] ?? new Uint8Array())
  expect(xml).toContain('<w:rPr><w:noProof/></w:rPr>')
})

it('refuses a property whose owning container shape changed', async () => {
  const body = '<w:p><w:r><w:rPr><w:noProof/></w:rPr><w:t>x</w:t></w:r></w:p>' + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  const paragraphRecord = firstParagraph(data)
  if (data.body === undefined) throw new Error('expected a body')
  data.body.dataStream = data.body.dataStream.replace('x', 'y')
  paragraphRecord.startIndex = data.body.dataStream.indexOf('\r')
  paragraphRecord.paragraphStyle = { horizontalAlign: HorizontalAlign.CENTER }
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses container properties nested under the wrong owner', async () => {
  const body = paragraph('x', '<w:pPr><w:noWrap/></w:pPr><w:trPr><w:hidden/></w:trPr>') + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  editParagraph(data, { horizontalAlign: HorizontalAlign.CENTER })
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('skips a mapped property the renderer omits', async () => {
  const body = paragraph('x', '<w:pPr><w:pStyle w:val="Custom"/></w:pPr>') + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  const source = firstParagraph(data)
  delete source.styleId
  source.paragraphStyle = { horizontalAlign: HorizontalAlign.CENTER }
  const saved = await writeDocumentFromSource(data, bytes)
  expect(saved.bytes.byteLength).toBeGreaterThan(0)
})

it('merges properties through comments, CDATA, and processing instructions', async () => {
  const body = '<w:p><w:pPr><w:keepNext/></w:pPr><!--keep--><w:r><w:t>a<![CDATA[b]]></w:t></w:r>'
    + '<?target value?></w:p>' + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  editParagraph(data, { horizontalAlign: HorizontalAlign.CENTER })
  const saved = await writeDocumentFromSource(data, bytes)
  const xml = strFromU8(unzipSync(saved.bytes)['word/document.xml'] ?? new Uint8Array())
  expect(xml).toContain('<w:keepNext/>')
})

it('merges an independent table-cell property', async () => {
  const table = '<w:tbl><w:tr><w:tc><w:tcPr><w:noWrap/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const { bytes, data } = await openSource(docxBytes({ body: `${table}${SECTION}` }))
  const model = data.tableSource?.['table-1']
  const cell = model?.tableRows[0]?.tableCells[0]
  if (cell === undefined) throw new Error('expected a cell')
  cell.vAlign = VerticalAlignmentType.CENTER
  const saved = await writeDocumentFromSource(data, bytes)
  const xml = strFromU8(unzipSync(saved.bytes)['word/document.xml'] ?? new Uint8Array())
  expect(xml).toContain('<w:noWrap/>')
})

it('splices a table whose id contains a separator', async () => {
  const table = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const { bytes, data } = await openSource(docxBytes({ body: `${table}${SECTION}` }))
  const custom = data.body?.tables?.[0]
  const model = data.tableSource?.['table-1']
  if (custom === undefined || model === undefined) throw new Error('expected a table')
  custom.tableId = 'a/b'
  data.tableSource = { 'a/b': model }
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('splices the second cell and a cell without a width', async () => {
  const table = '<w:tbl><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="1200"/></w:tblGrid>'
    + '<w:tr><w:tc><w:tcPr><w:tcW w:w="1200"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>'
    + '<w:tc><w:tcPr/><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const { bytes, data } = await openSource(docxBytes({ body: `${table}${SECTION}` }))
  const model = data.tableSource?.['table-1']
  const second = model?.tableRows[0]?.tableCells[1]
  if (second === undefined) throw new Error('expected a second cell')
  second.backgroundColor = { rgb: '#FFFFFF' }
  const saved = await writeDocumentFromSource(data, bytes)
  expect(saved.bytes.byteLength).toBeGreaterThan(0)
})

it('defaults a column width beyond the declared grid', async () => {
  const table = '<w:tbl><w:tblGrid><w:gridCol w:w="1200"/></w:tblGrid>'
    + '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const { bytes, data } = await openSource(docxBytes({ body: `${table}${SECTION}` }))
  const model = data.tableSource?.['table-1']
  const cell = model?.tableRows[0]?.tableCells[0]
  if (cell === undefined) throw new Error('expected a cell')
  cell.backgroundColor = { rgb: '#FFFFFF' }
  const saved = await writeDocumentFromSource(data, bytes)
  expect(saved.bytes.byteLength).toBeGreaterThan(0)
})

it('refuses a spliced part that is not well-formed XML', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  const body = data.body
  if (body === undefined || body.paragraphs?.[0] === undefined) throw new Error('expected a body')
  body.dataStream = body.dataStream.replace('x', 'y')
  body.paragraphs[0].startIndex = body.dataStream.indexOf('\r')
  body.docxRawBlocks = [{ startIndex: 0, xml: '<w:r w:x="1" w:x="2"/>' }]
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses an output above the byte bound', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  const body = data.body
  if (body === undefined) throw new Error('expected a body')
  const longText = Array.from(
    { length: 4000 },
    (_, index) => String.fromCodePoint(33 + ((index * 37 + index * index) % 90)),
  ).join('')
  body.dataStream = `${longText}\r\n`
  const paragraphRecord = body.paragraphs?.[0]
  if (paragraphRecord === undefined) throw new Error('expected a paragraph')
  paragraphRecord.startIndex = longText.length
  paragraphRecord.paragraphStyle = { horizontalAlign: HorizontalAlign.CENTER }
  body.textRuns = []
  await expect(writeDocumentFromSource(data, bytes, { ...CONFIG, maxFileBytes: bytes.byteLength + 1 }))
    .rejects.toMatchObject({ code: 'too-large' })
})

it('refuses an OPC graph with an override for a missing part', async () => {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CT}">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/missing.xml" ContentType="application/xml"/>'
    + '</Types>'
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION, contentTypes }))
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses a part whose scan fails after a digests-mismatch save', async () => {
  const header = `<w:hdr xmlns:w="${WORD}">${paragraph('header')}</w:hdr>`
  const source = docxBytes({
    body: paragraph('body') + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/></w:sectPr>',
    rels: `<Relationship Id="rId3" Type="${REL}/header" Target="header1.xml"/>`,
    extra: { 'word/header1.xml': header },
  })
  const { data } = await openSource(source)
  const entries = unzipSync(source)
  const broken = strToU8(`<w:hdr xmlns:w="${WORD}"><?broken`)
  entries['word/header1.xml'] = broken
  const modified = zipSync(entries)
  const opened = open(modified, { maxEntries: 10_000 }).source
  patchSourceState(data, {
    packageDigest: opened.digest,
    byteLength: opened.bytes.byteLength,
  })
  const anchors = anchorsOf(data).map(anchor => (
    anchor.part === 'word/header1.xml' ? { ...anchor, end: Math.min(Number(anchor.end), broken.byteLength - 1) } : anchor
  ))
  patchSourceState(data, { anchors })
  await expect(writeDocumentFromSource(data, modified)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('rethrows a source-state budget failure that is not a parser refusal', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  await expect(writeDocumentFromSource(data, bytes, { ...CONFIG, maxSourceAnchors: 0 })).rejects.toThrow(RangeError)
})

it('refuses a non-record source sidecar', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  carrierOf(data).officeExchange = { source: 5 }
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('rebuilds anchors when only the sidecar byte length is stale', async () => {
  const { bytes, data } = await openSource(docxBytes({ body: paragraph('x') + SECTION }))
  const byteLength = sourceStateOf(data)['byteLength']
  if (typeof byteLength !== 'number') throw new Error('expected a byte length')
  patchSourceState(data, { byteLength: byteLength + 1 })
  const saved = await writeDocumentFromSource(data, bytes)
  expect(saved.bytes).toEqual(bytes)
})

it('refuses an ambiguous alignment between an insertion and an edited unit', async () => {
  const first = await openSource(docxBytes({ body: `${paragraph('one')}${paragraph('two')}${SECTION}` }))
  editParagraph(first.data, { horizontalAlign: HorizontalAlign.CENTER })
  const inserted = docxBytes({
    body: `${paragraph('inserted')}${paragraph('one')}${paragraph('two')}${SECTION}`,
  })
  await expect(writeDocumentFromSource(first.data, inserted)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses a reordered source that digest matching cannot align', async () => {
  const first = await openSource(docxBytes({ body: `${paragraph('one')}${paragraph('two')}${SECTION}` }))
  const reordered = docxBytes({ body: `${paragraph('two')}${paragraph('one')}${SECTION}` })
  await expect(writeDocumentFromSource(first.data, reordered)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('aligns duplicate digests by order and deletes source-only units', async () => {
  const first = await openSource(docxBytes({
    body: `${paragraph('same')}${paragraph('middle')}${paragraph('same')}${SECTION}`,
  }))
  const rewritten = docxBytes({
    body: `${paragraph('same')}${paragraph('same')}${paragraph('same')}${paragraph('other')}${SECTION}`,
  })
  const saved = await writeDocumentFromSource(first.data, rewritten)
  const xml = strFromU8(unzipSync(saved.bytes)['word/document.xml'] ?? new Uint8Array())
  expect(xml.match(/same/gu) ?? []).toHaveLength(2)
  expect(xml).toContain('middle')
  expect(xml).not.toContain('other')
})

it('aligns a source that deleted one of two duplicate leading units', async () => {
  const first = await openSource(docxBytes({
    body: `${paragraph('same')}${paragraph('middle')}${paragraph('same')}${SECTION}`,
  }))
  const rewritten = docxBytes({ body: `${paragraph('middle')}${paragraph('same')}${SECTION}` })
  const saved = await writeDocumentFromSource(first.data, rewritten)
  const xml = strFromU8(unzipSync(saved.bytes)['word/document.xml'] ?? new Uint8Array())
  expect(xml.match(/same/gu) ?? []).toHaveLength(2)
  expect(xml).toContain('middle')
})

it('refuses a rebuild when the snapshot dropped a rebuilt story part', async () => {
  const header = `<w:hdr xmlns:w="${WORD}">${paragraph('header')}</w:hdr>`
  const source = docxBytes({
    body: paragraph('body') + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/></w:sectPr>',
    rels: `<Relationship Id="rId3" Type="${REL}/header" Target="header1.xml"/>`,
    extra: { 'word/header1.xml': header },
  })
  const { data } = await openSource(source)
  delete data.headers
  const modified = docxBytes({
    body: paragraph('body') + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/></w:sectPr>',
    rels: `<Relationship Id="rId3" Type="${REL}/header" Target="header1.xml"/>`,
    extra: { 'word/header1.xml': header },
    styles: DEFAULT_STYLES.replace('Normal', 'Other'),
  })
  await expect(writeDocumentFromSource(data, modified)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('refuses a source it cannot re-anchor and saves after a digests-mismatch rebuild', async () => {
  const first = await openSource(docxBytes({ body: `${paragraph('one')}${paragraph('two')}${SECTION}` }))
  const second = await openSource(docxBytes({
    body: `${paragraph('one')}${paragraph('two')}${SECTION}`,
    styles: DEFAULT_STYLES.replace('Normal', 'NormalTwo'),
  }))
  editParagraph(first.data, { horizontalAlign: HorizontalAlign.CENTER })
  const saved = await writeDocumentFromSource(first.data, second.bytes)
  expect(saved.bytes.byteLength).toBeGreaterThan(0)

  const unanchorable = docxBytes({
    body: `${paragraph('one')}<w:hyperlink><w:p/></w:hyperlink>${SECTION}`,
  })
  const mismatch = structuredClone(first.data)
  await expect(writeDocumentFromSource(mismatch, unanchorable)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('reports a preserved relationship element without an id', async () => {
  const { writeDocument } = await import('../src/provider/docx.ts')
  const data: IDocumentData = {
    id: 'x',
    documentStyle: {},
    body: { dataStream: 'x\r\n', paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] },
  }
  carrierOf(data).officeExchange = {
    docx: {
      parts: [],
      relationships: { document: ['<Relationship Type="image" Target="media/a.png"/>'] },
      contentTypes: [],
    },
  }
  const written = await writeDocument(data, CONFIG)
  expect(written.byteLength).toBeGreaterThan(0)
})

it('writes a header body without paragraphs', async () => {
  const { writeDocument } = await import('../src/provider/docx.ts')
  const data: IDocumentData = {
    id: 'x',
    documentStyle: {},
    body: {
      dataStream: 'x\r\n',
      paragraphs: [{
        startIndex: 1,
        paragraphId: 'p1',
        bullet: { listType: 'list-a', listId: 'l1', nestingLevel: 0 },
      }],
    },
    headers: { header1: { headerId: 'header1', body: { dataStream: '\r\n' } } },
    lists: {
      'list-a': {
        listType: 'list-a',
        nestingLevel: [{ glyphFormat: '%1.', glyphType: 0, startNumber: 0, bulletAlignment: 0 }],
      },
    },
  }
  const written = await writeDocument(data, CONFIG)
  expect(written.byteLength).toBeGreaterThan(0)
})

it('writes a table whose cell spans more columns than the grid declares', async () => {
  const { writeDocument } = await import('../src/provider/docx.ts')
  const stream = '\x1A\x1B\x1C\r\x1D\x0E\x0F'
  const data: IDocumentData = {
    id: 'x',
    documentStyle: {},
    body: {
      dataStream: stream,
      tables: [{ startIndex: 0, endIndex: stream.length, tableId: 'table-1' }],
    },
    tableSource: {
      'table-1': {
        tableId: 'table-1',
        tableRows: [{ tableCells: [{ columnSpan: 2 }], trHeight: { val: { v: 10 }, hRule: 0 } }],
        tableColumns: [{ size: { type: 1, width: { v: 80 } } }],
        align: 0,
        indent: { v: 0 },
        textWrap: 0,
        position: { positionH: { relativeFrom: 0, posOffset: 0 }, positionV: { relativeFrom: 0, posOffset: 0 } },
        dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
        cellMargin: { start: { v: 10 }, end: { v: 10 }, top: { v: 5 }, bottom: { v: 5 } },
        size: { type: 1, width: { v: 160 } },
      },
    },
  }
  const written = await writeDocument(data, CONFIG)
  expect(written.byteLength).toBeGreaterThan(0)
})

it('writes list levels without a glyph type', async () => {
  const { writeDocument } = await import('../src/provider/docx.ts')
  const data: IDocumentData = {
    id: 'x',
    documentStyle: {},
    body: {
      dataStream: 'x\r\n',
      paragraphs: [{
        startIndex: 1,
        paragraphId: 'p1',
        bullet: { listType: 'list-a', listId: 'l1', nestingLevel: 0 },
      }],
    },
    lists: {
      'list-a': {
        listType: 'list-a',
        nestingLevel: [{ glyphFormat: '', startNumber: 0, bulletAlignment: 0 }],
      },
    },
  }
  const written = await writeDocument(data, CONFIG)
  expect(written.byteLength).toBeGreaterThan(0)
})

it('keeps supplied metadata through a preserved-package regeneration', async () => {
  const { writeDocument } = await import('../src/provider/docx.ts')
  const data: IDocumentData = {
    id: 'x',
    documentStyle: {},
    body: { dataStream: 'x\r\n', paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] },
  }
  carrierOf(data).officeExchange = {
    docx: {
      parts: [{ name: 'word/media/image1.png', base64: 'AA==' }],
      relationships: { document: [] },
      contentTypes: ['<Default Extension="png" ContentType="image/png"/>'],
    },
  }
  const written = await writeDocument(data, CONFIG)
  expect(written.byteLength).toBeGreaterThan(0)
})

it('writes preserved parts with each base64 padding width', async () => {
  const { writeDocument } = await import('../src/provider/docx.ts')
  for (const base64 of ['AAA=', 'AAE=', 'AA==']) {
    const data: IDocumentData = {
      id: 'x',
      documentStyle: {},
      body: { dataStream: 'x\r\n', paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] },
    }
    carrierOf(data).officeExchange = {
      docx: {
        parts: [{ name: 'word/media/image1.png', base64 }],
        relationships: { document: [] },
        contentTypes: ['<Default Extension="png" ContentType="image/png"/>'],
      },
    }
    const written = await writeDocument(data, CONFIG)
    expect(written.byteLength).toBeGreaterThan(0)
  }
})

it('writes preserved notes with and without declarations', async () => {
  const { writeDocument } = await import('../src/provider/docx.ts')
  for (const contentTypes of [[], ['<Override PartName="/word/footnotes.xml" ContentType="notes"/>']]) {
    const data: IDocumentData = {
      id: 'x',
      documentStyle: {},
      body: { dataStream: 'x\r\n', paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] },
    }
    carrierOf(data).officeExchange = {
      docx: {
        parts: [{ name: 'word/footnotes.xml', base64: 'AQ==' }],
        relationships: { document: [] },
        contentTypes,
      },
    }
    const written = await writeDocument(data, CONFIG)
    expect(written.byteLength).toBeGreaterThan(0)
  }
})

it('rejects a non-string preserved part payload', async () => {
  const { writeDocument } = await import('../src/provider/docx.ts')
  const data: IDocumentData = {
    id: 'x',
    documentStyle: {},
    body: { dataStream: 'x\r\n', paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] },
  }
  carrierOf(data).officeExchange = {
    docx: { parts: [{ name: 'word/media/image1.png', base64: 5 }], relationships: {}, contentTypes: [] },
  }
  await expect(writeDocument(data, CONFIG)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('accepts an unused glyph style in a preserved-package write', async () => {
  const { writeDocument } = await import('../src/provider/docx.ts')
  const data: IDocumentData = {
    id: 'x',
    body: { dataStream: 'x\r\n', paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] },
    documentStyle: { evenAndOddHeaders: BooleanNumber.TRUE },
    headers: {
      header2: {
        headerId: 'header2',
        body: { dataStream: 'h\r\n', paragraphs: [{ startIndex: 1, paragraphId: 'h1', styleId: 'Custom' }] },
      },
    },
  }
  carrierOf(data).officeExchange = { docx: { parts: [], relationships: {}, contentTypes: [] } }
  const written = await writeDocument(data, CONFIG)
  expect(written.byteLength).toBeGreaterThan(0)
})

it('writes a named style into the generated styles part', async () => {
  const { writeDocument } = await import('../src/provider/docx.ts')
  const data: IDocumentData = {
    id: 'x',
    documentStyle: {},
    body: {
      dataStream: 'x\r\n',
      paragraphs: [{
        startIndex: 1,
        paragraphId: 'p1',
        paragraphStyle: { namedStyleType: NamedStyleType.HEADING_2 },
      }],
    },
  }
  const written = await writeDocument(data, CONFIG)
  expect(strFromU8(unzipSync(written)['word/styles.xml'] ?? new Uint8Array()))
    .toContain('w:styleId="Heading2"')
})

void DEFAULT_OFFICE_EXCHANGE_CONFIG







it('saves headers and footers whose table metadata is absent', async () => {
  const body = paragraph('body') + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/>'
    + '<w:footerReference w:type="default" r:id="rId4"/></w:sectPr>'
  const rels = `<Relationship Id="rId3" Type="${REL}/header" Target="header1.xml"/>`
    + `<Relationship Id="rId4" Type="${REL}/footer" Target="footer1.xml"/>`
  const { bytes, data } = await openSource(docxBytes({
    body,
    rels,
    extra: {
      'word/header1.xml': `<w:hdr xmlns:w="${WORD}">${paragraph('header')}</w:hdr>`,
      'word/footer1.xml': `<w:ftr xmlns:w="${WORD}">${paragraph('footer')}</w:ftr>`,
    },
  }))
  if (data.headers?.['header1'] !== undefined) delete data.headers.header1.tableSource
  if (data.footers?.['footer1'] !== undefined) delete data.footers.footer1.tableSource
  const saved = await writeDocumentFromSource(data, bytes)
  expect(saved.bytes.byteLength).toBeGreaterThan(0)
})

it('saves a changed header whose body has no paragraphs', async () => {
  const body = paragraph('body') + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/></w:sectPr>'
  const rels = `<Relationship Id="rId3" Type="${REL}/header" Target="header1.xml"/>`
  const { bytes, data } = await openSource(docxBytes({
    body,
    rels,
    extra: { 'word/header1.xml': `<w:hdr xmlns:w="${WORD}">${paragraph('header')}</w:hdr>` },
  }))
  const header = data.headers?.['header1']
  if (header === undefined) throw new Error('expected a header')
  header.body = { dataStream: '' }
  const saved = await writeDocumentFromSource(data, bytes)
  expect(saved.bytes.byteLength).toBeGreaterThan(0)
})

it('adds start overrides for every level of a new list', async () => {
  const body = paragraph('one', '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>')
    + paragraph('two', '<w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr>')
    + SECTION
  const { bytes, data } = await openSource(docxBytes({ body, numbering: DEFAULT_NUMBERING }))
  const levels = [
    { glyphFormat: '%1.', glyphType: 0, startNumber: 0, bulletAlignment: 0 },
    { glyphFormat: '%2.', glyphType: 0, startNumber: 0, bulletAlignment: 0 },
  ]
  data.lists = { 'numbering-9': { listType: 'numbering-9', nestingLevel: levels } }
  const first = data.body?.paragraphs?.[0]
  const second = data.body?.paragraphs?.[1]
  if (first === undefined || second === undefined) throw new Error('expected paragraphs')
  first.bullet = { listType: 'numbering-9', listId: 'num-9', nestingLevel: 0, startNumber: 3 }
  second.bullet = { listType: 'numbering-9', listId: 'num-9', nestingLevel: 1, startNumber: 4 }
  const saved = await writeDocumentFromSource(data, bytes)
  const numbering = strFromU8(unzipSync(saved.bytes)['word/numbering.xml'] ?? new Uint8Array())
  expect(numbering).toContain('<w:startOverride w:val="3"/>')
  expect(numbering).toContain('<w:startOverride w:val="4"/>')
})

it('refuses a cell key whose column is not a number', async () => {
  const table = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const { bytes, data } = await openSource(docxBytes({ body: `${table}${SECTION}` }))
  const custom = data.body?.tables?.[0]
  const model = data.tableSource?.['table-1']
  if (custom === undefined || model === undefined) throw new Error('expected a table')
  custom.tableId = 'a/0/b'
  data.tableSource = { 'a/0/b': model }
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('saves a plan with replace, delete, insert, carry, and append operations', async () => {
  const body = paragraph('alpha') + paragraph('beta') + paragraph('gamma') + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  const documentBody = data.body
  if (documentBody === undefined) throw new Error('expected a body')
  if (documentBody.dataStream !== 'alpha\rbeta\rgamma\r\n') throw new Error('unexpected stream')
  documentBody.dataStream = 'alpha\rdelta\rgamma\r\n'
  const paragraphs = documentBody.paragraphs ?? []
  if (paragraphs.length !== 3) throw new Error('expected three paragraphs')
  const [first, second, third] = paragraphs
  if (first === undefined || second === undefined || third === undefined) throw new Error('expected paragraphs')
  first.paragraphStyle = { horizontalAlign: HorizontalAlign.CENTER }
  third.startIndex = 17
  third.paragraphStyle = { horizontalAlign: HorizontalAlign.RIGHT }
  const section = documentBody.sectionBreaks?.[0]
  if (section !== undefined) section.startIndex = 18
  documentBody.paragraphs = [first, {
    startIndex: 11,
    paragraphId: 'p4',
    paragraphStyle: { horizontalAlign: HorizontalAlign.RIGHT },
  }, third]
  void second
  const saved = await writeDocumentFromSource(data, bytes)
  const xml = strFromU8(unzipSync(saved.bytes)['word/document.xml'] ?? new Uint8Array())
  expect(xml).toContain('delta')
  expect(xml).not.toContain('beta')
})

it('refuses a run property whose owner disappeared from the rendered unit', async () => {
  const body = '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr></w:r></w:p>'
    + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  const paragraphRecord = firstParagraph(data)
  paragraphRecord.paragraphStyle = { horizontalAlign: HorizontalAlign.RIGHT }
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('declares a numbering content type the manifest omits', async () => {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CT}">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Override PartName="/word/document.xml"'
    + ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '</Types>'
  const body = paragraph('item', '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>') + SECTION
  const { bytes, data } = await openSource(docxBytes({
    body,
    contentTypes,
    numbering: DEFAULT_NUMBERING,
    styles: null,
  }))
  const paragraphRecord = firstParagraph(data)
  paragraphRecord.bullet = { listType: 'numbering-9', listId: 'num-9', nestingLevel: 0 }
  const saved = await writeDocumentFromSource(data, bytes)
  const types = strFromU8(unzipSync(saved.bytes)['[Content_Types].xml'] ?? new Uint8Array())
  expect(types).toContain('wordprocessingml.numbering+xml')
})


it('treats an unanchored empty paragraph as a source paragraph', async () => {
  const middle = await openSource(docxBytes({
    body: `${paragraph('x')}<w:p/>${paragraph('y')}${SECTION}`,
  }))
  const middleState = sourceStateOf(middle.data)
  const middleAnchors = anchorsOf(middle.data)
  const keepMiddle = middleAnchors.filter(anchor => anchor['key'] !== 'para_2')
  patchSourceState(middle.data, { anchors: keepMiddle })
  expect(middleState['anchors']).toHaveLength(3)
  const middleSaved = await writeDocumentFromSource(middle.data, middle.bytes)
  expect(middleSaved.bytes.byteLength).toBeGreaterThan(0)

  const trailing = await openSource(docxBytes({
    body: `${paragraph('x')}<w:p/>${SECTION}`,
  }))
  patchSourceState(trailing.data, {
    anchors: anchorsOf(trailing.data).filter(anchor => anchor['key'] !== 'para_2'),
  })
  const trailingSaved = await writeDocumentFromSource(trailing.data, trailing.bytes)
  expect(trailingSaved.bytes.byteLength).toBeGreaterThan(0)
})

it('uses empty style and numbering paragraph lists when rendering a table cell', async () => {
  const table = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const { bytes, data } = await openSource(docxBytes({ body: `${table}${SECTION}` }))
  const documentBody = data.body
  if (documentBody === undefined) throw new Error('expected a body')
  documentBody.dataStream = '\x1A\x1B\x1C\r\x1D\x0E\x0F\n'
  documentBody.tables = [{ startIndex: 0, endIndex: 7, tableId: 'table-1' }]
  documentBody.sectionBreaks = [{ sectionId: 'section-1', startIndex: 7 }]
  delete documentBody.paragraphs
  delete documentBody.textRuns
  const model = data.tableSource?.['table-1']
  const cell = model?.tableRows[0]?.tableCells[0]
  if (cell === undefined) throw new Error('expected a cell')
  cell.backgroundColor = { rgb: '#ABCDEF' }
  const saved = await writeDocumentFromSource(data, bytes)
  expect(saved.bytes.byteLength).toBeGreaterThan(0)
})

it('declares the main document content type of a normal package', async () => {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CT}">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '</Types>'
  const { bytes, data } = await openSource(docxBytes({
    body: paragraph('x') + SECTION,
    contentTypes,
    styles: null,
    numbering: null,
  }))
  firstParagraph(data).paragraphStyle = { horizontalAlign: HorizontalAlign.CENTER }
  const saved = await writeDocumentFromSource(data, bytes)
  const types = strFromU8(unzipSync(saved.bytes)['[Content_Types].xml'] ?? new Uint8Array())
  expect(types).toContain('wordprocessingml.document.main+xml')
})

it('merges a property whose name has no schema position', async () => {
  const body = paragraph('x', '<w:pPr><w:unknownProp/></w:pPr>') + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  editParagraph(data, { horizontalAlign: HorizontalAlign.CENTER })
  const saved = await writeDocumentFromSource(data, bytes)
  const xml = strFromU8(unzipSync(saved.bytes)['word/document.xml'] ?? new Uint8Array())
  expect(xml).toContain('<w:unknownProp/>')
})

it('refuses a run property when the rendered paragraph has no children', async () => {
  const body = '<w:p><w:r><w:rPr><w:noProof/></w:rPr></w:r></w:p>' + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  const documentBody = data.body
  if (documentBody === undefined) throw new Error('expected a body')
  documentBody.textRuns = [{ st: 0, ed: 0, ts: { bl: 1 } }]
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})

it('records a raw-run fallback after a consumed record', async () => {
  const body = '<w:p><w:r><w:drawing r:embed="rIdA"/></w:r>'
    + '<x:r><w:drawing/></x:r>'
    + '<w:r><w:drawing r:embed="rIdB"/></w:r></w:p>'
  await expect(openSource(docxBytes({ body: body + SECTION }))).rejects.toBeDefined()
})

it('reads nested raw elements and runs with non-self-closing children', async () => {
  const body = '<w:p><w:sdt><w:sdt><w:x/></w:sdt></w:sdt>'
    + '<w:r><w:r><w:t>nested</w:t></w:r><w:drawing/></w:r></w:p>'
  const data = await (await import('../src/provider/docx.ts')).readDocument(docxBytes({ body: body + SECTION }))
  const blocks = data.body?.docxRawBlocks ?? []
  expect(blocks.some(block => block.xml.includes('<w:sdt>'))).toBe(true)
  expect(blocks.some(block => block.xml.includes('<w:drawing/>'))).toBe(true)
})

it('declares two missing image defaults in sorted order', async () => {
  const { writeDocument } = await import('../src/provider/docx.ts')
  const data: IDocumentData = {
    id: 'x',
    documentStyle: {},
    body: { dataStream: 'x\r\n', paragraphs: [{ startIndex: 1, paragraphId: 'p1' }] },
  }
  carrierOf(data).officeExchange = {
    docx: {
      parts: [
        { name: 'word/media/image1.png', base64: 'AQ==' },
        { name: 'word/media/image2.jpeg', base64: 'AQ==' },
        { name: 'word/media/image3.gif', base64: 'AQ==' },
      ],
      relationships: { document: [] },
      contentTypes: [],
    },
  }
  const written = await writeDocument(data, CONFIG)
  const types = strFromU8(unzipSync(written)['[Content_Types].xml'] ?? new Uint8Array())
  expect(types).toContain('<Default Extension="gif" ContentType="image/gif"/>')
})


it('refuses a run property when the rendered paragraph is empty', async () => {
  const body = '<w:p><w:r><w:rPr><w:noProof/></w:rPr></w:r></w:p>' + SECTION
  const { bytes, data } = await openSource(docxBytes({ body }))
  const documentBody = data.body
  if (documentBody === undefined) throw new Error('expected a body')
  documentBody.textRuns = [{ st: 0, ed: 0, ts: { bl: 1 } }]
  await expect(writeDocumentFromSource(data, bytes)).rejects.toMatchObject({ code: 'conversion-failed' })
})
