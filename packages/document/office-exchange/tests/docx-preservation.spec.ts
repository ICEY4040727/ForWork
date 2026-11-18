/** Opaque DOCX preservation: images and note parts survive an edit-and-save byte-for-byte. */
import { readFile } from 'node:fs/promises'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import type { IDocumentData } from '@univerjs/core'
import { expect, it } from 'vitest'
import { detectUnsupportedFeatures } from '../src/detect.ts'
import { readDocument, writeDocument } from '../src/provider/docx.ts'

const DOCUMENT_XML = 'word/document.xml'

interface PreservedPartView {
  readonly name: string
  readonly base64: string
}

interface PreservedPackageView {
  readonly parts?: readonly PreservedPartView[]
  readonly relationships?: {
    readonly document?: readonly string[]
    readonly stories?: Readonly<Record<string, readonly string[]>>
  }
  readonly contentTypes?: readonly string[]
}

function preservationOf(data: IDocumentData): PreservedPackageView | undefined {
  return (data as IDocumentData & { officeExchange?: { docx?: PreservedPackageView } }).officeExchange?.docx
}

function entryText(entries: Record<string, Uint8Array>, name: string): string {
  const content = entries[name]
  return content === undefined ? '' : strFromU8(content)
}

it('preserves an embedded JPEG and its relationship byte-for-byte', async () => {
  const source = new Uint8Array(await readFile(new URL('./fixtures/preserve-image.docx', import.meta.url)))
  expect(detectUnsupportedFeatures(source)).toEqual([])
  const data = await readDocument(source)
  const preserved = preservationOf(data)
  expect(preserved?.parts?.map(part => part.name)).toContain('word/media/image1.jpeg')
  expect(preserved?.relationships?.document?.join('')).toContain('Id="rId5"')
  expect(preserved?.contentTypes?.join('')).toContain('image/jpeg')
  const blocks = data.body?.docxRawBlocks ?? []
  expect(blocks).toHaveLength(1)
  expect(blocks[0]?.xml).toContain('<w:drawing>')
  expect(blocks[0]?.startIndex).toBe('Before image '.length)
  expect(data.body?.dataStream).toBe('Before image after image.\r\n')

  const written = unzipSync(await writeDocument(data))
  const sourceEntries = unzipSync(source)
  expect(written['word/media/image1.jpeg']).toEqual(sourceEntries['word/media/image1.jpeg'])
  expect(entryText(written, DOCUMENT_XML)).toContain(blocks[0]?.xml ?? '')
  expect(entryText(written, 'word/_rels/document.xml.rels')).toContain('Id="rId5"')
  expect(entryText(written, '[Content_Types].xml')).toContain('ContentType="image/jpeg"')

  const reopened = await readDocument(await writeDocument(data))
  expect(reopened.body?.docxRawBlocks).toEqual(data.body?.docxRawBlocks)
  expect(reopened.body?.dataStream).toBe(data.body?.dataStream)
})

it('preserves a linked image and its external relationship', async () => {
  const source = new Uint8Array(await readFile(new URL('./fixtures/preserve-external-image.docx', import.meta.url)))
  expect(detectUnsupportedFeatures(source)).toEqual([])
  const data = await readDocument(source)
  expect(preservationOf(data)?.parts ?? []).toEqual([])
  expect(preservationOf(data)?.relationships?.document?.join('')).toContain('TargetMode="External"')
  const blocks = data.body?.docxRawBlocks ?? []
  expect(blocks[0]?.xml).toContain('r:link="rId6"')

  const written = unzipSync(await writeDocument(data))
  expect(entryText(written, DOCUMENT_XML)).toContain(blocks[0]?.xml ?? '')
  const rels = entryText(written, 'word/_rels/document.xml.rels')
  expect(rels).toContain('Id="rId6"')
  expect(rels).toContain('https://example.test/linked.jpeg')
  expect(rels).toContain('TargetMode="External"')
  expect(written['word/media/image1.jpeg']).toBeUndefined()

  const reopened = await readDocument(await writeDocument(data))
  expect(reopened.body?.docxRawBlocks).toEqual(data.body?.docxRawBlocks)
})

it('preserves footnote and endnote parts, references, and relationships', async () => {
  const source = new Uint8Array(await readFile(new URL('./fixtures/preserve-notes.docx', import.meta.url)))
  expect(detectUnsupportedFeatures(source)).toEqual([])
  const data = await readDocument(source)
  const preserved = preservationOf(data)
  expect(preserved?.parts?.map(part => part.name)).toEqual([
    'word/endnotes.xml',
    'word/footnotes.xml',
  ])
  expect(preserved?.contentTypes?.join('')).toContain('/word/footnotes.xml')
  expect(preserved?.contentTypes?.join('')).toContain('/word/endnotes.xml')
  const blocks = data.body?.docxRawBlocks ?? []
  expect(blocks).toHaveLength(2)
  expect(blocks[0]?.xml).toContain('<w:footnoteReference w:id="1"/>')
  expect(blocks[1]?.xml).toContain('<w:endnoteReference w:id="2"/>')
  expect(data.body?.dataStream).toBe('Body text with an endnote\r\n')

  const written = unzipSync(await writeDocument(data))
  const sourceEntries = unzipSync(source)
  expect(written['word/footnotes.xml']).toEqual(sourceEntries['word/footnotes.xml'])
  expect(written['word/endnotes.xml']).toEqual(sourceEntries['word/endnotes.xml'])
  expect(entryText(written, DOCUMENT_XML)).toContain(blocks[0]?.xml ?? '')
  expect(entryText(written, DOCUMENT_XML)).toContain(blocks[1]?.xml ?? '')
  const rels = entryText(written, 'word/_rels/document.xml.rels')
  expect(rels).toContain('Id="rId7"')
  expect(rels).toContain('Id="rId8"')
  expect(entryText(written, '[Content_Types].xml')).toContain('/word/footnotes.xml')
  expect(entryText(written, '[Content_Types].xml')).toContain('/word/endnotes.xml')

  const reopened = await readDocument(await writeDocument(data))
  expect(reopened.body?.docxRawBlocks).toEqual(data.body?.docxRawBlocks)
  expect(preservationOf(reopened)?.parts).toEqual(preserved?.parts)
  expect(reopened.body?.dataStream).toBe(data.body?.dataStream)
})

it('preserves a header image and its story relationship', async () => {
  const drawing = '<w:r><w:drawing><wp:inline xmlns:wp="urn:probe">'
    + '<a:blip xmlns:a="urn:probe" r:embed="rId1"/></wp:inline></w:drawing></w:r>'
  const relationship = '<Relationship Id="rId1"'
    + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
    + ' Target="media/header.png"/>'
  const files = {
    '[Content_Types].xml': strToU8('<Types/>'),
    'word/_rels/document.xml.rels': strToU8('<Relationships>'
      + '<Relationship Id="rId2"'
      + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"'
      + ' Target="header1.xml"/></Relationships>'),
    'word/document.xml': strToU8('<w:document'
      + ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      + '<w:body><w:p/><w:sectPr>'
      + '<w:headerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/>'
      + '</w:sectPr></w:body></w:document>'),
    'word/header1.xml': strToU8('<w:hdr'
      + ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      + `<w:p>${drawing}</w:p></w:hdr>`),
    'word/_rels/header1.xml.rels': strToU8(`<Relationships>${relationship}</Relationships>`),
    'word/media/header.png': new Uint8Array([1, 2, 3, 4]),
  }
  const source = zipSync(files)
  expect(detectUnsupportedFeatures(source)).toEqual([])
  const data = await readDocument(source)
  const blocks = data.headers?.header1?.body.docxRawBlocks ?? []
  expect(blocks).toHaveLength(1)
  expect(blocks[0]?.xml).toContain('r:embed="rId1"')
  expect(preservationOf(data)?.relationships?.stories?.header1?.join('')).toContain('Id="rId1"')

  const written = unzipSync(await writeDocument(data))
  expect(entryText(written, 'word/header1.xml')).toContain(blocks[0]?.xml ?? '')
  expect(entryText(written, 'word/_rels/header1.xml.rels')).toContain(relationship)
  expect(written['word/media/header.png']).toEqual(files['word/media/header.png'])
  const reopened = await readDocument(await writeDocument(data))
  expect(reopened.headers?.header1?.body.docxRawBlocks).toEqual(data.headers?.header1?.body.docxRawBlocks)
})
