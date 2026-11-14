/** Unsupported-feature detection over in-memory OOXML packages. */
import { strToU8, zipSync } from 'fflate'
import { expect, it } from 'vitest'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG } from '../src/config.ts'
import { detectUnsupportedFeatures } from '../src/detect.ts'
import type { OfficeExchangeErrorCode } from '../src/errors.ts'
import { OfficeExchangeError } from '../src/errors.ts'

const DOCUMENT_XML = 'word/document.xml'
const PLAIN_DOCUMENT = '<w:document><w:body><w:p/></w:body></w:document>'

function packageOf(entries: Readonly<Record<string, string | Uint8Array>>): Uint8Array {
  const files: Record<string, Uint8Array> = {}
  for (const [name, content] of Object.entries(entries)) {
    files[name] = typeof content === 'string' ? strToU8(content) : content
  }
  return zipSync(files)
}

function codeOfThrown(run: () => void): OfficeExchangeErrorCode {
  try {
    run()
  } catch (error) {
    if (!(error instanceof OfficeExchangeError)) throw error
    return error.code
  }
  throw new Error('expected the call to throw')
}

it('reports no features for a clean document', () => {
  const bytes = packageOf({
    '[Content_Types].xml': '<Types/>',
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/styles.xml': '<styles/>',
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([])
})

it('detects macro parts', () => {
  for (const path of ['word/vbaProject.bin', 'xl/vbaProject.bin', 'ppt/vbaProject.bin']) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: new Uint8Array([0]) })
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'macro', detail: path }])
  }
})

it('detects tracked changes in the document body', () => {
  const ins = packageOf({ [DOCUMENT_XML]: '<w:document><w:ins w:id="1"/></w:document>' })
  expect(detectUnsupportedFeatures(ins)).toEqual([{ kind: 'tracked-changes', detail: DOCUMENT_XML }])

  const del = packageOf({ [DOCUMENT_XML]: '<w:document><w:del w:id="2"/></w:document>' })
  expect(detectUnsupportedFeatures(del)).toEqual([{ kind: 'tracked-changes', detail: DOCUMENT_XML }])
})

it('detects comment parts across Word, Excel, and PowerPoint', () => {
  for (const path of ['word/comments.xml', 'xl/comments1.xml', 'ppt/comments/comment1.xml']) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: '<comments/>' })
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'comment', detail: path }])
  }
})

it('detects chart parts', () => {
  for (const path of ['xl/charts/chart1.xml', 'ppt/charts/chart1.xml']) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: '<chart/>' })
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'chart', detail: path }])
  }
})

it('detects drawing and VML shape parts', () => {
  const drawing = packageOf({
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/drawings/drawing1.xml': '<drawing/>',
  })
  expect(detectUnsupportedFeatures(drawing)).toEqual([
    { kind: 'drawing', detail: 'word/drawings/drawing1.xml' },
  ])

  const shape = packageOf({
    [DOCUMENT_XML]: PLAIN_DOCUMENT,
    'word/vmlDrawing1.vml': '<vml/>',
  })
  expect(detectUnsupportedFeatures(shape)).toEqual([
    { kind: 'shape', detail: 'word/vmlDrawing1.vml' },
  ])
})

it('detects header and footer parts', () => {
  for (const path of ['word/header1.xml', 'word/footer2.xml']) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: '<hdr/>' })
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'header-footer', detail: path }])
  }
})

it('detects embedded objects', () => {
  for (const path of ['word/embeddings/oleObject1.bin', 'xl/embeddings/oleObject1.bin']) {
    const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT, [path]: new Uint8Array([0]) })
    expect(detectUnsupportedFeatures(bytes)).toEqual([{ kind: 'embedded-object', detail: path }])
  }
})

it('reports every kind once, ordered by the fixed kind order', () => {
  const bytes = packageOf({
    'word/embeddings/oleObject1.bin': new Uint8Array([0]),
    'word/header1.xml': '<hdr/>',
    'word/vmlDrawing1.vml': '<vml/>',
    'ppt/charts/chart1.xml': '<chart/>',
    'word/comments.xml': '<comments/>',
    [DOCUMENT_XML]: '<w:document><w:ins w:id="1"/></w:document>',
    'xl/vbaProject.bin': new Uint8Array([0]),
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([
    { kind: 'macro', detail: 'xl/vbaProject.bin' },
    { kind: 'tracked-changes', detail: DOCUMENT_XML },
    { kind: 'comment', detail: 'word/comments.xml' },
    { kind: 'chart', detail: 'ppt/charts/chart1.xml' },
    { kind: 'shape', detail: 'word/vmlDrawing1.vml' },
    { kind: 'header-footer', detail: 'word/header1.xml' },
    { kind: 'embedded-object', detail: 'word/embeddings/oleObject1.bin' },
  ])
})

it('keeps only the first part per kind', () => {
  const bytes = packageOf({
    'xl/charts/chart1.xml': '<chart/>',
    'xl/charts/chart2.xml': '<chart/>',
  })
  expect(detectUnsupportedFeatures(bytes)).toEqual([
    { kind: 'chart', detail: 'xl/charts/chart1.xml' },
  ])
})

it('rejects input over maxFileBytes before parsing', () => {
  const bytes = new Uint8Array(32)
  expect(codeOfThrown(() => detectUnsupportedFeatures(bytes, {
    ...DEFAULT_OFFICE_EXCHANGE_CONFIG,
    maxFileBytes: 8,
  }))).toBe('too-large')
})

it('rejects a valid package larger than maxFileBytes', () => {
  const bytes = packageOf({ [DOCUMENT_XML]: PLAIN_DOCUMENT })
  expect(codeOfThrown(() => detectUnsupportedFeatures(bytes, {
    ...DEFAULT_OFFICE_EXCHANGE_CONFIG,
    maxFileBytes: bytes.byteLength - 1,
  }))).toBe('too-large')
})

it('rejects packages with too many entries', () => {
  const entries: Record<string, Uint8Array> = {}
  for (let index = 0; index <= 10_000; index += 1) {
    entries[`word/media/part${index}.xml`] = new Uint8Array(0)
  }
  expect(codeOfThrown(() => detectUnsupportedFeatures(zipSync(entries)))).toBe('too-large')
})

it('classifies non-zip input as conversion-failed', () => {
  expect(codeOfThrown(() => detectUnsupportedFeatures(strToU8('not a zip archive')))).toBe('conversion-failed')
  expect(codeOfThrown(() => detectUnsupportedFeatures(new Uint8Array(0)))).toBe('conversion-failed')
})
