/** Provider dispatch branch coverage: dispatch modes, sidecar validation, and disposition normalization. */
import { readFile } from 'node:fs/promises'
import { Context } from '@forwork/cordis'
import { LocaleType } from '@univerjs/core'
import { expect, it } from 'vitest'
import LocalOfficeExchange, { saveDispositionOf } from '../src/provider/index.ts'
import { readDocumentWithSource } from '../src/provider/docx.ts'
import { open } from '../src/package/ooxml-package.ts'
import type { OfficeExchange } from '../src/service.ts'
import type { IDocumentData, IWorkbookData } from '@univerjs/core'

const FIXTURE = new URL('./fixtures/preserve-image.docx', import.meta.url)

interface ExchangeCarrier extends IDocumentData {
  officeExchange?: unknown
}

async function exchange(): Promise<OfficeExchange> {
  const ctx = new Context()
  const plugin = ctx.plugin(LocalOfficeExchange)
  await plugin
  return ctx.officeExchange
}

async function preservedData(): Promise<{ data: IDocumentData; bytes: Uint8Array }> {
  const bytes = new Uint8Array(await readFile(FIXTURE))
  const data = await readDocumentWithSource(bytes, 'docx')
  return { data, bytes }
}

function carrierOf(data: IDocumentData): ExchangeCarrier {
  return data
}

it('sorts save disposition names with every ASCII comparison', () => {
  const disposition = saveDispositionOf({
    schemaVersion: 1,
    verbatimCount: 0,
    spliced: { count: 3, names: ['b', 'a', 'a'] },
    added: { count: 3, names: ['a', 'c', 'b'] },
    removed: { count: 1, names: ['c'] },
    nameLimit: 256,
  })
  expect(disposition.spliced.names).toEqual(['a', 'a', 'b'])
  expect(disposition.added.names).toEqual(['a', 'b', 'c'])
  expect(disposition.removed.names).toEqual(['c'])
})

it('rejects a snapshot whose officeExchange is not an object', async () => {
  const office = await exchange()
  const data: IDocumentData = { id: 'x', documentStyle: {}, body: { dataStream: 'x\r\n' } }
  carrierOf(data).officeExchange = 5
  await expect(office.write({ kind: 'document', data }, 'docx')).rejects.toMatchObject({
    code: 'conversion-failed',
  })
})

it('rejects a non-record source sidecar without a source package', async () => {
  const office = await exchange()
  const data: IDocumentData = { id: 'x', documentStyle: {}, body: { dataStream: 'x\r\n' } }
  carrierOf(data).officeExchange = { source: 5 }
  await expect(office.write({ kind: 'document', data }, 'docx')).rejects.toMatchObject({
    code: 'conversion-failed',
  })
})

it('rejects source parts entries that are not strings or records', async () => {
  const office = await exchange()
  const data: IDocumentData = { id: 'x', documentStyle: {}, body: { dataStream: 'x\r\n' } }
  carrierOf(data).officeExchange = {
    source: {
      schemaVersion: 1,
      format: 'docx',
      packageDigest: 'a'.repeat(64),
      byteLength: 1,
      anchorMode: 'per-unit',
      sourceParts: ['word/document.xml', 5, { part: 6 }],
      anchors: [],
    },
  }
  await expect(office.write({ kind: 'document', data }, 'docx')).rejects.toMatchObject({
    code: 'conversion-failed',
  })
})

it('treats a non-record source sidecar as absent when a source package is supplied', async () => {
  const office = await exchange()
  const { data, bytes } = await preservedData()
  const source = open(bytes, { maxEntries: 10_000 }).source
  carrierOf(data).officeExchange = { source: 5 }
  await expect(office.write({ kind: 'document', data }, 'docx', source)).rejects.toMatchObject({
    code: 'conversion-failed',
  })
})

it('drops preserved parts when the source package is absent', async () => {
  const office = await exchange()
  const { data } = await preservedData()
  const carrier = carrierOf(data)
  carrier.officeExchange = {
    ...(carrier.officeExchange ?? {}),
    docx: { parts: [], relationships: {}, contentTypes: [] },
  }
  const written = await office.write({ kind: 'document', data }, 'docx')
  expect(written.bytes.byteLength).toBeGreaterThan(0)
})

it('reports a format mismatch before touching the sidecar', async () => {
  const office = await exchange()
  const data: IWorkbookData = {
    id: 'x',
    name: 'Book',
    appVersion: '1.0.0',
    locale: LocaleType.EN_US,
    styles: {},
    sheets: {},
    sheetOrder: [],
  }
  await expect(office.write({ kind: 'workbook', data }, 'docx')).rejects.toMatchObject({
    code: 'conversion-failed',
  })
})
