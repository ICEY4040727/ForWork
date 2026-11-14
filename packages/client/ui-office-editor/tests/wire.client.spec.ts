/**
 * The JSON projection and native-snapshot validation used by the editor
 * runtimes. These helpers run without Univer, so the wire boundary stays under
 * the per-file coverage gate even though the runtime modules need a browser.
 */
import { describe, expect, it } from 'vitest'
import { documentData, jsonValueOf, localePack, workbookData } from '../src/client/wire.ts'

const workbook = {
  id: 'book',
  name: 'Budget',
  appVersion: '1.4.0',
  locale: 'enUS',
  styles: {},
  sheets: {},
  sheetOrder: ['sheet-1'],
}

const document = { id: 'doc', documentStyle: {}, body: { dataStream: 'hello' } }

describe('jsonValueOf', () => {
  it('projects every JSON primitive and container', () => {
    expect(jsonValueOf(null)).toBeNull()
    expect(jsonValueOf(true)).toBe(true)
    expect(jsonValueOf('text')).toBe('text')
    expect(jsonValueOf(2.5)).toBe(2.5)
    expect(jsonValueOf([1, 'two', null])).toEqual([1, 'two', null])
    expect(jsonValueOf({ a: 1, b: [false], c: { d: 'e' } })).toEqual({ a: 1, b: [false], c: { d: 'e' } })
  })

  it('skips undefined object members and accepts null-prototype records', () => {
    expect(jsonValueOf({ kept: 1, dropped: undefined })).toEqual({ kept: 1 })
    const bare = Object.create(null) as Record<string, unknown>
    bare['value'] = 1
    expect(jsonValueOf(bare)).toEqual({ value: 1 })
  })

  it('rejects values JSON cannot represent', () => {
    expect(() => jsonValueOf(Number.NaN)).toThrow('non-finite')
    expect(() => jsonValueOf(Number.POSITIVE_INFINITY)).toThrow('non-finite')
    expect(() => jsonValueOf(new Date())).toThrow('non-JSON object')
    expect(() => jsonValueOf(() => {})).toThrow('non-JSON function')
    expect(() => jsonValueOf(1n)).toThrow('non-JSON bigint')
    expect(() => jsonValueOf([undefined])).toThrow('non-JSON undefined')
  })
})

describe('workbookData', () => {
  it('admits a workbook envelope and rejects a missing identity field', () => {
    expect(workbookData(workbook)).toBe(workbook)
    expect(() => workbookData(null)).toThrow('not a workbook data object')
    expect(() => workbookData([])).toThrow('not a workbook data object')
    expect(() => workbookData({ ...workbook, id: 1 })).toThrow('not a workbook data object')
    expect(() => workbookData({ ...workbook, name: 1 })).toThrow('not a workbook data object')
    expect(() => workbookData({ ...workbook, appVersion: 1 })).toThrow('not a workbook data object')
    expect(() => workbookData({ ...workbook, locale: 1 })).toThrow('not a workbook data object')
    expect(() => workbookData({ ...workbook, styles: [] })).toThrow('not a workbook data object')
    expect(() => workbookData({ ...workbook, sheets: [] })).toThrow('not a workbook data object')
    expect(() => workbookData({ ...workbook, sheetOrder: 'sheet-1' })).toThrow('not a workbook data object')
    expect(() => workbookData({ ...workbook, sheetOrder: [1] })).toThrow('not a workbook data object')
  })
})

describe('documentData', () => {
  it('admits a document envelope and rejects a missing document style', () => {
    expect(documentData(document)).toBe(document)
    expect(documentData({ id: 'doc', documentStyle: {} })).toEqual({ id: 'doc', documentStyle: {} })
    expect(() => documentData(null)).toThrow('not a document data object')
    expect(() => documentData({ id: 1, documentStyle: {} })).toThrow('not a document data object')
    expect(() => documentData({ id: 'doc' })).toThrow('not a document data object')
    expect(() => documentData({ id: 'doc', documentStyle: {}, body: 3 })).toThrow('not a document data object')
  })
})

describe('localePack', () => {
  it('hands the preset dictionary through unchanged', () => {
    const pack = { general: { ok: 'OK' } }
    expect(localePack(pack)).toBe(pack)
  })
})
