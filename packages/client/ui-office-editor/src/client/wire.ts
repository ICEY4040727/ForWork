/**
 * Conversions between the Remote's JSON projection and the Univer native
 * snapshot interfaces. The wire value is JSON-safe by contract, so each
 * direction validates the envelope at this boundary instead of asserting it.
 */
import type { IDocumentData, ILanguagePack, IWorkbookData } from '@univerjs/core'
import type { JsonValue } from '@deepseek-ai/dsh-util-values'

/**
 * Admit a preset's locale default export as a language pack.
 * @param value - the imported English dictionary.
 * @returns the dictionary the preset expects.
 */
export function localePack(value: Record<string, unknown>): ILanguagePack {
  return value as ILanguagePack
}

/**
 * Project a native snapshot onto the Remote's lossless JSON value.
 * @param value - the value `save()` returned from the editor.
 * @returns the JSON projection the Remote accepts.
 * @throws when the snapshot carries a value JSON cannot represent.
 */
export function jsonValueOf(value: unknown): JsonValue {
  if (value === null || typeof value === 'boolean' || typeof value === 'string') return value
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new Error('ui-office-editor: the snapshot contains a non-finite number')
    return value
  }
  if (Array.isArray(value)) return value.map(member => jsonValueOf(member))
  if (typeof value === 'object') {
    const prototype: unknown = Object.getPrototypeOf(value)
    if (prototype !== Object.prototype && prototype !== null) {
      throw new Error('ui-office-editor: the snapshot contains a non-JSON object')
    }
    const record: Record<string, JsonValue> = {}
    for (const [key, member] of Object.entries(value) as [string, unknown][]) {
      if (member === undefined) continue
      record[key] = jsonValueOf(member)
    }
    return record
  }
  throw new Error(`ui-office-editor: the snapshot contains a non-JSON ${typeof value} value`)
}

/**
 * Restore the native workbook value a sheets runtime loads.
 * @param value - the Remote's JSON projection.
 * @returns the workbook data object.
 * @throws when the envelope lacks the identity fields the editor dereferences.
 */
export function workbookData(value: JsonValue): IWorkbookData {
  if (!isWorkbookData(value)) throw new Error('ui-office-editor: the snapshot is not a workbook data object')
  return value
}

/**
 * Restore the native document value a docs runtime loads.
 * @param value - the Remote's JSON projection.
 * @returns the document data object.
 * @throws when the envelope lacks the identity fields the editor dereferences.
 */
export function documentData(value: JsonValue): IDocumentData {
  if (!isDocumentData(value)) throw new Error('ui-office-editor: the snapshot is not a document data object')
  return value
}

/** Whether a value carries the required workbook identity and containers. */
function isWorkbookData(value: unknown): value is IWorkbookData {
  if (!isRecord(value)) return false
  return typeof value['id'] === 'string'
    && typeof value['name'] === 'string'
    && typeof value['appVersion'] === 'string'
    && typeof value['locale'] === 'string'
    && isRecord(value['styles'])
    && isRecord(value['sheets'])
    && Array.isArray(value['sheetOrder'])
    && value['sheetOrder'].every(entry => typeof entry === 'string')
}

/** Whether a value carries the required document identity and containers. */
function isDocumentData(value: unknown): value is IDocumentData {
  if (!isRecord(value)) return false
  return typeof value['id'] === 'string'
    && isRecord(value['documentStyle'])
    && (value['body'] === undefined || isRecord(value['body']))
}

/** Whether a value is a plain object the mapping helpers can index. */
function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
