/** Format classification: extensions map to supported formats and editability. */
import { expect, it } from 'vitest'
import { formatOf, isEditableFormat, isLegacyFormat } from '../src/format.ts'

it('classifies supported extensions case-insensitively', () => {
  expect(formatOf('report.DOCX')).toBe('docx')
  expect(formatOf('budget.xlsx')).toBe('xlsx')
  expect(formatOf('deck.pptx')).toBe('pptx')
  expect(formatOf('old.doc')).toBe('doc')
  expect(formatOf('old.XLS')).toBe('xls')
  expect(formatOf('old.ppt')).toBe('ppt')
})

it('rejects unknown, missing, and extensionless paths', () => {
  expect(formatOf('notes.txt')).toBeUndefined()
  expect(formatOf('archive.docx.zip')).toBeUndefined()
  expect(formatOf('README')).toBeUndefined()
  expect(formatOf('')).toBeUndefined()
})

it('separates directly editable formats from legacy formats', () => {
  expect(isEditableFormat('docx')).toBe(true)
  expect(isEditableFormat('xlsx')).toBe(true)
  expect(isEditableFormat('pptx')).toBe(false)
  expect(isEditableFormat('doc')).toBe(false)
  expect(isLegacyFormat('doc')).toBe(true)
  expect(isLegacyFormat('xls')).toBe(true)
  expect(isLegacyFormat('ppt')).toBe(true)
  expect(isLegacyFormat('docx')).toBe(false)
})
