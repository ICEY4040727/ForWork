/** Legacy binary conversion: target mapping, failure paths, and a real engine round-trip. */
import { createHash, randomUUID } from 'node:crypto'
import { existsSync } from 'node:fs'
import { mkdtemp, readFile, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { ConversionError, createConverter } from '@deepseek-ai/libreoffice-kit'
import { expect, it } from 'vitest'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG, resolveConfig } from '../src/config.ts'
import { OfficeExchangeError } from '../src/errors.ts'
import { convertLegacyFile, legacyTarget } from '../src/provider/legacy.ts'
import { readWorkbook } from '../src/provider/xlsx.ts'

const FIXTURE = fileURLToPath(new URL('./fixtures/legacy.xls', import.meta.url))

let availability: Promise<string | undefined> | undefined

/**
 * Report why the conversion test must skip, or `undefined` when the bundled
 * LibreOffice engine can be created. A non-engine failure is not a skip.
 */
function unavailableReason(): Promise<string | undefined> {
  availability ??= probeEngine()
  return availability
}

async function probeEngine(): Promise<string | undefined> {
  try {
    const converter = await createConverter()
    await converter.dispose()
    return undefined
  } catch (error) {
    if (error instanceof ConversionError && error.code === 'unavailable') return error.message
    throw error
  }
}

function hashOf(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

async function rejectionOf(operation: Promise<unknown>): Promise<unknown> {
  try {
    await operation
  } catch (error) {
    return error
  }
  throw new Error('expected the operation to reject')
}

it('maps each legacy extension to its editable target', () => {
  expect(legacyTarget('report.doc')).toEqual({ targetPath: 'report.docx', format: 'docx' })
  expect(legacyTarget('budget.xls')).toEqual({ targetPath: 'budget.xlsx', format: 'xlsx' })
  expect(legacyTarget('deck.ppt')).toEqual({ targetPath: 'deck.pptx', format: 'pptx' })
})

it('keeps the directory and stem while replacing only the final extension', () => {
  expect(legacyTarget('/work/2026/report.doc')).toEqual({
    targetPath: '/work/2026/report.docx',
    format: 'docx',
  })
  expect(legacyTarget('C:\\work\\Q1.Review.DOC')).toEqual({
    targetPath: 'C:\\work\\Q1.Review.docx',
    format: 'docx',
  })
})

it('matches extensions case-insensitively', () => {
  expect(legacyTarget('REPORT.DOC')?.format).toBe('docx')
  expect(legacyTarget('Budget.Xls')?.format).toBe('xlsx')
  expect(legacyTarget('Deck.PPT')?.format).toBe('pptx')
})

it('returns undefined for unknown, editable, and extensionless paths', () => {
  expect(legacyTarget('notes.txt')).toBeUndefined()
  expect(legacyTarget('report.docx')).toBeUndefined()
  expect(legacyTarget('report')).toBeUndefined()
  expect(legacyTarget('dir.v2/report')).toBeUndefined()
})

it('rejects an unsupported extension before creating the output directory', async () => {
  const outputDir = join(tmpdir(), `dsh-office-exchange-unsupported-${randomUUID()}`)
  const error = await rejectionOf(convertLegacyFile('report.docx', outputDir, { config: DEFAULT_OFFICE_EXCHANGE_CONFIG }))
  if (!(error instanceof OfficeExchangeError)) throw error
  expect(error.code).toBe('unsupported-format')
  expect(existsSync(outputDir)).toBe(false)
})

it('rejects an already-aborted request without starting a conversion', async () => {
  const reason = new Error('cancelled before conversion')
  await expect(
    convertLegacyFile(FIXTURE, join(tmpdir(), 'dsh-office-exchange-aborted'), {
      config: DEFAULT_OFFICE_EXCHANGE_CONFIG,
      signal: AbortSignal.abort(reason),
    }),
  ).rejects.toBe(reason)
})

it('converts the .xls fixture to .xlsx without modifying the source', { timeout: 180_000 }, async (context) => {
  const reason = await unavailableReason()
  if (reason !== undefined) {
    context.skip(`LibreOffice unavailable: ${reason}`)
    return
  }
  const sourceBefore = new Uint8Array(await readFile(FIXTURE))
  const workspace = await mkdtemp(join(tmpdir(), 'dsh-office-exchange-legacy-'))
  try {
    const outputDir = join(workspace, 'converted')
    const outputPath = await convertLegacyFile(FIXTURE, outputDir, { config: DEFAULT_OFFICE_EXCHANGE_CONFIG })
    expect(outputPath).toBe(join(outputDir, 'legacy.xlsx'))
    const info = await stat(outputPath)
    expect(info.isFile()).toBe(true)
    expect(info.size).toBeGreaterThan(0)

    const workbook = await readWorkbook(new Uint8Array(await readFile(outputPath)))
    const sheetId = workbook.sheetOrder.find(id => workbook.sheets[id]?.name === 'Legacy')
    if (sheetId === undefined) throw new Error('converted workbook is missing the Legacy sheet')
    expect(workbook.sheets[sheetId]?.cellData?.[0]?.[0]?.v).toBe('Item')
    expect(workbook.sheets[sheetId]?.cellData?.[1]?.[0]?.v).toBe('Widget')

    const sourceAfter = new Uint8Array(await readFile(FIXTURE))
    expect(hashOf(sourceAfter)).toBe(hashOf(sourceBefore))
  } finally {
    await rm(workspace, { recursive: true, force: true })
  }
})

it('fails an over-budget conversion as too-large without modifying the source', { timeout: 180_000 }, async (context) => {
  const reason = await unavailableReason()
  if (reason !== undefined) {
    context.skip(`LibreOffice unavailable: ${reason}`)
    return
  }
  const sourceBefore = new Uint8Array(await readFile(FIXTURE))
  const workspace = await mkdtemp(join(tmpdir(), 'dsh-office-exchange-deadline-'))
  try {
    const error = await rejectionOf(
      convertLegacyFile(FIXTURE, join(workspace, 'converted'), { config: resolveConfig({ maxConversionMs: 1 }) }),
    )
    if (!(error instanceof OfficeExchangeError)) throw error
    expect(error.code).toBe('too-large')
    expect(error.message).toContain('maxConversionMs')
    expect(hashOf(new Uint8Array(await readFile(FIXTURE)))).toBe(hashOf(sourceBefore))
  } finally {
    await rm(workspace, { recursive: true, force: true })
  }
})
