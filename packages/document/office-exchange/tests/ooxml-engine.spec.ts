/**
 * LibreOffice import/render smoke for saved DOCX and XLSX outputs.
 *
 * This is an import/render smoke, not a no-repair proof: LibreOffice can repair
 * a malformed package silently and the kit reports no repair diagnostics, so a
 * successful PDF render does not show that the saved package opens without
 * repair. Desktop Word and Excel are unavailable on the release host, so no
 * matrix establishes that claim; it is a disclosed residual risk.
 */
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Context } from '@forwork/cordis'
import { ConversionError, createConverter } from '@deepseek-ai/libreoffice-kit'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import LocalOfficeExchange from '../src/index.ts'
import { open } from '../src/package/ooxml-package.ts'
import type { SourcePackage } from '../src/package/types.ts'

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url))
const DOCX_FIXTURE = join(FIXTURES, 'roundtrip.docx')
const XLSX_FIXTURE = join(FIXTURES, 'roundtrip.xlsx')

/** Per-render engine deadline, matching the legacy conversion test budget. */
const RENDER_TIMEOUT_MS = 180_000

/** Editable fixture formats this smoke saves and renders. */
type SmokeFormat = 'docx' | 'xlsx'

let ctx: Context
let workspace: string
let availability: Promise<string | undefined> | undefined

/** Report why the render smoke must skip, or `undefined` when the engine exists. */
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

async function sourcePackageOf(path: string): Promise<SourcePackage> {
  return open(new Uint8Array(await readFile(path)), { maxEntries: 10_000 }).source
}

/**
 * Save one fixture through the exchange writer with its source package, render
 * the saved bytes through the bundled engine, and assert the reported result and
 * the PDF signature.
 * @param format - Fixture and target format under test.
 */
async function renderSavedFixture(format: SmokeFormat): Promise<void> {
  const fixture = format === 'docx' ? DOCX_FIXTURE : XLSX_FIXTURE
  const exchange = ctx.officeExchange
  const read = await exchange.read(fixture)
  const written = await exchange.write(read.snapshot, format, await sourcePackageOf(fixture))
  expect(written.bytes.byteLength).toBeGreaterThan(0)

  const inputPath = join(workspace, `saved-roundtrip.${format}`)
  const outputPath = join(workspace, `rendered-roundtrip.${format}.pdf`)
  await writeFile(inputPath, written.bytes)

  const converter = await createConverter({ timeoutMs: RENDER_TIMEOUT_MS })
  try {
    const result = await converter.render({ inputPath, outputPath })
    expect(['native', 'wasm']).toContain(result.backend)
    expect(Array.isArray(result.missingFonts)).toBe(true)
    const pdf = new Uint8Array(await readFile(outputPath))
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-')
  } finally {
    await converter.dispose()
  }
}

describe('libreoffice-load-smoke', () => {
  beforeAll(async () => {
    ctx = new Context()
    await ctx.plugin(LocalOfficeExchange)
    workspace = await mkdtemp(join(tmpdir(), 'dsh-office-exchange-engine-'))
  })

  afterAll(async () => {
    await ctx.fiber.dispose()
    await rm(workspace, { recursive: true, force: true })
  })

  it('renders a saved DOCX as a PDF', { timeout: RENDER_TIMEOUT_MS }, async (context) => {
    const reason = await unavailableReason()
    if (reason !== undefined) {
      context.skip(`LibreOffice unavailable: ${reason}`)
      return
    }
    await renderSavedFixture('docx')
  })

  it('renders a saved XLSX as a PDF', { timeout: RENDER_TIMEOUT_MS }, async (context) => {
    const reason = await unavailableReason()
    if (reason !== undefined) {
      context.skip(`LibreOffice unavailable: ${reason}`)
      return
    }
    await renderSavedFixture('xlsx')
  })
})
