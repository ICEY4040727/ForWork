import { createHash } from 'node:crypto'
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { verifyBrandAssets, type FingerprintRow } from './verify-brand-assets.ts'

const DIGEST = '0'.repeat(64)

const temps: string[] = []

afterEach(() => {
  for (const root of temps.splice(0)) rmSync(root, { recursive: true, force: true })
})

/**
 * Build a minimal repository tree with one brand asset and a manifest.
 * @param asset - asset bytes.
 * @param row - manifest fields to override.
 * @param fingerprints - retired patterns to include.
 * @returns the temporary root.
 */
function fixture(
  asset: string,
  row: { width?: number; height?: number; sha256?: string } = {},
  fingerprints: FingerprintRow[] = [],
): string {
  const root = mkdtempSync(join(tmpdir(), 'forwork-brand-gate-'))
  temps.push(root)
  mkdirSync(join(root, 'assets/brand'), { recursive: true })
  mkdirSync(join(root, 'scripts'), { recursive: true })
  writeFileSync(join(root, 'assets/brand/marker.svg'), asset)
  writeFileSync(join(root, 'scripts/brand-assets.manifest.json'), JSON.stringify({
    version: 1,
    assets: [{
      path: 'assets/brand/marker.svg',
      kind: 'svg-mark',
      surface: 'brand-master',
      width: row.width ?? 1,
      height: row.height ?? 1,
      sha256: row.sha256 ?? createHash('sha256').update(asset).digest('hex'),
      generator: 'authored',
      required: true,
    }],
    evidence: [],
    fingerprints,
  }))
  return root
}

const MARKER = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1"><path d="M0 0h1v1H0z"/></svg>\n'

describe('verify-brand-assets', () => {
  it('accepts the committed manifest and evidence set', () => {
    const result = verifyBrandAssets()
    expect(result.errors).toEqual([])
    expect(result.assets).toBeGreaterThan(20)
    expect(result.evidence).toBe(13)
  })

  it('rejects a stale hash and stale dimensions', () => {
    const root = fixture(MARKER, { sha256: DIGEST, width: 9, height: 9 })
    const errors = verifyBrandAssets({ root }).errors
    expect(errors).toContain('assets/brand/marker.svg: dimensions 9x9 are stale; expected 1x1')
    expect(errors).toContain('assets/brand/marker.svg: SHA-256 is stale; run pnpm run verify-brand-assets --write')
  })

  it('rejects an unreferenced brand asset', () => {
    const root = fixture(MARKER)
    writeFileSync(join(root, 'assets/brand/extra.svg'), MARKER)
    expect(verifyBrandAssets({ root }).errors).toContain('assets/brand/extra.svg: brand asset is not referenced by the manifest')
  })

  it('rejects a retired SVG contour even when the manifest row is current', () => {
    const retired = readFileSync(new URL('./fixtures/brand/retired-whale.svg', import.meta.url), 'utf8')
    const root = fixture(retired, {}, [{
      artifact: 'deepseek-whale-mark',
      kind: 'svg-path',
      value: 'M22.9168 1.43018C22.6713 1.31018',
      replacement: 'assets/brand/forwork-mark.svg',
    }])
    expect(verifyBrandAssets({ root }).errors).toContain(
      'assets/brand/marker.svg: contains retired deepseek-whale-mark; expected assets/brand/forwork-mark.svg',
    )
  })

  it('rejects a retired raster by hash', () => {
    const root = fixture(MARKER, {}, [{
      artifact: 'deepseek-home-atlas',
      kind: 'sha256',
      value: createHash('sha256').update(MARKER).digest('hex'),
      replacement: 'assets/brand/forwork-mark.svg',
    }])
    expect(verifyBrandAssets({ root }).errors).toContain(
      'assets/brand/marker.svg: matches retired deepseek-home-atlas; expected assets/brand/forwork-mark.svg',
    )
  })

  it('rejects a missing manifest file', () => {
    const root = fixture(MARKER)
    rmSync(join(root, 'scripts/brand-assets.manifest.json'))
    expect(verifyBrandAssets({ root }).errors).toEqual(['scripts/brand-assets.manifest.json: missing'])
  })
})
