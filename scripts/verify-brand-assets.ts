/**
 * Verify that every shipped ForWork brand asset is owned, current, and free of
 * retired upstream artwork, and that the reviewed screenshot evidence covers
 * the documented surfaces.
 *
 * Check mode never writes. `--write` refreshes hashes and dimensions for rows
 * that already exist; it never creates a row for an unreferenced file.
 *
 * @module verify-brand-assets
 */

import { createHash } from 'node:crypto'
import { existsSync, globSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'

/** Repository root, resolved from this script. */
const ROOT = join(import.meta.dirname, '..')

/** The single brand inventory. */
const MANIFEST_PATH = 'scripts/brand-assets.manifest.json'

/** Image files that count as brand assets, by root pattern. */
const BRAND_PATTERNS = [
  'assets/brand/**/*.svg',
  'assets/brand/**/*.png',
  'assets/brand/**/*.ico',
  'assets/brand/**/*.webp',
  'apps/desktop/resources/icon*.svg',
  'apps/desktop/resources/icon*.png',
  'apps/desktop/resources/tray-windows.ico',
  'apps/desktop/installer/assets/brand*.png',
  'apps/desktop/installer/assets/uninstaller-sidebar.png',
  'apps/desktop/renderer/assets/welcome-brand.svg',
  'apps/web/public/icon-*.png',
  'apps/web/public/apple-touch-icon.png',
  'apps/web/public/favicon*.svg',
  'website/public/wordmark.svg',
  'website/public/favicon.svg',
  'packages/client/ui-brand-official/src/client/assets/**/*.svg',
  'packages/client/ui-brand-official/src/client/assets/**/*.png',
  'packages/client/ui-brand-official/src/client/assets/**/*.webp',
] as const

/** Reviewed screenshot files, by owner-local expected directory. */
const EVIDENCE_PATTERNS = [
  'apps/web/tests/expected/brand/*.png',
  'apps/desktop/tests/expected/brand/*.png',
  'website/tests/expected/brand/*.png',
] as const

/** Screenshot surfaces the plan promises; every id must be present. */
const REQUIRED_EVIDENCE = [
  'web-home-desktop-light',
  'web-home-desktop-dark',
  'web-home-hover',
  'web-home-mobile',
  'web-rail-collapsed',
  'web-icon-sheet',
  'desktop-welcome-light',
  'desktop-welcome-dark',
  'website-nav-desktop',
  'website-nav-mobile',
  'installer-light',
  'installer-dark',
  'installer-uninstaller',
] as const

/** A shipped brand asset row. */
export interface BrandAssetRow {
  path: string
  kind: string
  surface: string
  width: number
  height: number
  sha256: string
  generator: string
  required: boolean
}

/** A reviewed screenshot row. */
export interface EvidenceRow {
  id: string
  path: string
  surface: string
  viewport: { width: number; height: number }
  scale: number
  colorScheme: 'light' | 'dark'
  locale: string
  depicts: readonly string[]
  sha256: string
  capturedBy: string
}

/** A retired upstream graphic that must not return. */
export interface FingerprintRow {
  artifact: string
  kind: 'svg-path' | 'sha256'
  value: string
  replacement: string
}

/** The manifest shape. */
export interface BrandManifest {
  version: number
  assets: BrandAssetRow[]
  evidence: EvidenceRow[]
  fingerprints: FingerprintRow[]
}

/** Verification options. */
export interface VerifyOptions {
  /** Repository root override for tests. */
  root?: string
  /** Refresh existing rows instead of checking them. */
  write?: boolean
}

/** Result of one verification pass. */
export interface VerifyResult {
  errors: string[]
  assets: number
  evidence: number
}

/**
 * Normalize a glob result to a repository-relative POSIX path.
 * @param path - path returned by the glob walk.
 * @returns the normalized repository path.
 */
function normalized(path: string): string {
  return path.split('\\').join('/')
}

/**
 * List files matching a pattern set.
 * @param root - repository root.
 * @param patterns - glob patterns relative to the root.
 * @returns sorted unique normalized paths.
 */
function listFiles(root: string, patterns: readonly string[]): string[] {
  const found = new Set<string>()
  for (const pattern of patterns) {
    for (const match of globSync(pattern, { cwd: root })) {
      const path = normalized(match)
      if (existsSync(join(root, path)) && statSync(join(root, path)).isFile()) found.add(path)
    }
  }
  return [...found].sort()
}

/**
 * Hash one file.
 * @param root - repository root.
 * @param path - repository-relative file.
 * @returns lowercase SHA-256.
 */
export function fileHash(root: string, path: string): string {
  return createHash('sha256').update(readFileSync(join(root, path))).digest('hex')
}

/**
 * Read PNG dimensions.
 * @param bytes - PNG bytes.
 * @returns width and height.
 */
function pngDimensions(bytes: Buffer): { width: number; height: number } {
  if (bytes.length < 24 || bytes.toString('ascii', 1, 4) !== 'PNG') throw new Error('not a PNG')
  return { width: bytes.readUInt32BE(16), height: bytes.readUInt32BE(20) }
}

/**
 * Read the largest ICO entry dimensions.
 * @param bytes - ICO bytes.
 * @returns width and height.
 */
function icoDimensions(bytes: Buffer): { width: number; height: number } {
  const count = bytes.readUInt16LE(4)
  let width = 0
  let height = 0
  for (let index = 0; index < count; index += 1) {
    const at = 6 + index * 16
    const entryWidth = bytes.readUInt8(at) || 256
    const entryHeight = bytes.readUInt8(at + 1) || 256
    if (entryWidth > width) width = entryWidth
    if (entryHeight > height) height = entryHeight
  }
  return { width, height }
}

/**
 * Read WebP dimensions for the simple lossless/lossy/VP8X containers.
 * @param bytes - WebP bytes.
 * @returns width and height.
 */
function webpDimensions(bytes: Buffer): { width: number; height: number } {
  const chunk = bytes.toString('ascii', 12, 16)
  if (chunk === 'VP8X') {
    return {
      width: 1 + bytes.readUIntLE(24, 3),
      height: 1 + bytes.readUIntLE(27, 3),
    }
  }
  if (chunk === 'VP8 ') {
    return {
      width: bytes.readUInt16LE(26) & 0x3fff,
      height: bytes.readUInt16LE(28) & 0x3fff,
    }
  }
  if (chunk === 'VP8L') {
    const bits = bytes.readUInt32LE(21)
    return { width: (bits & 0x3fff) + 1, height: ((bits >> 14) & 0x3fff) + 1 }
  }
  throw new Error(`unsupported WebP chunk ${chunk}`)
}

/**
 * Read an image's pixel dimensions.
 * @param root - repository root.
 * @param path - repository-relative image.
 * @returns width and height.
 */
export function imageDimensions(root: string, path: string): { width: number; height: number } {
  const bytes = readFileSync(join(root, path))
  if (path.endsWith('.png')) return pngDimensions(bytes)
  if (path.endsWith('.ico')) return icoDimensions(bytes)
  if (path.endsWith('.webp')) return webpDimensions(bytes)
  if (path.endsWith('.svg')) {
    const text = bytes.toString('utf8')
    const viewBox = /viewBox="([^"]+)"/u.exec(text)?.[1]?.trim().split(/[\s,]+/u).map(Number)
    const boxWidth = viewBox?.[2]
    const boxHeight = viewBox?.[3]
    if (boxWidth !== undefined && boxHeight !== undefined && Number.isFinite(boxWidth) && Number.isFinite(boxHeight)) {
      return { width: Math.round(boxWidth), height: Math.round(boxHeight) }
    }
    const width = Number(/\bwidth="([\d.]+)"/u.exec(text)?.[1])
    const height = Number(/\bheight="([\d.]+)"/u.exec(text)?.[1])
    if (Number.isFinite(width) && Number.isFinite(height)) return { width, height }
    throw new Error(`${path}: missing viewBox and numeric width/height`)
  }
  throw new Error(`${path}: unsupported image type`)
}

/**
 * Verify or refresh the brand manifest.
 * @param options - repository root and write mode.
 * @returns errors plus the verified counts.
 */
export function verifyBrandAssets(options: VerifyOptions = {}): VerifyResult {
  const root = options.root ?? ROOT
  const manifestPath = join(root, MANIFEST_PATH)
  if (!existsSync(manifestPath)) {
    return { errors: [`${MANIFEST_PATH}: missing`], assets: 0, evidence: 0 }
  }
  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8')) as BrandManifest
  const errors: string[] = []
  const seen = new Set<string>()

  for (const asset of manifest.assets) {
    if (seen.has(asset.path)) errors.push(`${asset.path}: duplicate manifest row`)
    seen.add(asset.path)
    const path = join(root, asset.path)
    if (!existsSync(path)) {
      errors.push(`${asset.path}: manifest row points to a missing file`)
      continue
    }
    const dimensions = imageDimensions(root, asset.path)
    const hash = fileHash(root, asset.path)
    if (asset.width !== dimensions.width || asset.height !== dimensions.height) {
      if (options.write) {
        asset.width = dimensions.width
        asset.height = dimensions.height
      } else {
        errors.push(`${asset.path}: dimensions ${String(asset.width)}x${String(asset.height)} are stale; expected ${String(dimensions.width)}x${String(dimensions.height)}`)
      }
    }
    if (asset.sha256 !== hash) {
      if (options.write) {
        asset.sha256 = hash
      } else {
        errors.push(`${asset.path}: SHA-256 is stale; run pnpm run verify-brand-assets --write`)
      }
    }
  }

  for (const path of listFiles(root, BRAND_PATTERNS)) {
    if (!seen.has(path)) errors.push(`${path}: brand asset is not referenced by the manifest`)
  }

  for (const row of manifest.evidence) {
    const path = join(root, row.path)
    if (!existsSync(path)) {
      errors.push(`${row.id}: evidence file ${row.path} is missing`)
      continue
    }
    const dimensions = pngDimensions(readFileSync(path))
    const expectedWidth = row.viewport.width * row.scale
    const expectedHeight = row.viewport.height * row.scale
    if (dimensions.width !== expectedWidth || dimensions.height !== expectedHeight) {
      errors.push(`${row.id}: evidence is ${String(dimensions.width)}x${String(dimensions.height)}; expected ${String(expectedWidth)}x${String(expectedHeight)}`)
    }
    const hash = fileHash(root, row.path)
    if (row.sha256 !== hash) {
      if (options.write) {
        row.sha256 = hash
      } else {
        errors.push(`${row.id}: evidence SHA-256 is stale; run pnpm run verify-brand-assets --write`)
      }
    }
    for (const depicted of row.depicts) {
      if (!seen.has(depicted)) errors.push(`${row.id}: depicts ${depicted}, which is not a manifest asset`)
    }
  }

  const evidencePaths = new Set(manifest.evidence.map(row => row.path))
  for (const path of listFiles(root, EVIDENCE_PATTERNS)) {
    if (!evidencePaths.has(path)) errors.push(`${path}: evidence image is not referenced by the manifest`)
  }
  for (const id of REQUIRED_EVIDENCE) {
    if (!manifest.evidence.some(row => row.id === id)) errors.push(`${id}: required evidence row is missing`)
  }

  for (const fingerprint of manifest.fingerprints) {
    for (const asset of manifest.assets) {
      const path = join(root, asset.path)
      if (!existsSync(path)) continue
      if (fingerprint.kind === 'sha256') {
        if (fileHash(root, asset.path) === fingerprint.value) {
          errors.push(`${asset.path}: matches retired ${fingerprint.artifact}; expected ${fingerprint.replacement}`)
        }
        continue
      }
      if (!asset.path.endsWith('.svg')) continue
      if (readFileSync(path, 'utf8').includes(fingerprint.value)) {
        errors.push(`${asset.path}: contains retired ${fingerprint.artifact}; expected ${fingerprint.replacement}`)
      }
    }
  }

  if (options.write && errors.length === 0) {
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
  }

  return { errors, assets: manifest.assets.length, evidence: manifest.evidence.length }
}

/**
 * Run the verifier from the command line.
 * @returns the process exit code.
 */
export function main(): number {
  const write = process.argv.includes('--write')
  const result = verifyBrandAssets({ write })
  if (result.errors.length > 0) {
    for (const error of result.errors) console.error(`verify-brand-assets: ${error}`)
    return 1
  }
  console.log(`verify-brand-assets: ${String(result.assets)} assets and ${String(result.evidence)} evidence records verified`)
  return 0
}

if (process.argv[1] !== undefined && import.meta.filename === resolve(process.argv[1])) {
  process.exitCode = main()
}
