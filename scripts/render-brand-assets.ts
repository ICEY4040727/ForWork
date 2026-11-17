/**
 * Derive the desktop SVG icon variants from the canonical ForWork tile master.
 *
 * The 1024 tile master is the square Web crop source; the Windows and macOS
 * variants keep the same 1104 canvas, add their platform corner radius, and
 * center the door on the tile. `--check` compares the committed files with the
 * generated text without writing, so the brand gate can detect drift.
 *
 * @module render-brand-assets
 */

import { readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'

/** Repository root, resolved from this script. */
const root = join(import.meta.dirname, '..')

/** The canonical tile master. */
const MASTER = 'assets/brand/forwork-tile.svg'

/** A generated desktop source and its content. */
export interface DerivedBrandAsset {
  /** Repository-relative output path. */
  path: string
  /** Complete SVG text. */
  content: string
}

/** Shared drop shadows; the square source and both platform variants use the same depth. */
const DEFS = `<defs>
<filter id="forwork-tile-shadow" x="0" y="0" width="1104" height="1104" filterUnits="userSpaceOnUse" color-interpolation-filters="sRGB">
<feFlood flood-opacity="0" result="BackgroundImageFix"/>
<feColorMatrix in="SourceAlpha" type="matrix" values="0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 127 0" result="hardAlpha"/>
<feOffset dy="8"/>
<feGaussianBlur stdDeviation="20"/>
<feComposite in2="hardAlpha" operator="out"/>
<feColorMatrix type="matrix" values="0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0.06 0"/>
<feBlend mode="normal" in2="BackgroundImageFix" result="effect1_dropShadow"/>
<feBlend mode="normal" in="SourceGraphic" in2="effect1_dropShadow" result="shape"/>
</filter>
</defs>`

/** One SVG document with the shared filters appended. */
function document(body: string): string {
  return `<svg width="1104" height="1104" viewBox="0 0 1104 1104" fill="none" xmlns="http://www.w3.org/2000/svg">
${body}
${DEFS}
</svg>
`
}

/** One 1024 platform document that fits the 1104 canvas into a 1024 viewport. */
function platformDocument(body: string): string {
  return `<svg width="1024" height="1024" viewBox="0 0 1024 1024" fill="none" xmlns="http://www.w3.org/2000/svg">
<g transform="translate(32 32) scale(0.9375) translate(-40 -32)">
${body}
</g>
${DEFS}
</svg>
`
}

/**
 * Read the door group from the canonical tile master.
 * @returns the transformed door and handle group.
 */
function readDoorGroup(): string {
  const master = readFileSync(join(root, MASTER), 'utf8')
  const match = /<g transform="translate\(512 512\)[^"]*">[\s\S]*?<\/g>/u.exec(master)
  if (match === null) throw new Error(`${MASTER}: missing translated door group`)
  return match[0]
}

/**
 * Build every desktop SVG source from the tile master.
 * @returns the generated sources in write order.
 */
export function renderBrandAssets(): readonly DerivedBrandAsset[] {
  const door = `<g transform="translate(40 32)">${readDoorGroup()}</g>`
  return [
    {
      path: 'apps/desktop/resources/icon.svg',
      content: document(`<g filter="url(#forwork-tile-shadow)">
<rect width="1024" height="1024" transform="translate(40 32)" fill="#F4F2EC"/>
<g id="forwork-mark-artwork">
${door}
</g>
</g>`),
    },
    {
      path: 'apps/desktop/resources/icon-windows.svg',
      content: platformDocument(`<g filter="url(#forwork-tile-shadow)">
<rect rx="192" width="1024" height="1024" transform="translate(40 32)" fill="#F4F2EC"/>
<g id="tray-glyph">
${door}
</g>
</g>`),
    },
    {
      path: 'apps/desktop/resources/icon-macos.svg',
      content: platformDocument(`<g filter="url(#forwork-tile-shadow)">
<rect rx="230" width="1024" height="1024" transform="translate(40 32)" fill="#F4F2EC"/>
<g id="forwork-mark-artwork">
${door}
</g>
</g>`),
    },
  ]
}

/**
 * Write the derived sources, or compare them when `--check` is present.
 * @returns the process exit code.
 */
export function main(): number {
  const check = process.argv.includes('--check')
  let stale = 0
  for (const asset of renderBrandAssets()) {
    const path = join(root, asset.path)
    if (check) {
      const current = readFileSync(path, 'utf8')
      if (current !== asset.content) {
        console.error(`render-brand-assets: ${asset.path} is stale; run pnpm run render-brand-assets`)
        stale += 1
      }
      continue
    }
    writeFileSync(path, asset.content)
    console.log(`render-brand-assets: wrote ${asset.path}`)
  }
  return stale === 0 ? 0 : 1
}

if (process.argv[1] !== undefined && import.meta.filename === resolve(process.argv[1])) process.exitCode = main()
