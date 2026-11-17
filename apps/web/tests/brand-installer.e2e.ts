/** Finished-product evidence for the installer and uninstaller brand artwork. */

import { copyFileSync, mkdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { chromium, type Browser, type Page } from 'playwright'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

const ARTIFACTS = fileURLToPath(new URL('../../../.artifacts/brand/', import.meta.url))
const EXPECTED = fileURLToPath(new URL('../../desktop/tests/expected/brand/', import.meta.url))
const UPDATE = process.env.UPDATE_BRAND_SCREENSHOTS === '1'

const BRAND_LIGHT = fileURLToPath(new URL('../../desktop/installer/assets/brand.png', import.meta.url))
const BRAND_DARK = fileURLToPath(new URL('../../desktop/installer/assets/brand-dark.png', import.meta.url))
const UNINSTALLER = fileURLToPath(new URL('../../desktop/installer/assets/uninstaller-sidebar.png', import.meta.url))

/**
 * Capture one installer asset at its shipped size.
 * @param page - browser page.
 * @param name - evidence id and file stem.
 */
async function capture(page: Page, name: string): Promise<void> {
  mkdirSync(ARTIFACTS, { recursive: true })
  const file = join(ARTIFACTS, `${name}.png`)
  await page.screenshot({ path: file, animations: 'disabled' })
  if (UPDATE) {
    mkdirSync(EXPECTED, { recursive: true })
    copyFileSync(file, join(EXPECTED, `${name}.png`))
  }
  expect(readFileSync(file).length).toBeGreaterThan(1_000)
}

/**
 * Render one shipped asset to its native size.
 * @param page - browser page.
 * @param source - local PNG path.
 * @param width - native width.
 * @param height - native height.
 * @param background - installer background behind the transparent asset.
 */
async function show(page: Page, source: string, width: number, height: number, background: string): Promise<void> {
  const data = readFileSync(source).toString('base64')
  await page.setViewportSize({ width, height })
  await page.setContent(`<!doctype html><html><body style="margin:0;background:${background}">
    <img src="data:image/png;base64,${data}" width="${String(width)}" height="${String(height)}" alt="">
  </body></html>`, { waitUntil: 'load' })
  await page.waitForFunction(() => [...document.images].every(image => image.complete))
}

let browser: Browser

beforeAll(async () => {
  browser = await chromium.launch()
}, 60_000)

afterAll(async () => {
  await browser?.close()
})

describe('desktop e2e: installer brand artwork', () => {
  it('captures the shipped installer and uninstaller assets at native size', async () => {
    const page = await browser.newPage({ viewport: { width: 600, height: 196 }, deviceScaleFactor: 1 })
    try {
      await show(page, BRAND_LIGHT, 600, 196, '#ffffff')
      await capture(page, 'installer-light')
      await show(page, BRAND_DARK, 600, 196, '#141414')
      await capture(page, 'installer-dark')
      await show(page, UNINSTALLER, 164, 314, '#f4f4f4')
      await capture(page, 'uninstaller')
    } finally {
      await page.close()
    }
  }, 60_000)
})
