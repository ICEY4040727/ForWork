/** Finished-product evidence for the ForWork door mark on the shipped Web home. */

import { copyFileSync, mkdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { chromium, type Browser, type Page } from 'playwright'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { launchWebScaffold, type WebScaffold } from './scaffold.ts'

const ARTIFACTS = fileURLToPath(new URL('../../../.artifacts/brand/', import.meta.url))
const EXPECTED = fileURLToPath(new URL('./expected/brand/', import.meta.url))
const ICON = fileURLToPath(new URL('../../desktop/resources/icon.png', import.meta.url))
const UPDATE = process.env.UPDATE_BRAND_SCREENSHOTS === '1'

/**
 * Capture one finished-product screenshot into the CI artifacts and, when the
 * review set is being refreshed, into the committed expected directory.
 * @param page - page showing the finished surface.
 * @param name - manifest evidence id and file stem.
 */
async function capture(page: Page, name: string): Promise<void> {
  mkdirSync(ARTIFACTS, { recursive: true })
  const file = join(ARTIFACTS, `${name}.png`)
  await page.evaluate(() => document.fonts.ready.then(() => undefined))
  await page.evaluate(() => new Promise<void>((resolve) => {
    requestAnimationFrame(() => { requestAnimationFrame(() => { resolve() }) })
  }))
  await page.screenshot({ path: file, animations: 'disabled' })
  if (UPDATE) {
    mkdirSync(EXPECTED, { recursive: true })
    copyFileSync(file, join(EXPECTED, `${name}.png`))
  }
  expect(readFileSync(file).length).toBeGreaterThan(1_000)
}

/**
 * Open the blank home in one theme and return the page after the door mark renders.
 * @param browser - browser owning the page.
 * @param width - viewport width.
 * @param height - viewport height.
 * @param colorScheme - requested palette.
 * @returns the loaded page.
 */
async function openHome(
  browser: Browser,
  width: number,
  height: number,
  colorScheme: 'light' | 'dark',
): Promise<Page> {
  const page = await browser.newPage({
    viewport: { width, height },
    locale: 'en-US',
    timezoneId: 'Asia/Shanghai',
    colorScheme,
  })
  await page.goto(scaffold.authenticatedUrl, { waitUntil: 'load' })
  await page.locator('div[data-phase]').first().waitFor({ timeout: 30_000 })
  await page.getByText('Into the Unknown', { exact: true }).waitFor({ timeout: 15_000 })
  await expect.poll(() => page.evaluate(() => document.body.hasAttribute('data-ds-dark-theme')))
    .toBe(colorScheme === 'dark')
  return page
}

/** The hero mark hitbox immediately before the headline title group. */
function markHitbox(page: Page) {
  const headline = page.getByText('Into the Unknown', { exact: true })
  return headline.locator('xpath=../preceding-sibling::span[1]')
}

let scaffold: WebScaffold
let browser: Browser

beforeAll(async () => {
  scaffold = await launchWebScaffold({})
  browser = await chromium.launch()
}, 180_000)

afterAll(async () => {
  await browser?.close()
  await scaffold?.close()
})

describe('web e2e: ForWork brand artwork', () => {
  it('captures the blank home in light and dark mode', async () => {
    const light = await openHome(browser, 1440, 900, 'light')
    try {
      const mark = markHitbox(light).locator('svg')
      expect(await mark.count()).toBe(1)
      expect(await mark.locator('path').count()).toBe(2)
      await capture(light, 'home-desktop-light')
    } finally {
      await light.close()
    }

    const dark = await openHome(browser, 1440, 900, 'dark')
    try {
      const mark = markHitbox(dark).locator('svg')
      expect(await mark.count()).toBe(1)
      await capture(dark, 'home-desktop-dark')
    } finally {
      await dark.close()
    }
  }, 120_000)

  it('captures the door leaf in its hover-open state', async () => {
    const page = await openHome(browser, 1440, 900, 'light')
    try {
      const hitbox = markHitbox(page)
      const leaf = hitbox.locator('svg path').last()
      await hitbox.hover()
      await expect.poll(() => leaf.evaluate(node => getComputedStyle(node).transform))
        .not.toBe('none')
      await capture(page, 'home-hover')
    } finally {
      await page.close()
    }
  }, 120_000)

  it('captures the mobile drawer and the collapsed desktop rail', async () => {
    const mobile = await openHome(browser, 390, 844, 'light')
    try {
      await mobile.getByRole('button', { name: 'Open sidebar', exact: true }).click()
      await mobile.getByRole('dialog', { name: 'Navigation', exact: true }).waitFor()
      await capture(mobile, 'home-mobile')
    } finally {
      await mobile.close()
    }

    const rail = await openHome(browser, 1440, 900, 'light')
    try {
      await rail.getByRole('button', { name: 'Collapse sidebar', exact: true }).click()
      await expect.poll(() => rail.getByRole('button', { name: 'Open sidebar', exact: true }).count())
        .toBe(1)
      await capture(rail, 'rail-collapsed')
    } finally {
      await rail.close()
    }
  }, 120_000)

  it('captures the install icon at launcher sizes', async () => {
    const icon = readFileSync(ICON).toString('base64')
    const sizes = [16, 24, 32, 48, 64, 128, 192, 256]
    const page = await browser.newPage({ viewport: { width: 1024, height: 320 }, deviceScaleFactor: 1 })
    try {
      await page.setContent(`<!doctype html><html><body style="margin:0;background:#eef1f6;font:12px 'Segoe UI',sans-serif;color:#475569">
        <div style="display:flex;align-items:flex-end;gap:26px;padding:32px 36px">
          ${sizes.map(size => `<figure style="margin:0;display:flex;flex-direction:column;align-items:center;gap:8px">
            <img src="data:image/png;base64,${icon}" width="${String(size)}" height="${String(size)}" alt="">
            <figcaption>${String(size)}</figcaption>
          </figure>`).join('')}
        </div>
      </body></html>`, { waitUntil: 'load' })
      await page.waitForFunction(() => [...document.images].every(image => image.complete))
      await capture(page, 'icon-sheet')
    } finally {
      await page.close()
    }
  }, 60_000)
})
