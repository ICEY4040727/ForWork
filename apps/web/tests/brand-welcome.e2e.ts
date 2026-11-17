/** Finished-product evidence for the ForWork lockup in the native welcome renderer. */

import { createServer, type Server } from 'node:http'
import { readFile } from 'node:fs/promises'
import { copyFileSync, mkdirSync, readFileSync } from 'node:fs'
import { extname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { chromium, type Browser, type Page } from 'playwright'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { resolveDesktopLocale } from '../../desktop/src/locale.ts'

const ROOT = fileURLToPath(new URL('../../desktop/', import.meta.url))
const ARTIFACTS = fileURLToPath(new URL('../../../.artifacts/brand/', import.meta.url))
const EXPECTED = fileURLToPath(new URL('../../desktop/tests/expected/brand/', import.meta.url))
const UPDATE = process.env.UPDATE_BRAND_SCREENSHOTS === '1'

const MIME: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.woff2': 'font/woff2',
}

/**
 * Capture one welcome screenshot into the CI artifacts and the reviewed set.
 * @param page - loaded welcome page.
 * @param name - evidence id and file stem.
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
 * Open the built welcome renderer with a typed preload mock.
 * @param browser - browser owning the page.
 * @param origin - local renderer origin.
 * @param colorScheme - requested palette.
 * @returns the page after the lockup image loads.
 */
async function openWelcome(
  browser: Browser,
  origin: string,
  colorScheme: 'light' | 'dark',
): Promise<Page> {
  const api = resolveDesktopLocale('en')
  const page = await browser.newPage({
    viewport: { width: 600, height: 700 },
    deviceScaleFactor: 2,
    locale: 'en-US',
    colorScheme,
  })
  await page.addInitScript((copy) => {
    Object.defineProperty(window, 'dshWelcome', {
      value: {
        ...copy,
        analytics: async () => undefined,
        analyticsEnabled: async () => false,
        takeNotice: async () => undefined,
        startSignIn: async () => ({ status: 'signed-out', attempt: null }),
        cancelSignIn: async () => ({ status: 'signed-out', attempt: null }),
        copySignInLink: async () => undefined,
        onAccountState: () => () => undefined,
        saveApiKey: async () => ({ ok: true }),
        skip: async () => undefined,
      },
    })
  }, api)
  await page.goto(`${origin}/renderer/welcome.html`, { waitUntil: 'load' })
  const brand = page.locator('img.brand')
  await brand.waitFor({ timeout: 15_000 })
  await expect.poll(() => brand.evaluate((image: HTMLImageElement) => image.complete && image.naturalWidth > 0))
    .toBe(true)
  await expect.poll(() => page.evaluate(() => document.body.hasAttribute('data-ds-dark-theme')))
    .toBe(colorScheme === 'dark')
  return page
}

let server: Server
let origin: string
let browser: Browser

beforeAll(async () => {
  server = createServer((request, response) => {
    void (async () => {
      try {
        const pathname = decodeURIComponent(new URL(request.url ?? '/', 'http://localhost').pathname)
        const file = join(ROOT, pathname)
        const body = await readFile(file)
        response.writeHead(200, { 'content-type': MIME[extname(file)] ?? 'application/octet-stream' })
        response.end(body)
      } catch {
        response.writeHead(404)
        response.end('not found')
      }
    })()
  })
  await new Promise<void>((resolve) => { server.listen(0, '127.0.0.1', resolve) })
  const address = server.address()
  if (address === null || typeof address === 'string') throw new Error('brand welcome: no server port')
  origin = `http://127.0.0.1:${String(address.port)}`
  browser = await chromium.launch()
}, 60_000)

afterAll(async () => {
  await browser?.close()
  await new Promise<void>((resolve, reject) => {
    server?.close((error) => { if (error) reject(error); else resolve() })
  })
})

describe('desktop e2e: ForWork welcome brand', () => {
  it('captures the welcome lockup in light and dark mode', async () => {
    const light = await openWelcome(browser, origin, 'light')
    try {
      await capture(light, 'welcome-light')
    } finally {
      await light.close()
    }
    const dark = await openWelcome(browser, origin, 'dark')
    try {
      await capture(dark, 'welcome-dark')
    } finally {
      await dark.close()
    }
  }, 120_000)
})
