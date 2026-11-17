/** Finished-product evidence for the ForWork lockup in the built website header. */

import { createServer, type Server } from 'node:http'
import { readFile } from 'node:fs/promises'
import { copyFileSync, mkdirSync, readFileSync } from 'node:fs'
import { extname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { chromium, type Browser, type Page } from 'playwright'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

const DIST = fileURLToPath(new URL('../../../website/.dist', import.meta.url))
const ARTIFACTS = fileURLToPath(new URL('../../.artifacts/brand/', import.meta.url))
const EXPECTED = fileURLToPath(new URL('../../../website/tests/expected/brand/', import.meta.url))
const UPDATE = process.env.UPDATE_BRAND_SCREENSHOTS === '1'

const MIME: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.woff2': 'font/woff2',
}

/**
 * Capture one header screenshot into the CI artifacts and the reviewed set.
 * @param page - loaded website page.
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

let server: Server
let origin: string
let browser: Browser

beforeAll(async () => {
  server = createServer((request, response) => {
    void (async () => {
      try {
        const pathname = decodeURIComponent(new URL(request.url ?? '/', 'http://localhost').pathname)
        const base = join(DIST, pathname)
        const candidates = pathname.endsWith('/')
          ? [join(base, 'index.html')]
          : [base, `${base}.html`, join(base, 'index.html')]
        let body: Buffer | undefined
        let file = ''
        for (const candidate of candidates) {
          try {
            body = await readFile(candidate)
            file = candidate
            break
          } catch {
            // Try the next clean-URL candidate.
          }
        }
        if (body === undefined) throw new Error(`not found: ${pathname}`)
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
  if (address === null || typeof address === 'string') throw new Error('brand website: no server port')
  origin = `http://127.0.0.1:${String(address.port)}`
  browser = await chromium.launch()
}, 60_000)

afterAll(async () => {
  await browser?.close()
  await new Promise<void>((resolve, reject) => {
    server?.close((error) => { if (error) reject(error); else resolve() })
  })
})

describe('website e2e: ForWork header lockup', () => {
  it('captures the built header at desktop and mobile widths', async () => {
    const desktop = await browser.newPage({ viewport: { width: 1440, height: 240 }, locale: 'en-US' })
    try {
      await desktop.goto(`${origin}/en/`, { waitUntil: 'networkidle' })
      await desktop.locator('.dsh-wordmark').waitFor({ state: 'attached', timeout: 15_000 })
      await desktop.waitForTimeout(500)
      await capture(desktop, 'nav-desktop')
    } finally {
      await desktop.close()
    }

    const mobile = await browser.newPage({ viewport: { width: 390, height: 720 }, locale: 'en-US' })
    try {
      await mobile.goto(`${origin}/en/`, { waitUntil: 'networkidle' })
      await mobile.locator('.dsh-wordmark').waitFor({ state: 'attached', timeout: 15_000 })
      await mobile.waitForTimeout(500)
      await capture(mobile, 'nav-mobile')
    } finally {
      await mobile.close()
    }
  }, 120_000)
})
