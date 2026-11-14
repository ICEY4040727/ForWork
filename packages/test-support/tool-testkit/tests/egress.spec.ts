/** Egress recorder and allowlist assertions: destination capture, restoration, and refusal rules. */
import http, { type ClientRequest } from 'node:http'
import https from 'node:https'
import { afterEach, describe, expect, it } from 'vitest'
import { assertNoEgress, recordEgress } from '../src/index.ts'

const originalFetch = globalThis.fetch
const originalHttpRequest = http.request
const originalHttpsRequest = https.request

afterEach(() => {
  globalThis.fetch = originalFetch
  http.request = originalHttpRequest
  https.request = originalHttpsRequest
})

/** A Node request stand-in that records nothing and never opens a socket. */
function refuseRequest(): ClientRequest {
  throw new Error('test transport refused to open a socket')
}

describe('recordEgress', () => {
  it('records fetch destinations while forwarding to the installed stub', async () => {
    const seen: string[] = []
    const stub: typeof fetch = async (input) => {
      seen.push(typeof input === 'string' ? input : input instanceof URL ? input.href : input.url)
      return new Response('ok')
    }
    globalThis.fetch = stub

    const capture = await recordEgress(async () => {
      const response = await fetch('https://documents.invalid/upload')
      expect(await response.text()).toBe('ok')
      await fetch(new URL('https://models.invalid/v1/chat'))
      await fetch(new Request('https://models.invalid/v1/embeddings'))
      return 'finished'
    })

    expect(capture.result).toBe('finished')
    expect(capture.urls).toEqual([
      'https://documents.invalid/upload',
      'https://models.invalid/v1/chat',
      'https://models.invalid/v1/embeddings',
    ])
    expect(seen).toEqual(capture.urls)
    expect(globalThis.fetch).toBe(stub)
  })

  it('restores fetch after the body throws', async () => {
    const stub: typeof fetch = async () => new Response('ok')
    globalThis.fetch = stub

    await expect(recordEgress(async () => {
      await fetch('https://documents.invalid/upload')
      throw new Error('body failed')
    })).rejects.toThrow('body failed')

    expect(globalThis.fetch).toBe(stub)
  })

  it('refuses a nested recorder without disturbing the outer recording', async () => {
    await expect(recordEgress(async () => {
      await expect(recordEgress(async () => 'inner')).rejects.toThrow('already active')
      return 'outer'
    })).resolves.toEqual({ result: 'outer', urls: [] })

    expect(globalThis.fetch).toBe(originalFetch)
  })

  it('records http and https request destinations without opening a socket', async () => {
    http.request = refuseRequest
    https.request = refuseRequest

    const capture = await recordEgress(async () => {
      expect(() => http.request({ hostname: 'ingest.invalid', port: 8080, path: '/upload' })).toThrow('refused')
      expect(() => http.request({ hostname: 'secure.invalid', protocol: 'https:', path: '/v1' })).toThrow('refused')
      expect(() => http.request({ host: 'legacy.invalid' })).toThrow('refused')
      expect(() => http.request({})).toThrow('refused')
      expect(() => http.request('http://files.invalid/raw')).toThrow('refused')
      expect(() => http.request('http://files.invalid/raw', () => {})).toThrow('refused')
      expect(() => http.request({ hostname: 'callbacks.invalid' }, () => {})).toThrow('refused')
      expect(() => https.request(new URL('https://cloud.invalid/v1/2'), { method: 'POST' })).toThrow('refused')
      return 'finished'
    })

    expect(capture.urls).toEqual([
      'http://ingest.invalid:8080/upload',
      'https://secure.invalid/v1',
      'http://legacy.invalid/',
      'http://localhost/',
      'http://files.invalid/raw',
      'http://files.invalid/raw',
      'http://callbacks.invalid/',
      'https://cloud.invalid/v1/2',
    ])
    expect(http.request).toBe(refuseRequest)
    expect(https.request).toBe(refuseRequest)
  })
})

describe('assertNoEgress', () => {
  it('passes an empty list', () => {
    expect(() => { assertNoEgress([]) }).not.toThrow()
  })

  it('fails a destination outside the explicit allowlist', () => {
    expect(() => { assertNoEgress(['https://documents.invalid/upload']) })
      .toThrow('Unexpected network egress to https://documents.invalid/upload; allowed hosts: (none)')
    expect(() => { assertNoEgress(['https://api.deepseek.com/v1'], { allow: ['api.deepseek.com'] }) }).not.toThrow()
    expect(() => { assertNoEgress(['https://API.DeepSeek.com/v1'], { allow: ['api.deepseek.com'] }) }).not.toThrow()
    expect(() => { assertNoEgress(['https://documents.invalid/upload'], { allow: ['api.deepseek.com'] }) })
      .toThrow('allowed hosts: api.deepseek.com')
  })

  it('fails a destination that is not an absolute URL', () => {
    expect(() => { assertNoEgress(['not a URL'], { allow: ['api.deepseek.com'] }) })
      .toThrow('Unexpected network egress to not a URL')
  })
})
