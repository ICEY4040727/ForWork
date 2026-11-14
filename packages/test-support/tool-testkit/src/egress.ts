/**
 * Network-egress observation for privacy tests. Records the destination of
 * every `fetch`, `node:http.request`, and `node:https.request` call made by a
 * test body without replacing, blocking, or replaying the transport itself.
 * Assertions compare recorded hostnames with an explicit allowlist so a test
 * can prove document flows only reach the configured model endpoint.
 *
 * The recorder is process-global and deliberately non-nestable: only one
 * recorded body may run at a time, so a concurrent suite cannot attribute
 * another suite's egress to the current assertion.
 *
 * @module @deepseek-ai/dsh-tool-testkit/egress
 */

import http from 'node:http'
import https from 'node:https'
import type { ClientRequest, IncomingMessage, RequestOptions } from 'node:http'

/** One recorded body's result and every destination it contacted. */
export interface EgressCapture<T> {
  /** The body's own return value. */
  result: T
  /** Destinations in call order, as absolute URLs where the call supplied enough information. */
  urls: readonly string[]
}

/** Hostname allowlist for {@link assertNoEgress}. */
export interface AssertNoEgressOptions {
  /**
   * Hostnames permitted to receive traffic, such as the configured model
   * endpoint. Comparison is case-insensitive and ignores the port. Omitting
   * the list forbids every destination.
   */
  allow?: readonly string[]
}

/** Callback shape shared by the patched `http.request` and `https.request`. */
type RequestCallback = (response: IncomingMessage) => void

/** Whether a recorder currently owns the process-global transport functions. */
let recording = false

/**
 * Run one body while recording every network destination it contacts, then
 * restore `fetch`, `http.request`, and `https.request` exactly as they were.
 * The body still performs its own calls; the recorder only observes them, so
 * a test stub installed before this call remains the transport that runs.
 * Nested or concurrent recorders throw instead of sharing one destination
 * list.
 * @param fn - the body to observe.
 * @returns the body's result plus the recorded destinations in call order.
 */
export async function recordEgress<T>(fn: () => Promise<T>): Promise<EgressCapture<T>> {
  if (recording) {
    throw new Error('recordEgress: a recorder is already active; nested or concurrent recorders are not supported')
  }
  recording = true
  const urls: string[] = []
  const originalFetch = globalThis.fetch
  const originalHttpRequest = http.request
  const originalHttpsRequest = https.request
  const record = (destination: string): void => {
    urls.push(destination)
  }
  try {
    globalThis.fetch = (input, init) => {
      urls.push(fetchDestination(input))
      return originalFetch(input, init)
    }
    http.request = wrapRequest(originalHttpRequest, 'http:', record)
    https.request = wrapRequest(originalHttpsRequest, 'https:', record)
    const result = await fn()
    return { result, urls }
  } finally {
    globalThis.fetch = originalFetch
    http.request = originalHttpRequest
    https.request = originalHttpsRequest
    recording = false
  }
}

/**
 * Assert that no recorded destination reached a host outside the allowlist.
 * An empty list passes; a destination whose URL cannot be parsed, or whose
 * hostname is not allowed, fails the assertion.
 * @param urls - destinations from {@link recordEgress}.
 * @param options - optional hostname allowlist; the model endpoint belongs here.
 */
export function assertNoEgress(urls: readonly string[], options: AssertNoEgressOptions = {}): void {
  const allowed = new Set((options.allow ?? []).map(hostname => hostname.toLowerCase()))
  const unexpected = urls.filter((url) => {
    const hostname = destinationHostname(url)
    return hostname === undefined || !allowed.has(hostname)
  })
  if (unexpected.length === 0) return
  const allowlist = [...allowed].join(', ') || '(none)'
  throw new Error(`Unexpected network egress to ${unexpected.join(', ')}; allowed hosts: ${allowlist}`)
}

/** Read the destination out of any supported `fetch` input. */
function fetchDestination(input: Parameters<typeof fetch>[0]): string {
  if (typeof input === 'string') return input
  if (input instanceof URL) return input.href
  return input.url
}

/** Patch one Node request function, preserving both of its calling forms. */
function wrapRequest(
  original: typeof http.request,
  protocol: string,
  record: (destination: string) => void,
): typeof http.request {
  const wrapped = (
    first: string | URL | RequestOptions,
    second?: RequestOptions | RequestCallback,
    third?: RequestCallback,
  ): ClientRequest => {
    record(requestDestination(first, protocol))
    if (typeof first === 'string' || first instanceof URL) {
      if (typeof second === 'function') return original(first, second)
      if (second === undefined) return original(first, third)
      return original(first, second, third)
    }
    return original(first, typeof second === 'function' ? second : third)
  }
  return wrapped
}

/** Build an absolute destination from a request URL or options object. */
function requestDestination(first: string | URL | RequestOptions, protocol: string): string {
  if (typeof first === 'string') return first
  if (first instanceof URL) return first.href
  const hostname = first.hostname ?? first.host ?? 'localhost'
  const port = first.port === undefined ? '' : `:${String(first.port)}`
  return `${first.protocol ?? protocol}//${hostname}${port}${first.path ?? '/'}`
}

/** Extract the lowercased hostname, or undefined when the destination is not an absolute URL. */
function destinationHostname(destination: string): string | undefined {
  try {
    return new URL(destination).hostname
  } catch {
    return undefined
  }
}
