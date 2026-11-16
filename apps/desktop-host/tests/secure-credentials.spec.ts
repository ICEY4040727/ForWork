import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Context } from '@forwork/cordis'
import { credentialKey, credentialRef, SecureStorageUnavailableError, type SecureStore } from '@forwork/credentials'
import DesktopSecureCredentials from '../src/secure-credentials.ts'

const REF = credentialRef('FORWORK_API_KEY')
const LEGACY_REF = credentialRef('LEGACY_KEY')
const RECORD_KEY = credentialKey('desktop-test', 'provider')

/** In-memory OS secure store that records availability like Electron's `safeStorage`. */
class MemoryStore implements SecureStore {
  available = true
  readonly entries = new Map<string, string>()

  get(key: string): Promise<string | undefined> {
    if (!this.available) return Promise.reject(new SecureStorageUnavailableError())
    return Promise.resolve(this.entries.get(key))
  }

  set(key: string, value: string): Promise<void> {
    if (!this.available) return Promise.reject(new SecureStorageUnavailableError())
    this.entries.set(key, value)
    return Promise.resolve()
  }

  delete(key: string): Promise<void> {
    this.entries.delete(key)
    return Promise.resolve()
  }
}

/** Mount the provider with an injected store; only class plugins run their service init. */
function mount(ctx: Context, dshHome: string, store: SecureStore): Promise<unknown> {
  class TestProvider extends DesktopSecureCredentials {
    constructor(context: Context, config: ConstructorParameters<typeof DesktopSecureCredentials>[1]) {
      super(context, config, store)
    }
  }
  return Promise.resolve(ctx.plugin(TestProvider, { dshHome }))
}

describe('DesktopSecureCredentials', () => {
  let root: string
  let store: MemoryStore

  beforeEach(async () => {
    root = await mkdtemp(join(tmpdir(), 'forwork-desktop-host-credentials-'))
    store = new MemoryStore()
  })

  afterEach(async () => {
    await rm(root, { recursive: true, force: true })
  })

  it('stores model references only in the secure store', async () => {
    const ctx = new Context()
    await mount(ctx, root, store)

    await ctx.credentials.set(REF, 'sk-secure')
    expect(store.entries.get(REF)).toBe('sk-secure')
    await expect(ctx.credentials.resolve(REF)).resolves.toEqual({ value: 'sk-secure', source: 'secure-store' })
    expect(await ctx.credentials.describe(REF)).toEqual({ configured: true, source: 'secure-store', writable: true })

    await ctx.credentials.unset(REF)
    expect(store.entries.has(REF)).toBe(false)
    await expect(ctx.credentials.resolve(REF)).resolves.toBeUndefined()
  })

  it('moves legacy plaintext references into the secure store and out of the local document', async () => {
    const document = join(root, '.credentials.yaml')
    await writeFile(document, 'version: 1\nrefs:\n  LEGACY_KEY: sk-legacy\n')
    const ctx = new Context()
    await mount(ctx, root, store)

    expect(store.entries.get(LEGACY_REF)).toBe('sk-legacy')
    await expect(ctx.credentials.resolve(LEGACY_REF)).resolves.toEqual({ value: 'sk-legacy', source: 'secure-store' })
    expect(await readFile(document, 'utf8')).not.toContain('sk-legacy')
  })

  it('delegates account records to the local document', async () => {
    const ctx = new Context()
    await mount(ctx, root, store)
    const events: string[] = []
    ctx.on('credentials/record-updated', key => void events.push(key))

    await ctx.credentials.modifyRecord(RECORD_KEY, async () => ({ kind: 'api-key', key: 'record-secret' }))
    await expect(ctx.credentials.readRecord(RECORD_KEY)).resolves.toEqual({ kind: 'api-key', key: 'record-secret' })
    expect(events).toEqual([RECORD_KEY])
    expect(await ctx.credentials.listRecords()).toEqual([{ key: RECORD_KEY, kind: 'api-key' }])
    await ctx.credentials.deleteRecord(RECORD_KEY)
    await expect(ctx.credentials.readRecord(RECORD_KEY)).resolves.toBeUndefined()
  })

  it('refuses model writes while the secure store is unavailable instead of using plaintext', async () => {
    store.available = false
    const ctx = new Context()
    await mount(ctx, root, store)

    await expect(ctx.credentials.set(REF, 'sk-refused')).rejects.toBeInstanceOf(SecureStorageUnavailableError)
    expect(store.entries.has(REF)).toBe(false)
    await expect(ctx.credentials.describe(REF)).resolves.toEqual({ configured: false, writable: false })
  })
})
