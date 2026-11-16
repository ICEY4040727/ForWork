import { mkdtemp, readFile, readdir, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import type { SafeStorageLike } from '@forwork/credentials'
import { DesktopSecureCredentials, isSecureStoreRequest } from '../src/secure-credentials.ts'

const SECRET = 'sk-desktop-model-credential-do-not-store-me'

/** Reversible stand-in for Electron's OS-keychain encryption. */
class FakeSafeStorage implements SafeStorageLike {
  available = true

  isEncryptionAvailable(): boolean {
    return this.available
  }

  encryptString(plainText: string): Buffer {
    return Buffer.from(Buffer.from(plainText, 'utf8').map(byte => byte ^ 0x31))
  }

  decryptString(encrypted: Buffer): string {
    return Buffer.from(encrypted.map(byte => byte ^ 0x31)).toString('utf8')
  }
}

describe('DesktopSecureCredentials', () => {
  let root: string
  let safeStorage: FakeSafeStorage
  let credentials: DesktopSecureCredentials

  beforeEach(async () => {
    root = await mkdtemp(join(tmpdir(), 'forwork-desktop-secure-credentials-'))
    safeStorage = new FakeSafeStorage()
    credentials = new DesktopSecureCredentials(safeStorage, root)
  })

  afterEach(async () => {
    await rm(root, { recursive: true, force: true })
  })

  it('writes and reads model credentials through the safeStorage store', async () => {
    expect(credentials.available()).toBe(true)

    await expect(credentials.handle({
      type: 'secure-store-request', requestId: 1, action: 'set', key: 'FORWORK_API_KEY', value: SECRET,
    })).resolves.toEqual({ type: 'secure-store-response', requestId: 1, ok: true })

    await expect(credentials.handle({
      type: 'secure-store-request', requestId: 2, action: 'get', key: 'FORWORK_API_KEY',
    })).resolves.toEqual({ type: 'secure-store-response', requestId: 2, ok: true, value: SECRET })

    const stored = await readFile(join(root, 'desktop', 'secure-credentials.json'), 'utf8')
    expect(stored).not.toContain(SECRET)
    expect(stored).toContain(safeStorage.encryptString(SECRET).toString('base64'))

    await expect(credentials.handle({
      type: 'secure-store-request', requestId: 3, action: 'delete', key: 'FORWORK_API_KEY',
    })).resolves.toEqual({ type: 'secure-store-response', requestId: 3, ok: true })

    await expect(credentials.handle({
      type: 'secure-store-request', requestId: 4, action: 'get', key: 'FORWORK_API_KEY',
    })).resolves.toEqual({ type: 'secure-store-response', requestId: 4, ok: true })
  })

  it('refuses to read or write while encryption is unavailable and writes no plaintext', async () => {
    safeStorage.available = false
    expect(credentials.available()).toBe(false)

    await expect(credentials.handle({
      type: 'secure-store-request', requestId: 1, action: 'set', key: 'FORWORK_API_KEY', value: SECRET,
    })).resolves.toEqual({
      type: 'secure-store-response', requestId: 1, ok: false, code: 'secure-storage-unavailable',
    })
    await expect(credentials.handle({
      type: 'secure-store-request', requestId: 2, action: 'get', key: 'FORWORK_API_KEY',
    })).resolves.toEqual({
      type: 'secure-store-response', requestId: 2, ok: false, code: 'secure-storage-unavailable',
    })

    const directory = join(root, 'desktop')
    await expect(readdir(directory).catch(() => [])).resolves.toEqual([])
  })
})

describe('isSecureStoreRequest', () => {
  it('accepts well-formed get, set, and delete requests', () => {
    expect(isSecureStoreRequest({ type: 'secure-store-request', requestId: 1, action: 'get', key: 'K' })).toBe(true)
    expect(isSecureStoreRequest({
      type: 'secure-store-request', requestId: 1, action: 'set', key: 'K', value: 'secret',
    })).toBe(true)
    expect(isSecureStoreRequest({ type: 'secure-store-request', requestId: 1, action: 'delete', key: 'K' })).toBe(true)
  })

  it.each([
    null,
    { type: 'secure-store-request', requestId: 1, action: 'read', key: 'K' },
    { type: 'secure-store-request', requestId: 1.5, action: 'get', key: 'K' },
    { type: 'secure-store-request', requestId: 1, action: 'get', key: '' },
    { type: 'secure-store-request', requestId: 1, action: 'get', key: 'K', value: 'unexpected' },
    { type: 'secure-store-request', requestId: 1, action: 'set', key: 'K' },
    { type: 'secure-store-request', requestId: 1, action: 'set', key: 'K', value: '' },
  ])('rejects %j', (value) => {
    expect(isSecureStoreRequest(value)).toBe(false)
  })
})
