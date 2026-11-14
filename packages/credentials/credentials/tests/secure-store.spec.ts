import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { SafeStorageSecureStore, SecureStorageUnavailableError } from '../src/index.ts'
import type { SafeStorageLike } from '../src/index.ts'

const SECRET = 'sk-live-1234567890-do-not-store-me'

/** Reversible stand-in for Electron's OS-keychain encryption. */
class FakeSafeStorage implements SafeStorageLike {
  available = true

  isEncryptionAvailable(): boolean {
    return this.available
  }

  encryptString(plainText: string): Buffer {
    return Buffer.from(Buffer.from(plainText, 'utf8').map(byte => byte ^ 0x5a))
  }

  decryptString(encrypted: Buffer): string {
    return Buffer.from(encrypted.map(byte => byte ^ 0x5a)).toString('utf8')
  }
}

async function onlyFile(directory: string): Promise<string> {
  const files = await readdir(directory)
  expect(files).toHaveLength(1)
  const file = files[0]
  if (file === undefined) throw new Error('expected one stored file')
  return join(directory, file)
}

describe('SafeStorageSecureStore', () => {
  let directory: string
  let safeStorage: FakeSafeStorage
  let store: SafeStorageSecureStore

  beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'dsh-secure-store-'))
    safeStorage = new FakeSafeStorage()
    store = new SafeStorageSecureStore({ safeStorage, directory })
  })

  afterEach(async () => {
    await rm(directory, { recursive: true, force: true })
  })

  it('round-trips set, get, and delete', async () => {
    expect(await store.get('DEEPSEEK_API_KEY')).toBeUndefined()

    await store.set('DEEPSEEK_API_KEY', SECRET)
    await store.set('OTHER_API_KEY', 'other-secret')
    expect(await store.get('DEEPSEEK_API_KEY')).toBe(SECRET)
    expect(await store.get('OTHER_API_KEY')).toBe('other-secret')

    await store.delete('DEEPSEEK_API_KEY')
    expect(await store.get('DEEPSEEK_API_KEY')).toBeUndefined()
    expect(await store.get('OTHER_API_KEY')).toBe('other-secret')

    await store.delete('ABSENT_KEY')
  })

  it('throws a typed error and writes nothing while encryption is unavailable', async () => {
    safeStorage.available = false

    await expect(store.set('DEEPSEEK_API_KEY', SECRET)).rejects.toMatchObject({
      name: 'SecureStorageUnavailableError',
      code: 'secure-storage-unavailable',
    })
    await expect(store.get('DEEPSEEK_API_KEY')).rejects.toBeInstanceOf(SecureStorageUnavailableError)
    expect(await readdir(directory)).toEqual([])
  })

  it('persists ciphertext without the plaintext secret', async () => {
    await store.set('DEEPSEEK_API_KEY', SECRET)

    const file = await onlyFile(directory)
    const bytes = await readFile(file)
    expect(bytes.includes(Buffer.from(SECRET, 'utf8'))).toBe(false)
    expect(bytes.toString('utf8')).not.toContain(SECRET)
    expect(bytes.toString('utf8')).toContain(safeStorage.encryptString(SECRET).toString('base64'))
  })

  it('propagates a read failure that is not absence', async () => {
    await store.set('DEEPSEEK_API_KEY', SECRET)
    const file = await onlyFile(directory)
    await rm(file)
    await mkdir(file)

    await expect(store.get('DEEPSEEK_API_KEY')).rejects.toMatchObject({ code: 'EISDIR' })
  })

  describe('stored document validation', () => {
    const malformedDocuments: Array<[name: string, text: string]> = [
      ['invalid JSON', '{'],
      ['a non-object root', '"text"'],
      ['a null root', 'null'],
      ['an array root', '[]'],
      ['an unknown version', JSON.stringify({ version: 2, entries: {} })],
      ['a non-mapping entries section', JSON.stringify({ version: 1, entries: [] })],
      ['a non-string entry value', JSON.stringify({ version: 1, entries: { KEY: 1 } })],
      ['an empty entry value', JSON.stringify({ version: 1, entries: { KEY: '' } })],
    ]

    it.each(malformedDocuments)('rejects %s', async (_name, text) => {
      await store.set('SEED', 'seed')
      await writeFile(await onlyFile(directory), text)

      await expect(store.get('SEED')).rejects.toThrow()
    })
  })
})
