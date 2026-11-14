/**
 * OS-secure-store adapter for model credentials. The seam is a three-method
 * key/value store; this module provides the Electron-backed implementation
 * over an injected `SafeStorageLike`, so the caller supplies the app-data
 * directory and this package never imports Electron itself.
 *
 * Only `encryptString` output reaches disk, base64-encoded in one JSON file.
 * Encryption availability is checked on every `get` and `set`; an unavailable
 * store raises {@link SecureStorageUnavailableError} instead of falling back to
 * plaintext.
 * @module @deepseek-ai/dsh-credentials/secure-store
 */

import { randomUUID } from 'node:crypto'
import { mkdir, readFile, rename, rm, writeFile } from 'node:fs/promises'
import { join } from 'node:path'

/** Basename of the ciphertext document inside the supplied directory. */
const SECURE_STORE_FILENAME = 'secure-credentials.json'

/** The document layout this build reads and writes. */
const SECURE_STORE_VERSION = 1

/**
 * Stable discriminant of a failed secure-store operation.
 */
export type SecureStoreErrorCode = 'secure-storage-unavailable'

/** The subset of Electron's `safeStorage` this adapter needs. */
export interface SafeStorageLike {
  /**
   * Whether the OS secure store can encrypt and decrypt on this host.
   * @returns true when encryption is available.
   */
  isEncryptionAvailable(): boolean

  /**
   * Encrypt a plaintext secret with an OS-managed key.
   * @param plainText - the secret to encrypt.
   * @returns the ciphertext.
   */
  encryptString(plainText: string): Buffer

  /**
   * Decrypt bytes produced by {@link SafeStorageLike.encryptString}.
   * @param encrypted - the ciphertext.
   * @returns the original plaintext.
   */
  decryptString(encrypted: Buffer): string
}

/**
 * Minimal key/value seam for secrets that must be protected by the operating
 * system at rest. Keys are opaque strings; values are plaintext only in
 * memory.
 */
export interface SecureStore {
  /**
   * Read one stored value.
   * @param key - the entry to read.
   * @returns the plaintext value, or `undefined` while absent.
   * @throws SecureStorageUnavailableError when the OS secure store is unavailable.
   */
  get(key: string): Promise<string | undefined>

  /**
   * Store one value, replacing any current value at the same key.
   * @param key - the entry to write.
   * @param value - the plaintext secret to protect.
   * @throws SecureStorageUnavailableError when the OS secure store is unavailable.
   */
  set(key: string, value: string): Promise<void>

  /**
   * Remove one stored value; removing an absent key is a no-op.
   * @param key - the entry to remove.
   */
  delete(key: string): Promise<void>
}

/** Raised instead of writing plaintext when the OS secure store is unavailable. */
export class SecureStorageUnavailableError extends Error {
  /** Machine-readable discriminant for callers that branch on the failure. */
  readonly code: SecureStoreErrorCode = 'secure-storage-unavailable'

  constructor() {
    super('secure-storage-unavailable: the operating system secure store is unavailable')
    this.name = 'SecureStorageUnavailableError'
  }
}

/** Construction options for {@link SafeStorageSecureStore}. */
export interface FileSecureStoreOptions {
  /** Electron `safeStorage` or a test double with the same surface. */
  safeStorage: SafeStorageLike
  /** Application-data directory that receives the ciphertext document. */
  directory: string
}

/**
 * {@link SecureStore} backed by Electron `safeStorage` and one ciphertext
 * document under the supplied application-data directory.
 *
 * Each entry is the base64 encoding of `safeStorage.encryptString(value)`;
 * plaintext is never serialized. Operations are serialized within one
 * instance so a read-modify-write cannot interleave with another operation on
 * the same document. `delete` works without encryption because removing
 * ciphertext needs no key, so a locked keychain does not block clearing a
 * stored secret.
 */
export class SafeStorageSecureStore implements SecureStore {
  private readonly safeStorage: SafeStorageLike
  private readonly directory: string
  private readonly filename: string
  private operations: Promise<void> = Promise.resolve()

  /** @param options - the injected safe storage and the directory to persist under. */
  constructor(options: FileSecureStoreOptions) {
    this.safeStorage = options.safeStorage
    this.directory = options.directory
    this.filename = join(options.directory, SECURE_STORE_FILENAME)
  }

  /**
   * Read and decrypt one entry.
   * @param key - the entry to read.
   * @returns the plaintext value, or `undefined` while absent.
   * @throws SecureStorageUnavailableError when the OS secure store is unavailable.
   */
  get(key: string): Promise<string | undefined> {
    return this.enqueue(async () => {
      this.assertAvailable()
      const entries = await this.readEntries()
      const encrypted = entries.get(key)
      if (encrypted === undefined) return undefined
      return this.safeStorage.decryptString(Buffer.from(encrypted, 'base64'))
    })
  }

  /**
   * Encrypt and persist one entry, replacing any current value at the key.
   * @param key - the entry to write.
   * @param value - the plaintext secret to protect.
   * @throws SecureStorageUnavailableError when the OS secure store is unavailable.
   */
  set(key: string, value: string): Promise<void> {
    return this.enqueue(async () => {
      this.assertAvailable()
      const encrypted = this.safeStorage.encryptString(value)
      const entries = await this.readEntries()
      entries.set(key, encrypted.toString('base64'))
      await this.writeEntries(entries)
    })
  }

  /**
   * Remove one entry; removing an absent key is a no-op.
   * @param key - the entry to remove.
   */
  delete(key: string): Promise<void> {
    return this.enqueue(async () => {
      const entries = await this.readEntries()
      if (!entries.has(key)) return
      entries.delete(key)
      await this.writeEntries(entries)
    })
  }

  /** Refuse an operation that would otherwise need a plaintext fallback. */
  private assertAvailable(): void {
    if (!this.safeStorage.isEncryptionAvailable()) throw new SecureStorageUnavailableError()
  }

  /** Queue one exclusive operation behind every earlier operation on this instance. */
  private enqueue<T>(operation: () => Promise<T>): Promise<T> {
    const task = this.operations.then(operation)
    this.operations = task.then(() => undefined, () => undefined)
    return task
  }

  /** Read the ciphertext document, treating an absent file as an empty store. */
  private async readEntries(): Promise<Map<string, string>> {
    let text: string
    try {
      text = await readFile(this.filename, 'utf8')
    } catch (error) {
      if (isENOENT(error)) return new Map()
      throw error
    }
    return parseEntries(text, this.filename)
  }

  /** Replace the ciphertext document through a mode-restricted temporary file. */
  private async writeEntries(entries: Map<string, string>): Promise<void> {
    await mkdir(this.directory, { recursive: true, mode: 0o700 })
    const document = JSON.stringify(
      { version: SECURE_STORE_VERSION, entries: Object.fromEntries(entries) },
      null,
      2,
    )
    const temporary = `${this.filename}.${process.pid}.${randomUUID()}.tmp`
    try {
      await writeFile(temporary, `${document}\n`, { encoding: 'utf8', mode: 0o600, flag: 'wx' })
      await rename(temporary, this.filename)
    } finally {
      await rm(temporary, { force: true })
    }
  }
}

/** Whether a filesystem error means absence; every non-ENOENT failure must surface. */
function isENOENT(error: unknown): boolean {
  return (error as NodeJS.ErrnoException | null)?.code === 'ENOENT'
}

/**
 * Parse the ciphertext document. The file is a durable boundary written only
 * by this adapter, so a malformed document fails loud instead of reading as an
 * empty store. Diagnostics name the file and the offending key, never a value.
 * @param text - the document's text.
 * @param filename - absolute path, quoted in errors.
 * @returns the stored ciphertext entries.
 */
function parseEntries(text: string, filename: string): Map<string, string> {
  let raw: unknown
  try {
    raw = JSON.parse(text)
  } catch {
    throw new Error(`secure-store: ${filename} is not valid JSON`)
  }
  if (!isRecord(raw)) throw new TypeError(`secure-store: ${filename} must hold a mapping`)
  if (raw['version'] !== SECURE_STORE_VERSION) {
    throw new Error(
      `secure-store: ${filename} declares version ${JSON.stringify(raw['version'])};`
      + ` this build reads version ${SECURE_STORE_VERSION}`,
    )
  }
  const section = raw['entries']
  if (!isRecord(section)) throw new TypeError(`secure-store: "entries" in ${filename} must be a mapping`)
  const entries = new Map<string, string>()
  for (const [key, value] of Object.entries(section)) {
    if (typeof value !== 'string' || value.length === 0) {
      throw new TypeError(`secure-store: the value for "${key}" in ${filename} must be a non-empty string`)
    }
    entries.set(key, value)
  }
  return entries
}

/** Whether a parsed JSON value is a plain mapping. */
function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
