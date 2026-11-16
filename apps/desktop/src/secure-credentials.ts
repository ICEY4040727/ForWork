/**
 * Electron-owned access to the OS secure store for Desktop model credentials.
 *
 * The store itself belongs to `@forwork/credentials`; this module only injects
 * Electron's `safeStorage` and the application's `userData` directory, then
 * answers the private Host bridge. Only ciphertext reaches disk.
 * @module secure-credentials
 */

import { join } from 'node:path'
import {
  SafeStorageSecureStore,
  SecureStorageUnavailableError,
  type SafeStorageLike,
} from '@forwork/credentials'

/** One secure-store operation requested by the Desktop Host. */
export interface SecureStoreRequest {
  readonly type: 'secure-store-request'
  readonly requestId: number
  readonly action: 'get' | 'set' | 'delete'
  readonly key: string
  readonly value?: string
}

/** The Host bridge's answer; `code` is present only when the operation failed. */
export interface SecureStoreResponse {
  readonly type: 'secure-store-response'
  readonly requestId: number
  readonly ok: boolean
  readonly value?: string
  readonly code?: string
}

/** Whether a failure is the typed "the OS secure store is unavailable" refusal. */
function isSecureStorageUnavailable(error: unknown): boolean {
  return error instanceof SecureStorageUnavailableError
    || (typeof error === 'object' && error !== null && 'code' in error
      && error.code === 'secure-storage-unavailable')
}

/** Whether a value is a well-formed Host bridge request. */
export function isSecureStoreRequest(value: unknown): value is SecureStoreRequest {
  if (typeof value !== 'object' || value === null) return false
  const candidate = value as Record<string, unknown>
  if (candidate.type !== 'secure-store-request' || !Number.isSafeInteger(candidate.requestId)) return false
  if (candidate.action !== 'get' && candidate.action !== 'set' && candidate.action !== 'delete') return false
  if (typeof candidate.key !== 'string' || candidate.key.length === 0) return false
  if (candidate.action === 'set') return typeof candidate.value === 'string' && candidate.value.length > 0
  return candidate.value === undefined
}

/**
 * Desktop-owned `SafeStorageSecureStore` plus the Host bridge that reaches it.
 *
 * The ciphertext document lives at `userData/desktop/secure-credentials.json`,
 * beside the other Desktop-owned user preferences.
 */
export class DesktopSecureCredentials {
  /** Ciphertext store shared with the Desktop Host over private IPC. */
  readonly store: SafeStorageSecureStore

  /**
   * @param safeStorage - Electron `safeStorage` or a test double with the same surface.
   * @param userData - Electron `app.getPath('userData')`; never a Renderer-supplied path.
   */
  constructor(private readonly safeStorage: SafeStorageLike, userData: string) {
    this.store = new SafeStorageSecureStore({
      safeStorage,
      directory: join(userData, 'desktop'),
    })
  }

  /**
   * Whether the OS secure store can protect a value right now. A `false`
   * answer means model credentials cannot be saved without a plaintext
   * fallback, which this application never performs.
   * @returns `safeStorage.isEncryptionAvailable()` for the current process.
   */
  available(): boolean {
    return this.safeStorage.isEncryptionAvailable()
  }

  /**
   * Answer one Host bridge request.
   * @param request - a validated request from the Desktop Host.
   * @returns the operation value or a stable failure code; values never enter diagnostics.
   */
  async handle(request: SecureStoreRequest): Promise<SecureStoreResponse> {
    try {
      switch (request.action) {
        case 'get': {
          const value = await this.store.get(request.key)
          return value === undefined
            ? { type: 'secure-store-response', requestId: request.requestId, ok: true }
            : { type: 'secure-store-response', requestId: request.requestId, ok: true, value }
        }
        case 'set':
          await this.store.set(request.key, request.value ?? '')
          return { type: 'secure-store-response', requestId: request.requestId, ok: true }
        case 'delete':
          await this.store.delete(request.key)
          return { type: 'secure-store-response', requestId: request.requestId, ok: true }
      }
    } catch (error) {
      if (isSecureStorageUnavailable(error)) {
        return {
          type: 'secure-store-response',
          requestId: request.requestId,
          ok: false,
          code: 'secure-storage-unavailable',
        }
      }
      // The store's diagnostics name paths and keys, never secret values.
      console.error('desktop secure credentials: secure store operation failed', error)
      return { type: 'secure-store-response', requestId: request.requestId, ok: false, code: 'secure-store-failed' }
    }
  }
}
