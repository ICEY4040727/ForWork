/**
 * Desktop credential provider: model references live in Electron's OS secure
 * store, while account records and ambient fallbacks stay in the profile's
 * local credential document.
 *
 * The store backend runs in the Electron main process, where `safeStorage`
 * exists; this module only speaks the private Host bridge. A reference read or
 * write that cannot reach the OS secure store never falls back to plaintext.
 * @module @forwork/desktop-host/secure-credentials
 */

import { readFile } from 'node:fs/promises'
import { Context, Service } from '@forwork/cordis'
import z from '@forwork/schemastery'
import {
  CredentialProvider,
  SecureStorageUnavailableError,
  credentialRef,
  type CredentialInfo,
  type CredentialKey,
  type CredentialRecord,
  type CredentialRecordEntry,
  type CredentialRecordInfo,
  type CredentialRef,
  type ResolvedCredential,
  type SecureStore,
} from '@forwork/credentials'
import LocalCredentialProvider, {
  parseCredentialsDocument,
  resolveSpec,
  type Config as LocalConfig,
} from '@forwork/credentials-local'
import { launchEnvironmentOf } from '@forwork/launch-environment'

/** Name of the private Host bridge messages carrying secure-store operations. */
const SECURE_STORE_REQUEST = 'secure-store-request'
const SECURE_STORE_RESPONSE = 'secure-store-response'

/** A secure-store operation unanswered this long fails instead of hanging a model request. */
const BRIDGE_TIMEOUT_MS = 30_000

/** One response the Electron main process sends for a secure-store request. */
interface BridgeResponse {
  readonly type: typeof SECURE_STORE_RESPONSE
  readonly requestId: number
  readonly ok: boolean
  readonly value?: string
  readonly code?: string
}

/** Whether an incoming message is a well-formed bridge response. */
function isBridgeResponse(value: unknown): value is BridgeResponse {
  if (typeof value !== 'object' || value === null) return false
  const candidate = value as Record<string, unknown>
  return candidate.type === SECURE_STORE_RESPONSE
    && Number.isSafeInteger(candidate.requestId)
    && typeof candidate.ok === 'boolean'
    && (candidate.value === undefined || typeof candidate.value === 'string')
    && (candidate.code === undefined || typeof candidate.code === 'string')
}

/** Whether a failure is the typed refusal raised while the OS secure store is unavailable. */
function isSecureStorageUnavailable(error: unknown): boolean {
  return error instanceof SecureStorageUnavailableError
    || (typeof error === 'object' && error !== null && 'code' in error
      && error.code === 'secure-storage-unavailable')
}

/** Send one request to the Electron main process; the callback reports a transport failure. */
function sendToMain(message: object, onError: (error: Error) => void): void {
  if (!process.connected || process.send === undefined) {
    onError(new Error('secure credentials: the Desktop secure-store bridge is unavailable'))
    return
  }
  try {
    process.send(message, (error) => {
      if (error !== null) onError(error)
    })
  } catch (error) {
    onError(error instanceof Error ? error : new Error(String(error)))
  }
}

/** {@link SecureStore} over the private Host bridge to Electron's `safeStorage`. */
export class HostSecureStore implements SecureStore {
  private sequence = 0
  private disposed = false
  private readonly pending = new Map<number, {
    readonly resolve: (value: string | undefined) => void
    readonly reject: (error: Error) => void
    readonly timer: ReturnType<typeof setTimeout>
  }>()

  private readonly onMessage = (message: unknown): void => {
    if (!isBridgeResponse(message)) return
    const request = this.pending.get(message.requestId)
    if (request === undefined) return
    this.pending.delete(message.requestId)
    clearTimeout(request.timer)
    if (message.ok) {
      request.resolve(message.value)
      return
    }
    if (message.code === 'secure-storage-unavailable') {
      request.reject(new SecureStorageUnavailableError())
      return
    }
    request.reject(new Error(`secure credentials: the OS secure store refused an operation (${message.code ?? 'unknown'})`))
  }

  /** @param timeoutMs - deadline for one main-process answer. */
  constructor(private readonly timeoutMs = BRIDGE_TIMEOUT_MS) {
    process.on('message', this.onMessage)
  }

  /**
   * Read one stored value.
   * @param key - the entry to read.
   * @returns the plaintext value, or `undefined` while absent.
   * @throws SecureStorageUnavailableError when the OS secure store is unavailable.
   */
  get(key: string): Promise<string | undefined> {
    return this.request('get', key)
  }

  /**
   * Store one value.
   * @param key - the entry to write.
   * @param value - the plaintext secret to protect.
   * @throws SecureStorageUnavailableError when the OS secure store is unavailable.
   */
  set(key: string, value: string): Promise<void> {
    return this.request('set', key, value).then(() => undefined)
  }

  /**
   * Remove one entry.
   * @param key - the entry to remove.
   */
  delete(key: string): Promise<void> {
    return this.request('delete', key).then(() => undefined)
  }

  /** Detach from the bridge and fail every operation still waiting for main. */
  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    process.off('message', this.onMessage)
    for (const request of this.pending.values()) {
      clearTimeout(request.timer)
      request.reject(new Error('secure credentials: the Desktop Host is stopping'))
    }
    this.pending.clear()
  }

  /** Issue one correlated bridge request and settle it on the matching response. */
  private request(action: 'get' | 'set' | 'delete', key: string, value?: string): Promise<string | undefined> {
    if (this.disposed) return Promise.reject(new Error('secure credentials: the secure-store bridge is disposed'))
    const requestId = ++this.sequence
    return new Promise<string | undefined>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(requestId)
        reject(new Error('secure credentials: the OS secure store did not answer'))
      }, this.timeoutMs)
      timer.unref()
      this.pending.set(requestId, { resolve, reject, timer })
      sendToMain({
        type: SECURE_STORE_REQUEST,
        requestId,
        action,
        key,
        ...(value === undefined ? {} : { value }),
      }, (error) => {
        this.pending.delete(requestId)
        clearTimeout(timer)
        reject(error)
      })
    })
  }
}

/** Plugin configuration; the local document holds account records and ambient fallbacks. */
export interface Config {
  /** Credentials document path; defaults to `.credentials.yaml` under the harness home. */
  path?: string
  /** Harness home used when `path` is omitted; defaults to `$DSH_HOME` or `~/.dsh`. */
  dshHome?: string
  /** Watcher write-settle window in milliseconds; defaults to 100. */
  debounceMs?: number
}

/**
 * The Desktop profile's credential provider.
 *
 * Model references are written only through {@link HostSecureStore}. Account
 * records, environment fallbacks, and legacy reference migration use the local
 * provider mounted in an isolated service scope, so the two halves never
 * contend for `ctx.credentials`.
 */
export default class DesktopSecureCredentials extends CredentialProvider {
  static Config: z<Config> = z.object({
    path: z.string(),
    dshHome: z.string(),
    debounceMs: z.number().min(0).default(100),
  })

  private readonly localContext: Context
  private readonly ready: Promise<void>
  private local!: LocalCredentialProvider

  /**
   * @param ctx - the profile scope that receives `ctx.credentials`.
   * @param config - legacy document location.
   * @param store - an already-constructed store; the default bridges to Electron main.
   */
  constructor(ctx: Context, public config: Config = {}, private readonly store: SecureStore = new HostSecureStore()) {
    super(ctx)
    this.localContext = ctx.isolate('credentials')
    this.ready = this.initialize()
  }

  /** Mount the local provider and move any legacy plaintext references into the secure store. */
  private async initialize(): Promise<void> {
    const localConfig: LocalConfig = {
      ...(this.config.path === undefined ? {} : { path: this.config.path }),
      ...(this.config.dshHome === undefined ? {} : { dshHome: this.config.dshHome }),
      watch: false,
      debounceMs: this.config.debounceMs ?? 100,
    }
    await this.localContext.plugin(LocalCredentialProvider, localConfig)
    const local = this.localContext.get('credentials')
    if (!(local instanceof LocalCredentialProvider)) {
      throw new Error('desktop secure credentials: the local credential provider did not mount')
    }
    this.local = local
    await this.migrateLegacyReferences()
  }

  /**
   * Move references written by earlier builds out of the plaintext document.
   * An unavailable store leaves them in place rather than blocking Host boot;
   * {@link resolve} retries the move and never serves a legacy file value.
   */
  private async migrateLegacyReferences(): Promise<void> {
    const spec = resolveSpec({ ...this.config, watch: false })
    let text: string
    try {
      text = await readFile(spec.filename, 'utf8')
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return
      throw error
    }
    const document = parseCredentialsDocument(text, spec.filename)
    for (const [name, value] of document.refs) {
      try {
        await this.store.set(name, value)
      } catch (error) {
        if (!isSecureStorageUnavailable(error)) throw error
        this.ctx.logger.warn('desktop secure credentials: the OS secure store is unavailable; legacy references stay in the local document')
        return
      }
      // The secure copy is committed first; a failed removal keeps a value the next boot retries.
      await this.local.unset(credentialRef(name)).catch((error: unknown) => {
        this.ctx.logger.warn('desktop secure credentials: could not remove a migrated reference from the local document')
        this.ctx.logger.warn(error)
      })
    }
  }

  /** The inherited-environment value for a reference, or `undefined` when empty or unset. */
  private ambient(ref: CredentialRef): ResolvedCredential | undefined {
    const environment = launchEnvironmentOf(this.ctx)
    const inherited = environment.getFrom(ref, ['process'])
    if (inherited !== undefined && inherited.value.length > 0) return { value: inherited.value, source: 'env' }
    const fallback = environment.getFrom(ref, ['project-env', 'user-env'])
    if (fallback !== undefined && fallback.value.length > 0) {
      return { value: fallback.value, source: fallback.source }
    }
    return undefined
  }

  /**
   * Resolve one reference from the secure store, then ambient fallbacks.
   * @param ref - the reference to resolve.
   * @returns the value and its source, or `undefined` while unconfigured.
   */
  override async resolve(ref: CredentialRef): Promise<ResolvedCredential | undefined> {
    await this.ready
    try {
      const stored = await this.store.get(ref)
      if (stored !== undefined) return { value: stored, source: 'secure-store' }
    } catch (error) {
      if (!isSecureStorageUnavailable(error)) throw error
      return this.ambient(ref)
    }
    const legacy = await this.local.resolve(ref)
    if (legacy?.source !== 'file') return legacy
    // A late migration or a document edit: never serve the plaintext copy.
    try {
      await this.store.set(ref, legacy.value)
    } catch (error) {
      if (!isSecureStorageUnavailable(error)) throw error
      return this.ambient(ref)
    }
    await this.local.unset(ref).catch(() => undefined)
    return { value: legacy.value, source: 'secure-store' }
  }

  /**
   * Describe one reference without exposing its value.
   * @param ref - the reference to describe.
   * @returns configured state, supplying source, and writability.
   */
  override async describe(ref: CredentialRef): Promise<CredentialInfo> {
    await this.ready
    try {
      const stored = await this.store.get(ref)
      if (stored !== undefined) return { configured: true, source: 'secure-store', writable: true }
    } catch (error) {
      if (!isSecureStorageUnavailable(error)) throw error
      const ambient = this.ambient(ref)
      return ambient === undefined
        ? { configured: false, writable: false }
        : { configured: true, source: ambient.source, writable: false }
    }
    const local = await this.local.describe(ref)
    return { ...local, writable: true }
  }

  /**
   * Durably store one model reference in the OS secure store.
   * @param ref - the reference to store.
   * @param value - the non-empty secret value.
   */
  override async set(ref: CredentialRef, value: string): Promise<void> {
    if (value.length === 0) {
      throw new Error(`desktop secure credentials: an empty value cannot be stored for "${ref}"; use unset`)
    }
    await this.ready
    await this.store.set(ref, value)
    await this.local.unset(ref).catch((error: unknown) => {
      this.ctx.logger.warn('desktop secure credentials: could not clear a superseded local reference')
      this.ctx.logger.warn(error)
    })
    this.notifyUpdated(ref)
  }

  /**
   * Remove one model reference.
   * @param ref - the reference to remove.
   */
  override async unset(ref: CredentialRef): Promise<void> {
    await this.ready
    await this.store.delete(ref)
    await this.local.unset(ref).catch((error: unknown) => {
      this.ctx.logger.warn('desktop secure credentials: could not clear a superseded local reference')
      this.ctx.logger.warn(error)
    })
    this.notifyUpdated(ref)
  }

  /**
   * Read one stored record from the local credential document.
   * @param key - the record to read.
   * @returns the record, or `undefined` while none is stored.
   */
  override async readRecord(key: CredentialKey): Promise<CredentialRecord | undefined> {
    await this.ready
    return this.local.readRecord(key)
  }

  /**
   * Describe one stored record without exposing its value.
   * @param key - the record to describe.
   * @returns presence, discriminant, and writability.
   */
  override async describeRecord(key: CredentialKey): Promise<CredentialRecordInfo> {
    await this.ready
    return this.local.describeRecord(key)
  }

  /**
   * Enumerate every stored record's address and tag.
   * @returns every stored record, values excluded.
   */
  override async listRecords(): Promise<readonly CredentialRecordEntry[]> {
    await this.ready
    return this.local.listRecords()
  }

  /**
   * Serialized read-modify-write over one local record.
   * @param key - the record to modify.
   * @param mutate - receives the current record and returns its replacement, or `undefined` to leave it.
   * @returns the record after the write, or the current one when `mutate` declined.
   */
  override async modifyRecord(
    key: CredentialKey,
    mutate: (current: CredentialRecord | undefined) => Promise<CredentialRecord | undefined>,
  ): Promise<CredentialRecord | undefined> {
    await this.ready
    return this.local.modifyRecord(key, mutate)
  }

  /**
   * Remove one local record.
   * @param key - the record to remove.
   */
  override async deleteRecord(key: CredentialKey): Promise<void> {
    await this.ready
    await this.local.deleteRecord(key)
  }

  /** Detach the bridge when this provider unloads. */
  async *[Service.init](): AsyncGenerator<() => void, void, void> {
    const store = this.store
    if (store instanceof HostSecureStore) yield () => { store.dispose() }
    await this.ready
  }
}
