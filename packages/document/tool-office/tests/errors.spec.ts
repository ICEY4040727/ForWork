/** Stable failure mapping, guarded execution, and document identity resolution. */
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { OfficeExchangeError } from '@deepseek-ai/dsh-office-exchange'
import { SessionId } from '@deepseek-ai/dsh-session'
import { RemoteError } from '@deepseek-ai/dsh-typert-protocol'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  confinePath,
  confineToolPath,
  documentIdFor,
  documentTarget,
} from '../src/documents.ts'
import { failure, guarded, mapFailure, OfficeToolResultError } from '../src/errors.ts'

const SESSION = SessionId('session-1')
let root: string

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'dsh-tool-office-paths-'))
})
afterEach(async () => {
  await rm(root, { recursive: true, force: true })
})

describe('failure mapping', () => {
  it('keeps a tool refusal and its optional detail', () => {
    expect(mapFailure(new OfficeToolResultError('pending-edits', 'save first'))).toEqual({
      status: 'error',
      code: 'pending-edits',
      message: 'save first',
    })
    expect(failure('conflict', 'stale', { currentVersion: 2 })).toEqual({
      status: 'error',
      code: 'conflict',
      message: 'stale',
      currentVersion: 2,
    })
    expect(failure('unsupported-feature', 'macro', { features: [{ kind: 'macro' }] })).toMatchObject({
      features: [{ kind: 'macro' }],
    })
  })

  it.each([
    ['office-document/conflict', 'conflict'],
    ['office-document/unsupported-feature', 'unsupported-feature'],
    ['office-document/unsupported-format', 'unsupported-feature'],
    ['office-document/out-of-scope', 'out-of-scope'],
    ['office-document/not-found', 'document-missing'],
    ['office-document/bound-exceeded', 'bound-exceeded'],
    ['office-document/conversion-failed', 'conversion-failed'],
  ] as const)('maps Remote %s to %s', (remote, code) => {
    expect(mapFailure(new RemoteError(remote, 'remote failure', { path: 'x' })).code).toBe(code)
  })

  it.each([
    ['gateway/internal', 'conversion-failed'],
    ['gateway/bad-request', 'conversion-failed'],
  ] as const)('maps Remote %s to %s', (remote, code) => {
    expect(mapFailure(new RemoteError(remote, 'remote failure', {})).code).toBe(code)
  })

  it('rethrows a cancelled Remote call so the executor owns cancellation', () => {
    const cause = new RemoteError('gateway/cancelled', 'cancelled', {})
    expect(() => mapFailure(cause)).toThrow(cause)
  })

  it.each([
    ['unsupported-feature', 'unsupported-feature'],
    ['unsupported-format', 'unsupported-feature'],
    ['too-large', 'bound-exceeded'],
    ['out-of-scope', 'out-of-scope'],
    ['conversion-failed', 'conversion-failed'],
    ['not-implemented', 'conversion-failed'],
  ] as const)('maps Exchange %s to %s', (exchange, code) => {
    expect(mapFailure(new OfficeExchangeError(exchange, 'exchange failure')).code).toBe(code)
  })

  it('rethrows an unclassified failure', () => {
    const cause = new Error('unexpected')
    expect(() => mapFailure(cause)).toThrow(cause)
  })

  it('returns a refused step through guarded', async () => {
    await expect(guarded(async () => { throw new OfficeToolResultError('timeout', 'no answer') })).resolves.toEqual({
      status: 'error',
      code: 'timeout',
      message: 'no answer',
    })
    await expect(guarded(async () => ({ status: 'ok' as const, value: 1 }))).resolves.toEqual({ status: 'ok', value: 1 })
  })
})

describe('document identity and confinement', () => {
  it('builds and splits the Session-scoped identity', () => {
    expect(documentIdFor(SESSION, 'reports/q1.xlsx')).toBe('session-1:reports/q1.xlsx')
    const target = documentTarget('session-1:reports/q1.xlsx', SESSION, root)
    expect(target.documentId).toBe('session-1:reports/q1.xlsx')
    expect(target.path).toBe('reports/q1.xlsx')
    expect(target.absolutePath).toBe(join(root, 'reports/q1.xlsx'))
  })

  it('refuses an identity from another Session and an empty path', () => {
    expect(() => documentTarget('session-2:q1.xlsx', SESSION, root)).toThrow(OfficeToolResultError)
    expect(() => documentTarget('session-1:', SESSION, root)).toThrow(/names no path/)
  })

  it('confines paths to the workspace root', async () => {
    await mkdir(join(root, 'reports'), { recursive: true })
    await writeFile(join(root, 'reports', 'q1.xlsx'), 'x')
    await expect(confinePath(join(root, 'reports', 'q1.xlsx'), root)).resolves.toBe(join(root, 'reports', 'q1.xlsx'))
    await expect(confineToolPath('reports/q1.xlsx', root)).resolves.toBe(join(root, 'reports', 'q1.xlsx'))
    await expect(confineToolPath('../outside.xlsx', root)).rejects.toThrow(OfficeExchangeError)
  })
})
