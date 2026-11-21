/**
 * An open classifies the package once through the Exchange read and reuses
 * that report instead of running a second full package inflate.
 */
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Context } from '@forwork/cordis'
import { LocalFileSystem } from '@forwork/fs-local'
import { SessionId } from '@forwork/session/types'
import SessionStore from '@forwork/session'
import type { WorkspaceFileScope } from '@forwork/api-workspace-files'
import { expect, it, vi } from 'vitest'
import { LocalOfficeExchange, OfficeDocumentRemote } from '../src/index.ts'

const counts = vi.hoisted(() => ({ packageScans: 0, preservedScans: 0 }))

vi.mock('../src/detect.ts', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/detect.ts')>()
  return {
    ...actual,
    detectOfficePackage: (...args: Parameters<typeof actual.detectOfficePackage>) => {
      counts.packageScans += 1
      return actual.detectOfficePackage(...args)
    },
    detectPreserved: (...args: Parameters<typeof actual.detectPreserved>) => {
      counts.preservedScans += 1
      return actual.detectPreserved(...args)
    },
  }
})

it('classifies the package once per open and reuses the read report', async () => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-detection-'))
  const ctx = new Context()
  try {
    await ctx.plugin(LocalFileSystem, { cwd: root })
    await ctx.plugin(SessionStore)
    await ctx.plugin(OfficeDocumentRemote, { maxBytes: 1_000_000 })
    await ctx.plugin(LocalOfficeExchange)
    const scope: WorkspaceFileScope = { sessionId: SessionId('remote-detection'), workspaceRoot: root }
    ctx.sessions.create(scope.sessionId, { meta: { cwd: root } })
    const bytes = await readFile(new URL('./fixtures/roundtrip.xlsx', import.meta.url))
    await writeFile(join(root, 'roundtrip.xlsx'), bytes)
    counts.packageScans = 0
    counts.preservedScans = 0

    const read = await ctx.officeDocument.read(
      scope,
      'roundtrip.xlsx',
      new AbortController().signal,
    )

    expect(read.preserved).toEqual(['document-properties', 'theme'])
    expect(counts.packageScans).toBe(1)
    expect(counts.preservedScans).toBe(0)
  } finally {
    await ctx.fiber.dispose()
    await rm(root, { recursive: true, force: true })
  }
})
