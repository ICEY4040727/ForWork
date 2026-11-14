/** Write-root confinement: real containment, traversal, missing targets, and symlink escapes. */
import { expect, it, onTestFinished } from 'vitest'
import { mkdir, mkdtemp, realpath, rm, symlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { assertWithinRoots } from '../src/authority.ts'
import { OfficeExchangeError } from '../src/errors.ts'

interface Workspace {
  readonly base: string
  readonly root: string
  readonly outside: string
}

async function workspace(): Promise<Workspace> {
  const base = await mkdtemp(join(tmpdir(), 'dsh-office-authority-'))
  onTestFinished(async () => { await rm(base, { recursive: true, force: true }) })
  const root = join(base, 'root')
  const outside = join(base, 'outside')
  await mkdir(root)
  await mkdir(outside)
  return { base, root, outside }
}

async function rejectionOf(promise: Promise<unknown>): Promise<OfficeExchangeError> {
  try {
    await promise
  } catch (error) {
    if (error instanceof OfficeExchangeError) return error
    throw error
  }
  throw new Error('expected the write-root check to reject')
}

it('returns the resolved path for a target inside a root', async () => {
  const { root } = await workspace()
  const file = join(root, 'report.xlsx')
  await writeFile(file, '')
  await expect(assertWithinRoots(file, [root])).resolves.toBe(await realpath(file))
})

it('accepts a root itself as the target', async () => {
  const { root } = await workspace()
  await expect(assertWithinRoots(root, [root])).resolves.toBe(await realpath(root))
})

it('accepts a target inside any of several roots', async () => {
  const { base, root } = await workspace()
  const second = join(base, 'second')
  await mkdir(second)
  const file = join(second, 'report.xlsx')
  await writeFile(file, '')
  await expect(assertWithinRoots(file, [root, second])).resolves.toBe(await realpath(file))
})

it('rejects a target outside every root and names the target and roots', async () => {
  const { root, outside } = await workspace()
  const target = join(outside, 'report.xlsx')
  await writeFile(target, '')
  const error = await rejectionOf(assertWithinRoots(target, [root]))
  expect(error.code).toBe('out-of-scope')
  expect(error.message).toContain(target)
  expect(error.message).toContain(root)
})

it('rejects every target when no roots are configured', async () => {
  const { root } = await workspace()
  const error = await rejectionOf(assertWithinRoots(join(root, 'report.xlsx'), []))
  expect(error.code).toBe('out-of-scope')
})

it('rejects traversal that leaves the root', async () => {
  const { root, outside } = await workspace()
  await writeFile(join(outside, 'report.xlsx'), '')
  const error = await rejectionOf(assertWithinRoots(join(root, '..', 'outside', 'report.xlsx'), [root]))
  expect(error.code).toBe('out-of-scope')
})

it('rejects an absolute sibling whose name shares the root prefix', async () => {
  const { root } = await workspace()
  const sibling = `${root}-sibling`
  await mkdir(sibling)
  const target = join(sibling, 'report.xlsx')
  await writeFile(target, '')
  const error = await rejectionOf(assertWithinRoots(target, [root]))
  expect(error.code).toBe('out-of-scope')
})

it('allows a missing file under an existing root', async () => {
  const { root } = await workspace()
  const target = join(root, 'new.xlsx')
  await expect(assertWithinRoots(target, [root], { allowMissing: true })).resolves.toBe(join(await realpath(root), 'new.xlsx'))
})

it('rejects a missing file only when allowMissing is false', async () => {
  const { root } = await workspace()
  const existing = join(root, 'report.xlsx')
  await writeFile(existing, '')
  await expect(assertWithinRoots(existing, [root], { allowMissing: false })).resolves.toBe(await realpath(existing))
  const missing = join(root, 'new.xlsx')
  const error = await rejectionOf(assertWithinRoots(missing, [root], { allowMissing: false }))
  expect(error.code).toBe('out-of-scope')
  expect(error.message).toContain(missing)
})

it('rejects a missing file below a missing subtree outside the root', async () => {
  const { root, outside } = await workspace()
  const target = join(outside, 'missing', 'deeper', 'new.xlsx')
  const error = await rejectionOf(assertWithinRoots(target, [root]))
  expect(error.code).toBe('out-of-scope')
})

it('resolves a target below a regular file through its nearest existing ancestor', async () => {
  const { root } = await workspace()
  const blocker = join(root, 'blocker')
  await writeFile(blocker, '')
  await expect(assertWithinRoots(join(blocker, 'report.xlsx'), [root])).resolves.toBe(join(await realpath(blocker), 'report.xlsx'))
})

it('rejects a symlink inside the root that points outside', async (context) => {
  const { root, outside } = await workspace()
  await writeFile(join(outside, 'secret.xlsx'), '')
  try {
    await symlink(outside, join(root, 'escape'), process.platform === 'win32' ? 'junction' : 'dir')
  } catch (error) {
    context.skip(`the filesystem refused to create a symlink: ${String(error)}`)
    return
  }
  const error = await rejectionOf(assertWithinRoots(join(root, 'escape', 'secret.xlsx'), [root]))
  expect(error.code).toBe('out-of-scope')
})
