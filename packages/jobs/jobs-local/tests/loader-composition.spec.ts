import { afterEach, describe, expect, it } from 'vitest'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { Context } from '@forwork/cordis'
import Include from '@forwork/cordis-plugin-include'
import Loader, { type ModuleLoaderV2 } from '@forwork/cordis-plugin-loader'
import LocalJobRegistry from '@forwork/jobs-local'

/**
 * A Loader-internal stub that serves this test's plugin modules. The real Node
 * loader's module machinery is unused by the composition under test.
 * @param load - resolves one plugin specifier.
 * @returns a v2 loader whose `import` calls {@link load}.
 */
function moduleLoaderStub(load: ModuleLoaderV2['import']): ModuleLoaderV2 {
  const unsupported = (): never => { throw new Error('test Loader stub does not implement Node module machinery') }
  return {
    version: 'v2',
    loadCache: {} as ModuleLoaderV2['loadCache'],
    import: load,
    register: unsupported,
    getOrCreateModuleJob: unsupported,
    resolveSync: unsupported,
    load: unsupported,
  }
}

let root: string | undefined
let context: Context | undefined

afterEach(async () => {
  await context?.fiber.dispose()
  context = undefined
  if (root !== undefined) await rm(root, { recursive: true, force: true })
  root = undefined
})

describe('jobs-local through a real Loader composition', () => {
  it('applies the provider-owned admission config from a Cordis row', async () => {
    root = await mkdtemp(join(tmpdir(), 'dsh-jobs-local-loader-'))
    const configPath = join(root, 'cordis.yml')
    await writeFile(configPath, [
      "- name: '@forwork/jobs-local'",
      '  config:',
      '    maxConcurrentJobsPerOwner: 1',
      '',
    ].join('\n'))

    context = new Context()
    context.baseUrl = pathToFileURL(root).href + '/'
    await context.plugin(Loader)
    context.loader.builtins.include = Include
    context.loader.internal = moduleLoaderStub(async (specifier: string) => {
      if (specifier === '@forwork/jobs-local') return LocalJobRegistry
      throw new Error(`unexpected Loader import: ${specifier}`)
    })
    await context.loader.create({
      name: 'cordis:include',
      config: { path: pathToFileURL(configPath).href },
    })
    await context.loader.await()

    expect(context.jobs).toBeInstanceOf(LocalJobRegistry)
    context.jobs.attachController('loader-test')
    let settle!: (outcome: { status: 'killed' }) => void
    context.jobs.start({
      kind: 'bash',
      label: 'hold loader slot',
      run: () => ({
        cancel: () => { settle({ status: 'killed' }) },
        done: new Promise((resolve) => { settle = resolve }),
      }),
    })
    expect(() => context!.jobs.start({
      kind: 'bash',
      label: 'blocked loader job',
      run: () => ({ cancel: () => {}, done: Promise.resolve({ status: 'completed' }) }),
    })).toThrow('(limit: 1)')
  })
})
