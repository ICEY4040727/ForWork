import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'
import { Context } from '@forwork/cordis'
import Loader, { type ModuleLoaderV2 } from '@forwork/cordis-plugin-loader'
import Include from '@forwork/cordis-plugin-include'
import SessionProjectionRegistry from '@forwork/session-projection'
import TokenMeter from '@forwork/token-meter'
import ToolResultPruner from '@forwork/compaction-tool-result-pruner'

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

describe('compaction-tool-result-pruner real Loader composition', () => {
  it('loads and resolves the flat YAML plugin shape', async () => {
    root = await mkdtemp(join(tmpdir(), 'dsh-compact-tool-result-prune-loader-'))
    const configPath = join(root, 'cordis.yml')
    await writeFile(configPath, [
      "- name: '@forwork/session-projection'",
      "- name: '@forwork/token-meter'",
      "- name: '@forwork/compaction-tool-result-pruner'",
      '  config:',
      '    thresholdChars: 100',
      '    headChars: 20',
      '    tailChars: 10',
      '',
    ].join('\n'))

    context = new Context()
    context.baseUrl = pathToFileURL(root).href + '/'
    await context.plugin(Loader)
    context.loader.builtins.include = Include
    context.loader.internal = moduleLoaderStub(async (specifier: string) => {
      if (specifier === '@forwork/session-projection') return SessionProjectionRegistry
      if (specifier === '@forwork/token-meter') return TokenMeter
      if (specifier === '@forwork/compaction-tool-result-pruner') return ToolResultPruner
      throw new Error(`unexpected Loader import: ${specifier}`)
    })
    await context.loader.create({
      name: 'cordis:include',
      config: { path: pathToFileURL(configPath).href },
    })
    await context.loader.await()

    expect(context.get('toolResultPruner')).toBeInstanceOf(ToolResultPruner)
    expect(context.toolResultPruner.config).toEqual({
      thresholdChars: 100,
      headChars: 20,
      tailChars: 10,
    })
  })

  it('rejects stale config after plugin schema normalization', async () => {
    context = new Context()
    // Satisfy the declared injections first: config normalization runs in the
    // service constructor, which a pending fiber never reaches.
    await context.plugin(SessionProjectionRegistry)
    await context.plugin(TokenMeter)
    await expect(context.plugin(ToolResultPruner, {
      maxChars: 100,
    } as never)).rejects.toThrow(/unknown key "maxChars"/)
  })
})
