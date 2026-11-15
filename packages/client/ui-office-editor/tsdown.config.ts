import type { UserConfig } from 'tsdown'
import { clientBundle } from '../tsdown.client.ts'

const bundle = clientBundle(import.meta.dirname, ['lib/types/index.js'])

export default (options: Parameters<typeof bundle>[0]): UserConfig[] => bundle(options).map(config =>
  config.name?.endsWith('/client') === true
    ? {
        ...config,
        // The Univer presets reach CommonJS-only React helpers through the
        // default `main` field. Rolldown then adds the CommonJS interop helper
        // to the dynamically imported runtime chunk and requires that chunk
        // synchronously from the entry, which the client module table cannot
        // answer. Resolving each dependency's browser ESM build keeps the
        // runtime chunk free of those wrappers.
        inputOptions: {
          ...config.inputOptions,
          resolve: { ...config.inputOptions?.resolve, mainFields: ['browser', 'module', 'main'] },
        },
      }
    : config,
)
