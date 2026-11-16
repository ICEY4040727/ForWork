/**
 * Owner-side Liveshare Web bundle plugin.
 * @module @forwork/experimental-liveshare
 */

export { apply, name, inject, type Config } from './host/share-host.ts'
export { LiveshareShares } from './host/share-remote.ts'

import type { LiveshareShares } from './host/share-remote.ts'

declare module '@forwork/cordis' {
  interface Context {
    /** Owner share management served to the owner page over Typert Remote. */
    liveshareShares: LiveshareShares
  }
}
