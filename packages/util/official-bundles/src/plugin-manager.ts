/** The package names the plugin manager protects, hides, orders, or marks experimental. */

/**
 * Plugin packages the plugin manager refuses to disable or remove, because management itself needs them; a bundle that
 * contains one is read-only as well. Removing a name drops only this `management-required` refusal; the manager's own
 * entry, installation-carried bundles, and unaddressable rows stay protected by their own rules.
 */
export const PROTECTED_PLUGINS: ReadonlySet<string> = new Set([
  '@forwork/plugin-manager', '@forwork/cordis-plugin-loader',
  '@forwork/cordis-plugin-include', '@forwork/api-gateway',
  '@forwork/host-webserver', '@forwork/client-modules',
  '@forwork/client-ui-settings-plugin-inventory', '@forwork/client-ui-plugin-manager',
  '@forwork/host-plugin-inventory', '@forwork/typert-registry',
  '@forwork/api-remotes', '@forwork/api-remote-access',
  '@forwork/cordis-plugin-timer', '@forwork/client-connection',
  '@forwork/host-frontend-static', '@forwork/tools',
  '@forwork/hmr',
])

/**
 * Built-in profile bundles, persistent terminals, and the standalone worktree tool, which have no plugin manager page.
 * Their Host availability and activation stay intact.
 */
export const HIDDEN_BUNDLES: ReadonlySet<string> = new Set([
  '@forwork/base',
  '@forwork/web-app',
  '@forwork/headless',
  '@forwork/sdk-app',
  '@forwork/acp-app',
  '@forwork/sdk-minimal',
  '@forwork/experimental-terminal-bundle',
  '@forwork/experimental-tool-worktree',
])

/**
 * The optional bundles the plugin manager's main and More lists rank first, in this order. {@link MORE_BUNDLES}, not
 * this list, decides whether a bundle appears on the main list.
 */
export const EXTENSION_BUNDLES: readonly string[] = [
  '@forwork/experimental-voice-input-bundle',
  '@forwork/experimental-plugin-market-bundle',
  '@forwork/experimental-auto-review',
]

/**
 * The optional bundles the plugin manager's main list leaves out until the profile enables them; its More list,
 * which includes every eligible bundle, ranks them last in this order.
 */
export const MORE_BUNDLES: readonly string[] = [
  '@forwork/personalized-recommendations',
  '@forwork/experimental-inspector-profile',
  '@forwork/cua-bundle',
  '@forwork/experimental-cot-translation-bundle',
  '@forwork/experimental-session-search',
  '@forwork/experimental-session-titles-bundle',
  '@forwork/experimental-badge-skill-bundle',
  '@forwork/experimental-ralph-bundle',
  '@forwork/subagent-codex',
  '@forwork/subagent-claude-code',
]

/** Official bundles the plugin manager marks with the Experimental badge. */
export const EXPERIMENTAL_BUNDLES: ReadonlySet<string> = new Set([
  '@forwork/personalized-recommendations',
  '@forwork/cua-bundle',
  '@forwork/experimental-voice-input-bundle',
  '@forwork/experimental-auto-review',
  '@forwork/experimental-inspector-profile',
  '@forwork/experimental-forge-bundle',
  '@forwork/experimental-tool-worktree',
  '@forwork/experimental-cot-translation-bundle',
  '@forwork/experimental-session-search',
  '@forwork/experimental-session-titles-bundle',
  '@forwork/experimental-badge-skill-bundle',
  '@forwork/experimental-ralph-bundle',
  '@forwork/experimental-plugin-market-bundle',
  '@forwork/experimental-liveshare',
])
