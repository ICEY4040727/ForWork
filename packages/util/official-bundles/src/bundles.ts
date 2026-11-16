/** The bundle package names the dsh installation selects, offers, installs on demand, or retires. */

/** Installation-owned defaults used when a shipped profile is first opened. */
export interface ProfileTemplate {
  /** Ordered bundle layer list. */
  bundles: readonly string[]
}

/** The shipped profile templates auto-initialized on first use, by name. */
export const PROFILE_TEMPLATES: Record<string, ProfileTemplate> = {
  acp: {
    bundles: ['@forwork/base', '@forwork/acp-app'],
  },
  web: {
    bundles: ['@forwork/base', '@forwork/web-app'],
  },
  headless: {
    bundles: ['@forwork/base', '@forwork/headless'],
  },
  sdk: {
    bundles: ['@forwork/base', '@forwork/sdk-app'],
  },
  'sdk-minimal': {
    bundles: ['@forwork/sdk-minimal'],
  },
}

/**
 * Installation-owned bundle tuples that an earlier release wrote into a shipped profile; app-boot replaces an exact
 * match with the {@link PROFILE_TEMPLATES} entry of the same name and leaves every other list untouched.
 */
export const INSTALLATION_OWNED_PROFILE_TUPLES: Readonly<Record<string, readonly string[]>> = {
  headless: ['@forwork/base', '@forwork/web-app', '@forwork/headless'],
}

/** Bundles an earlier release shipped and the installation no longer carries; app-boot removes them from profiles. */
export const RETIRED_BUNDLES: ReadonlySet<string> = new Set([
  // The Web composition mounts Schedule itself
  // ([upgrade guide](../../../../docs/upgrade-guide/v0.2.0-rc.2/schedule-bundle-retired/guide.md)).
  '@forwork/experimental-schedule-bundle',
  '@forwork/experimental-agent-team-profile',
])

/** The bundle list a `forwork plugin` init uses for a name with no shipped template. */
export const DEFAULT_PROFILE_BUNDLES: readonly string[] = ['@forwork/base']

/**
 * The bundles the dsh installation ships for a person to switch on: each a
 * runtime dependency of the installation that declares `dsh.bundle.patch`,
 * an `icon`, and `./locale/*.json` display metadata, selected by no shipped
 * template, and offered switched off by the plugin manager except persistent
 * terminals, which have no page switch
 * ([rationale](../../../../.agents/notes/implemented/process/2026-09-15-shipped-optional-bundles.md),
 * [admission](../../../../.agents/notes/implemented/architecture/2026-09-21-experimental-capabilities-as-optional-bundles.md)).
 */
export const OPTIONAL_BUNDLES: readonly string[] = [
  '@forwork/personalized-recommendations',
  '@forwork/cua-bundle',
  '@forwork/experimental-session-search',
  '@forwork/experimental-ralph-bundle',
  '@forwork/experimental-terminal-bundle',
  '@forwork/experimental-badge-skill-bundle',
  '@forwork/experimental-session-titles-bundle',
  '@forwork/experimental-voice-input-bundle',
  '@forwork/experimental-cot-translation-bundle',
  '@forwork/experimental-auto-review',
  '@forwork/experimental-inspector-profile',
  '@forwork/experimental-liveshare',
  '@forwork/experimental-plugin-market-bundle',
  '@forwork/experimental-tool-worktree',
  '@forwork/experimental-forge-bundle',
]

/** Public DSH release members available through the ordinary on-demand bundle installer. */
export const ON_DEMAND_BUNDLES: readonly string[] = [
  '@forwork/subagent-claude-code',
  '@forwork/subagent-codex',
]
