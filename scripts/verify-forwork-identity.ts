/** Reject DeepSeek Harness identity tokens from shipped surfaces. */

import { readdirSync, readFileSync } from 'node:fs'
import { extname, join, relative, resolve, sep } from 'node:path'

const root = resolve(import.meta.dirname, '..')

/** One forbidden identity token and its documented exceptions. */
interface IdentityRule {
  id: string
  label: string
  pattern: RegExp
  /** Matched tokens equal to one of these literals are known, reviewed exceptions. */
  allow?: readonly string[]
  /** Lines containing one of these literals are known, reviewed exceptions. */
  allowLines?: readonly string[]
  /**
   * Repository-relative files or directory prefixes whose matches this rule
   * admits after review. Keep every entry narrow and name its owner here.
   */
  allowPaths?: readonly string[]
}

/** Known third-party scope that keeps the upstream DeepSeek package name; see
 * ../.agents/notes/implemented/process/2026-11-17-forwork-identity-exceptions.md for the exception. */
const ALLOWED_DEEPSEEK_PACKAGES = ['@deepseek-ai/libreoffice-kit'] as const

/** Python distribution and module names stay `deepseek_harness*` until T102
 * renames them; see the exceptions note. */
const T102_PYTHON_PATHS = [
  'python/sdk-runtime/hatch_build.py',
  'python/sdk-runtime/platforms.json',
  'python/sdk-runtime/pyproject.toml',
  'python/sdk-runtime/runtime-bootstrap.mjs',
  'python/sdk-runtime/src/deepseek_harness_runtime/__init__.py',
  'python/sdk-runtime/src/deepseek_harness_runtime/_resources.py',
  'python/sdk-runtime/src/deepseek_harness_runtime/deepseek-harness-runtime.json',
  'python/sdk/pyproject.toml',
  'python/sdk/src/deepseek_harness/client.py',
  'python/sdk/tests/manual_sdk_agent_smoke.py',
  'python/sdk/tests/test_bundled_runtime.py',
  'python/sdk/tests/test_client.py',
  'python/sdk/tests/test_release_version.py',
  'python/sdk/tests/test_runtime_resolution.py',
  'docs/user/guide/python-sdk.md',
  'docs/user/guide/python-sdk.zh.md',
] as const

/** The github-review integration watches the upstream repository by design; the
 * guide documents that configuration rather than this product's identity. */
const GITHUB_REVIEW_PATHS = [
  'docs/user/guide/github-review.md',
  'docs/user/guide/github-review.zh.md',
] as const

/** Required citation source: the published citation and the historical evidence path
 * recorded in a postmortem. */
const CITATION_LINES = [
  '@misc{deepseek-harness2026',
  'howpublished={\\url{https://github.com/deepseek-ai/deepseek-harness}}',
] as const

/** The published citation's author field is part of the same citation. */
const CITATION_AUTHOR_LINES = ['author={DeepSeek-AI}'] as const

/** The user-visible default workspace directory keeps its current name until a
 * data migration renames it; see the exceptions note. */
const WORKSPACE_DIRECTORY_LINES = [
  "paths.join(directory, 'deepseek-harness', DEFAULT_WORKSPACE_DIRECTORY)",
  '`<Documents>/deepseek-harness`',
] as const

/** DeepSeek provider packages and the client copy that names the DeepSeek
 * account, owned by T102 until those packages are renamed. */
const T102_PROVIDER_PATHS = [
  'packages/client/ui-chat/',
  'packages/client/ui-model-selection/',
  'packages/client/ui-settings-account/',
  'packages/client/ui-settings-models/',
  'packages/client/ui-settings-session-log/',
  'packages/client/ui-settings-web-search/',
  'packages/credentials/deepseek-account-platform/',
  'packages/experimental/client-ui-cot-translation/',
  'packages/llm/deepseek-llm-api-extensions/',
  'packages/llm/llm-deepseek-account/',
  'packages/llm/llm-deepseek-api-key/',
  'packages/llm/llm-deepseek/',
  'packages/llm/llm-pi-ai/',
  'packages/llm/plugin-package-inventory-deepseek/',
  'packages/session/session-log-deepseek/',
  'packages/web/web-search-deepseek/',
] as const

/** Source comments and configuration descriptions that name the external
 * DeepSeek model, search, and web services. T102 owns renaming the provider;
 * these files only describe its API or protocol. */
const PROVIDER_REFERENCE_PATHS = [
  'apps/desktop/src/welcome-backend.ts',
  'packages/bundle/base/cordis.patch.yml',
  'packages/bundle/sdk-minimal/package.json',
  'packages/client/ui-attachment/',
  'packages/experimental/computer-use-cua-driver-native/',
  'packages/experimental/translator/',
  'packages/extensions/tool-cordis/src/api-catalog.ts',
  'packages/llm/llm/src/index.ts',
  'packages/llm/llm/src/types.ts',
  'packages/mcp/mcp-client/',
  'packages/sdk/server/src/server.ts',
  'packages/subagent/subagent-dsh-sdk/src/index.ts',
  'packages/test-support/session-snapshot/src/harness.ts',
  'packages/web/tool-web/src/search.ts',
  'packages/web/web-api/',
  'packages/web/web-fetch-http/',
  'packages/web/web/src/types.ts',
] as const

/** Provider documentation that describes the external DeepSeek service rather
 * than presenting it as this product's identity. */
const PROVIDER_DOC_PATHS = [
  'docs/capability-seams.md',
  'docs/capability-seams.zh.md',
  'docs/config-catalog.md',
  'docs/config-catalog.zh.md',
  'docs/deepseek-llm-api-wire-extensions.md',
  'docs/deepseek-llm-api-wire-extensions.zh.md',
  'docs/development.md',
  'docs/development.zh.md',
  'docs/persistence-catalog.md',
  'docs/persistence-catalog.zh.md',
  'docs/subsystems/feedback.md',
  'docs/subsystems/feedback.zh.md',
  'docs/subsystems/llm-streaming.md',
  'docs/subsystems/llm-streaming.zh.md',
  'docs/subsystems/web.md',
  'docs/subsystems/web.zh.md',
  'docs/user/develop/practice/llm-adapter.md',
  'docs/user/develop/practice/llm-adapter.zh.md',
  'docs/user/guide/index.md',
  'docs/user/guide/index.zh.md',
  'docs/user/guide/mcp-memory.md',
  'docs/user/guide/mcp-memory.zh.md',
  'docs/user/guide/providers.md',
  'docs/user/guide/providers.zh.md',
  'docs/user/guide/python-sdk.md',
  'docs/user/guide/python-sdk.zh.md',
] as const

/** Frozen release history and upstream references; these documents describe what
 * shipped earlier and the upstream project, so their references stay. */
const UPSTREAM_REFERENCE_PATHS = [
  'vendor/README.md',
  'docs/architecture.md',
  'docs/architecture.zh.md',
  'docs/i18n/style-samples.md',
  'docs/i18n/style-samples.zh.md',
  'docs/i18n/translation-rules.md',
  'docs/i18n/translation-rules.zh.md',
  'docs/persistence-changes/',
  'docs/postmortem/0003-web-agent-gui-feedback-loop.md',
  'docs/postmortem/0003-web-agent-gui-feedback-loop.zh.md',
  'docs/upgrade-guide/',
] as const

/** Protocol names carried in requests to the DeepSeek API and in wire clients,
 * analogous to the `dsh.*` protocol vocabulary: `x-deepseek-harness-*` headers,
 * the wire-stable SDK runtime server name, and the ACP/Codex client names. They
 * move only with a wire compatibility change. */
const WIRE_PROTOCOL_LINES = [
  '[deepseek-harness] vendored modification',
  "engine: 'deepseek-harness'",
  "name: 'deepseek-harness'",
  "name: 'deepseek-harness-acp'",
  "name: 'deepseek-harness-acp-test-client'",
  '`deepseek-harness-sdk-runtime`',
  "name: 'deepseek-harness-sdk-runtime'",
  "name: 'deepseek-harness-subagent-acp'",
  'repository: deepseek-ai/deepseek-harness',
  'x-deepseek-harness-',
  ...WORKSPACE_DIRECTORY_LINES,
  ...CITATION_LINES,
] as const

/** Frozen release history; these documents describe what shipped earlier, so the
 * commands and environment names they record stay. */
const RELEASE_HISTORY_PATHS = [
  'docs/persistence-changes/',
  'docs/upgrade-guide/',
] as const

const RULES: readonly IdentityRule[] = [
  {
    id: 'deepseek-scope',
    label: 'DeepSeek package scope',
    pattern: /@deepseek-ai\/[a-z0-9][a-z0-9._~-]*/gu,
    allow: ALLOWED_DEEPSEEK_PACKAGES,
  },
  {
    id: 'deepseek-harness-token',
    label: 'lowercase deepseek-harness product token',
    pattern: /deepseek-harness/gu,
    allowLines: WIRE_PROTOCOL_LINES,
    allowPaths: [...T102_PYTHON_PATHS, ...GITHUB_REVIEW_PATHS, ...UPSTREAM_REFERENCE_PATHS],
  },
  {
    id: 'deepseek-brand-word',
    label: 'standalone DeepSeek brand word',
    pattern: /\bDeepSeek\b/gu,
    allowLines: [...CITATION_AUTHOR_LINES],
    allowPaths: [...T102_PROVIDER_PATHS, ...PROVIDER_REFERENCE_PATHS, ...PROVIDER_DOC_PATHS, ...UPSTREAM_REFERENCE_PATHS],
  },
  {
    id: 'deepseek-brand',
    label: 'DeepSeek Harness brand',
    pattern: /DeepSeek Harness/gu,
    allowPaths: RELEASE_HISTORY_PATHS,
  },
  {
    id: 'deepseek-api-key',
    label: 'legacy DEEPSEEK_API_KEY environment name',
    pattern: /DEEPSEEK_API_KEY/gu,
    allowPaths: RELEASE_HISTORY_PATHS,
  },
  {
    id: 'deepseek-base-url',
    label: 'legacy DEEPSEEK_BASE_URL environment name',
    pattern: /DEEPSEEK_BASE_URL/gu,
    allowPaths: RELEASE_HISTORY_PATHS,
  },
  {
    id: 'cli-package',
    label: 'legacy @forwork/dsh package name',
    pattern: /@forwork\/dsh\b/gu,
    allowPaths: RELEASE_HISTORY_PATHS,
  },
  {
    id: 'cli-bin',
    label: 'legacy dsh bin name',
    pattern: /\.name\(['"]dsh['"]\)/gu,
  },
  {
    id: 'cli-invocation',
    label: 'legacy dsh CLI invocation',
    pattern: new RegExp(
      String.raw`(?<![\w@/.-])dsh(?=\s+(?:--profile|--from-default-profile|--patch|` +
        String.raw`web\b|headless\b|sdk\b|acp\b|plugin\b|rescue\b|tui\b|desktop\b|help\b))`,
      'gu',
    ),
    allowPaths: RELEASE_HISTORY_PATHS,
  },
  {
    id: 'artifact-prefix',
    label: 'legacy dsh- artifact or DeepSeek Harness artifact prefix',
    pattern: /\bdsh-(?:desktop|web-frontend|desktop-host|plugin-market-generator)\b|\bdeepseek-harness-(?:\d|\$\{)/gu,
    allowPaths: RELEASE_HISTORY_PATHS,
  },
]

/** Generated catalogs that project shipped package and configuration names. */
const GENERATED_CATALOGS = [
  'docs/dependency-catalog.json',
  'docs/config-catalog.md',
  'docs/config-catalog.zh.md',
  'docs/tool-catalog.md',
  'docs/tool-catalog.zh.md',
  'docs/module-graph.md',
  'docs/module-graph.zh.md',
  'docs/capability-seams.md',
  'docs/capability-seams.zh.md',
  'docs/cli-help.md',
  'docs/cli-help.zh.md',
  'scripts/dependency-catalog/package-lock.json',
] as const

const SKIP_DIRECTORIES = new Set([
  '.agents', '.artifacts', '.dsh-build', '.generated', '.git', 'archived',
  'coverage', 'dist', 'lib', 'node_modules',
])

const SURFACE_EXTENSIONS = new Set(['.c', '.css', '.cts', '.html', '.js', '.json', '.mjs', '.mts', '.ts', '.tsx', '.yaml', '.yml'])

const PYTHON_EXTENSIONS = new Set(['.json', '.mjs', '.py', '.toml'])

const APP_EXTRA_EXTENSIONS = new Set([...SURFACE_EXTENSIONS, '.md', '.ps1', '.sh'])

const ROOT_FILES = ['README.md', 'README.zh.md'] as const

/** One violation of an identity rule. */
export interface IdentityViolation {
  rule: string
  label: string
  file: string
  line: number
  text: string
}

function normalized(relativePath: string): string {
  return relativePath.replaceAll(sep, '/')
}

function shouldScan(relativePath: string): boolean {
  if (relativePath.endsWith('package.json')) return true
  if (/\/src\//u.test(relativePath)) return true
  if (/\/locale\/[^/]+\.json$/u.test(relativePath)) return true
  return /\/cordis(?:\.patch)?\.ya?ml$/u.test(relativePath)
}

function isAppExtra(relativePath: string): boolean {
  if (relativePath.endsWith('.env.example')) return true
  if (!APP_EXTRA_EXTENSIONS.has(extname(relativePath))) return false
  return /^apps\/[^/]+\/config\//u.test(relativePath) || relativePath.startsWith('apps/desktop/scripts/')
}

function collectFiles(repoRoot: string, directory: string, found: Set<string>, admit: (relativePath: string) => boolean): void {
  let entries
  try {
    entries = readdirSync(directory, { withFileTypes: true })
  } catch {
    return
  }
  for (const entry of entries) {
    if (entry.isSymbolicLink()) continue
    const full = join(directory, entry.name)
    if (entry.isDirectory()) {
      if (SKIP_DIRECTORIES.has(entry.name)) continue
      collectFiles(repoRoot, full, found, admit)
      continue
    }
    const relativePath = normalized(relative(repoRoot, full))
    if (admit(relativePath)) found.add(full)
  }
}

/**
 * Collect every repository path the identity scan treats as a shipped surface.
 * @param repoRoot - repository root to scan.
 * @returns absolute file paths in no particular order.
 */
export function collectForworkIdentityFiles(repoRoot: string): string[] {
  const found = new Set<string>()
  for (const directory of ['packages', 'apps', 'vendor', 'native']) {
    collectFiles(repoRoot, join(repoRoot, directory), found, path => SURFACE_EXTENSIONS.has(extname(path)) && shouldScan(path))
  }
  collectFiles(repoRoot, join(repoRoot, 'apps'), found, isAppExtra)
  collectFiles(repoRoot, join(repoRoot, 'python'), found, path => PYTHON_EXTENSIONS.has(extname(path)))
  collectFiles(repoRoot, join(repoRoot, 'docs'), found, path => path.endsWith('.md'))
  for (const file of ROOT_FILES) found.add(join(repoRoot, file))
  found.add(join(repoRoot, 'website/.vitepress/config.ts'))
  for (const catalog of GENERATED_CATALOGS) found.add(join(repoRoot, catalog))
  return [...found]
}

function pathAllowed(relativePath: string, rule: IdentityRule): boolean {
  return rule.allowPaths?.some(allowed =>
    allowed.endsWith('/') ? relativePath.startsWith(allowed) : relativePath === allowed,
  ) ?? false
}

function matchesForbiddenToken(line: string, rule: IdentityRule): boolean {
  if (rule.allowLines?.some(literal => line.includes(literal))) return false
  for (const match of line.matchAll(rule.pattern)) {
    const allowed = rule.allow?.some(literal => match[0] === literal || match[0].startsWith(`${literal}-`)) ?? false
    if (!allowed) return true
  }
  return false
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function manifestBinViolations(relativePath: string, text: string): IdentityViolation[] {
  const manifest: unknown = JSON.parse(text)
  if (!isRecord(manifest) || !isRecord(manifest.bin) || !Object.hasOwn(manifest.bin, 'dsh')) return []
  return [{
    rule: 'cli-bin', label: 'legacy dsh bin name', file: relativePath, line: 1,
    text: JSON.stringify(manifest.bin),
  }]
}

/**
 * Scan shipped source, manifests, documentation, and generated catalogs for retained identity tokens.
 * @param repoRoot - repository root to scan.
 * @returns every violation in deterministic path and line order.
 */
export function scanForworkIdentity(repoRoot: string): IdentityViolation[] {
  const files = collectForworkIdentityFiles(repoRoot)
  const violations: IdentityViolation[] = []
  for (const file of files.sort()) {
    const relativePath = normalized(relative(repoRoot, file))
    const text = readFileSync(file, 'utf8')
    const lines = text.split(/\r?\n/u)
    for (const rule of RULES) {
      if (pathAllowed(relativePath, rule)) continue
      for (const [index, line] of lines.entries()) {
        if (!matchesForbiddenToken(line, rule)) continue
        violations.push({ rule: rule.id, label: rule.label, file: relativePath, line: index + 1, text: line.trim() })
      }
    }
    if (relativePath.endsWith('package.json')) violations.push(...manifestBinViolations(relativePath, text))
  }
  return violations
}

if (import.meta.main) {
  const violations = scanForworkIdentity(root)
  if (violations.length === 0) {
    console.log('verify-forwork-identity: shipped surfaces carry ForWork identity')
  } else {
    for (const violation of violations) {
      console.error(`${violation.file}:${violation.line}: ${violation.label}: ${violation.text}`)
    }
    console.error(`verify-forwork-identity: ${violations.length} violation(s)`)
    process.exitCode = 1
  }
}
