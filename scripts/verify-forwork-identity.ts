/** Reject DeepSeek Harness identity tokens from shipped surfaces. */

import { readdirSync, readFileSync } from 'node:fs'
import { join, relative, resolve, sep } from 'node:path'

const root = resolve(import.meta.dirname, '..')

/** One forbidden identity token and its documented exceptions. */
interface IdentityRule {
  id: string
  label: string
  pattern: RegExp
  /** Lines containing one of these literals are known, reviewed exceptions. */
  allow?: readonly string[]
}

/** Known third-party scope that keeps the upstream DeepSeek package name. */
const ALLOWED_DEEPSEEK_PACKAGES = ['@deepseek-ai/libreoffice-kit'] as const

const RULES: readonly IdentityRule[] = [
  {
    id: 'deepseek-scope',
    label: 'DeepSeek package scope',
    pattern: /@deepseek-ai\/[a-z0-9][a-z0-9._~-]*/gu,
    allow: ALLOWED_DEEPSEEK_PACKAGES,
  },
  {
    id: 'deepseek-brand',
    label: 'DeepSeek Harness brand',
    pattern: /DeepSeek Harness/gu,
  },
  {
    id: 'deepseek-api-key',
    label: 'legacy DEEPSEEK_API_KEY environment name',
    pattern: /DEEPSEEK_API_KEY/gu,
  },
  {
    id: 'deepseek-base-url',
    label: 'legacy DEEPSEEK_BASE_URL environment name',
    pattern: /DEEPSEEK_BASE_URL/gu,
  },
  {
    id: 'cli-package',
    label: 'legacy @forwork/dsh package name',
    pattern: /@forwork\/dsh\b/gu,
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
  },
  {
    id: 'artifact-prefix',
    label: 'legacy dsh- artifact or DeepSeek Harness artifact prefix',
    pattern: /\bdsh-(?:desktop|web-frontend|desktop-host|plugin-market-generator)\b|\bdeepseek-harness-(?:\d|\$\{)/gu,
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
  '.artifacts', '.dsh-build', '.git', 'coverage', 'dist', 'lib', 'node_modules',
])

const SURFACE_EXTENSIONS = new Set(['.c', '.css', '.cts', '.html', '.js', '.json', '.mjs', '.mts', '.ts', '.tsx', '.yaml', '.yml'])

/** One violation of an identity rule. */
export interface IdentityViolation {
  rule: string
  label: string
  file: string
  line: number
  text: string
}

function shouldScan(relativePath: string): boolean {
  if (relativePath.endsWith('package.json')) return true
  if (/\/src\//u.test(relativePath.replaceAll(sep, '/'))) return true
  if (/\/locale\/[^/]+\.json$/u.test(relativePath.replaceAll(sep, '/'))) return true
  return /\/cordis(?:\.patch)?\.ya?ml$/u.test(relativePath.replaceAll(sep, '/'))
}

function collectSurfaceFiles(directory: string, found: string[]): void {
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    if (entry.isSymbolicLink()) continue
    const full = join(directory, entry.name)
    if (entry.isDirectory()) {
      if (SKIP_DIRECTORIES.has(entry.name)) continue
      collectSurfaceFiles(full, found)
      continue
    }
    const extension = entry.name.slice(entry.name.lastIndexOf('.'))
    if (!SURFACE_EXTENSIONS.has(extension)) continue
    const relativePath = relative(root, full)
    if (shouldScan(relativePath)) found.push(full)
  }
}

function matchesForbiddenToken(line: string, rule: IdentityRule): boolean {
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
 * Scan shipped source, manifests, and generated catalogs for retained identity tokens.
 * @param repoRoot - repository root to scan.
 * @returns every violation in deterministic path and line order.
 */
export function scanForworkIdentity(repoRoot: string): IdentityViolation[] {
  const files: string[] = []
  for (const directory of ['packages', 'apps', 'vendor', 'native']) collectSurfaceFiles(join(repoRoot, directory), files)
  for (const catalog of GENERATED_CATALOGS) files.push(join(repoRoot, catalog))
  const violations: IdentityViolation[] = []
  for (const file of files.sort()) {
    const relativePath = relative(repoRoot, file).replaceAll(sep, '/')
    const text = readFileSync(file, 'utf8')
    const lines = text.split(/\r?\n/u)
    for (const rule of RULES) {
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
