/**
 * Assert that the committed third-party notice list preserves the MIT notice
 * for the harness base and the Apache-2.0 notice for the Univer editor core
 * (FR-012). Runs without an install: it reads `THIRD_PARTY_NOTICES.md`, the
 * root `LICENSE`, and the workspace manifests only.
 */

import { existsSync, globSync, readFileSync } from 'node:fs'
import { resolve, sep } from 'node:path'

const ROOT = resolve(import.meta.dirname, '..')
const NOTICES = 'THIRD_PARTY_NOTICES.md'
const LICENSE = 'LICENSE'
const VENDORED_HEADING = '## Vendored source (`vendor/`)'
const RUNTIME_HEADING = '## Runtime npm dependencies'
const UNIVER_SCOPE = '@univerjs/'
const UNIVER_CORE = '@univerjs/core'
const UNIVER_LICENSE = 'Apache-2.0'
const MANIFEST_KINDS = ['dependencies', 'devDependencies', 'optionalDependencies', 'peerDependencies'] as const

/** One notice table entry: the package identity and the license column verbatim. */
interface NoticeEntry {
  name: string
  license: string
}

/** Result of checking the committed notice list against the workspace. */
export interface LicenseNoticeReport {
  /** Number of vendored harness packages whose notice entry was checked. */
  vendoredPackageCount: number
  /** Number of Univer packages whose notice entry was checked. */
  univerPackageCount: number
  /** Repository-relative diagnostics for missing or mismatched notices. */
  failures: string[]
}

function readText(root: string, file: string): string | undefined {
  const path = resolve(root, file)
  // Normalize checkout line endings so section and table parsing match on every platform.
  return existsSync(path) ? readFileSync(path, 'utf8').replaceAll('\r\n', '\n') : undefined
}

function readManifest(root: string, file: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(readFileSync(resolve(root, file), 'utf8'))
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error(`verify-license-notices: ${file} must contain a JSON object.`)
  }
  return parsed as Record<string, unknown>
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((entry: unknown) => typeof entry === 'string')
}

function isStringRecord(value: unknown): value is Record<string, string> {
  return typeof value === 'object'
    && value !== null
    && !Array.isArray(value)
    && Object.values(value).every((entry: unknown) => typeof entry === 'string')
}

function workspaceManifestPaths(root: string): string[] {
  const workspaces = readManifest(root, 'package.json').workspaces
  if (!isStringArray(workspaces)) {
    throw new Error('verify-license-notices: package.json workspaces must be a string array.')
  }
  const files = new Set<string>()
  for (const pattern of workspaces) {
    for (const file of globSync(`${pattern}/package.json`, { cwd: root })) {
      files.add(file.split(sep).join('/'))
    }
  }
  return [...files].sort()
}

/**
 * Collect every non-workspace `@univerjs/` dependency the workspace declares.
 * @param root - absolute repository root.
 * @returns sorted package identities from every manifest dependency section.
 */
export function declaredUniverPackages(root: string): string[] {
  const names = new Set<string>()
  for (const file of workspaceManifestPaths(root)) {
    const manifest = readManifest(root, file)
    for (const kind of MANIFEST_KINDS) {
      const section = manifest[kind]
      if (!isStringRecord(section)) continue
      for (const [name, range] of Object.entries(section)) {
        if (name.startsWith(UNIVER_SCOPE) && !range.startsWith('workspace:')) names.add(name)
      }
    }
  }
  return [...names].sort()
}

/** The lines of one level-two section, stopping before the next one. */
function sectionBody(notices: string, heading: string): string | undefined {
  const start = notices.indexOf(`${heading}\n`)
  if (start === -1) return undefined
  const body = notices.slice(start + heading.length + 1)
  const next = body.indexOf('\n## ')
  return (next === -1 ? body : body.slice(0, next)).trim()
}

/** Parse every package/license table row in a section, skipping headers and separators. */
function noticeEntries(section: string): NoticeEntry[] {
  const entries: NoticeEntry[] = []
  for (const line of section.split('\n')) {
    if (!line.startsWith('|')) continue
    const cells = line.slice(1, -1).split('|').map(cell => cell.trim())
    const first = cells[0] ?? ''
    const name = /\[`([^`]+)`\]/.exec(first)?.[1] ?? /^`([^`]+)`$/.exec(first)?.[1]
    const license = cells.at(-1) ?? ''
    if (name === undefined || license === 'License' || license === '') continue
    entries.push({ name, license })
  }
  return entries
}

/**
 * Check the committed notice list for the MIT harness base and the Apache-2.0
 * Univer editor core.
 * @param root - absolute repository root containing `THIRD_PARTY_NOTICES.md`.
 * @returns the checked entry counts and every missing or mismatched notice.
 */
export function inspectLicenseNotices(root: string): LicenseNoticeReport {
  const failures: string[] = []
  const notices = readText(root, NOTICES)
  if (notices === undefined) {
    return {
      vendoredPackageCount: 0,
      univerPackageCount: 0,
      failures: [`${NOTICES}: missing; the packaged notice list is not committed.`],
    }
  }

  const license = readText(root, LICENSE)
  if (license === undefined) {
    failures.push(`${LICENSE}: missing; the MIT harness base license text is not committed.`)
  } else if (!license.startsWith('MIT License') || !license.includes('Permission is hereby granted, free of charge')) {
    failures.push(`${LICENSE}: does not contain the MIT harness base license text.`)
  }
  if (!notices.includes('ForWork is licensed under [MIT](LICENSE).')) {
    failures.push(`${NOTICES}: does not state the MIT license for the harness base.`)
  }

  let vendoredPackageCount = 0
  const vendored = sectionBody(notices, VENDORED_HEADING)
  if (vendored === undefined) {
    failures.push(`${NOTICES}: no "${VENDORED_HEADING}" section for the MIT harness base.`)
  } else {
    const entries = noticeEntries(vendored)
    vendoredPackageCount = entries.length
    if (entries.length === 0) {
      failures.push(`${NOTICES}: the vendored harness section lists no MIT notice entries.`)
    }
    for (const entry of entries) {
      if (entry.license !== 'MIT') {
        failures.push(`${NOTICES}: vendored harness package ${entry.name} declares ${entry.license}; expected MIT.`)
      }
    }
    if (!entries.some(entry => entry.name === '@forwork/cordis')) {
      failures.push(`${NOTICES}: the vendored harness section has no @forwork/cordis MIT notice entry.`)
    }
  }

  const univerPackages = declaredUniverPackages(root)
  if (univerPackages.length === 0) {
    failures.push('packages/*/*/package.json: no @univerjs/ dependency is declared; the editor core declaration is missing.')
  } else if (!univerPackages.includes(UNIVER_CORE)) {
    failures.push(`packages/*/*/package.json: the Univer editor core ${UNIVER_CORE} is no longer declared.`)
  }

  const runtime = sectionBody(notices, RUNTIME_HEADING)
  if (runtime === undefined) {
    failures.push(`${NOTICES}: no "${RUNTIME_HEADING}" section for the Apache-2.0 editor core notice.`)
  } else {
    const runtimeEntries = new Map(noticeEntries(runtime).map(entry => [entry.name, entry.license]))
    for (const name of univerPackages) {
      const declared = runtimeEntries.get(name)
      if (declared === undefined) {
        failures.push(`${NOTICES}: the runtime notice list has no entry for ${name}; the Apache-2.0 editor core notice is missing.`)
      } else if (declared !== UNIVER_LICENSE) {
        failures.push(`${NOTICES}: runtime ${name} is listed as ${declared}; expected ${UNIVER_LICENSE}.`)
      }
    }
  }

  return { vendoredPackageCount, univerPackageCount: univerPackages.length, failures }
}

if (process.argv[1] !== undefined && import.meta.filename === resolve(process.argv[1])) {
  const report = inspectLicenseNotices(ROOT)
  if (report.failures.length > 0) {
    process.stderr.write('verify-license-notices: MIT harness base or Apache-2.0 editor core notice missing:\n')
    for (const failure of report.failures) process.stderr.write(`  ${failure}\n`)
    process.exitCode = 1
  } else {
    process.stdout.write(
      `verify-license-notices: ${String(report.vendoredPackageCount)} MIT harness base and `
      + `${String(report.univerPackageCount)} Apache-2.0 Univer notice entries checked in ${NOTICES}.\n`,
    )
  }
}
