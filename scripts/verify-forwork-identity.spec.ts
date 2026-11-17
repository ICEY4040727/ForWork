/** Rejection coverage for the ForWork identity scan. */

import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { describe, expect, it, onTestFinished } from 'vitest'
import { collectForworkIdentityFiles, scanForworkIdentity } from './verify-forwork-identity.ts'

/** Files the scanner reads unconditionally; the real gate expects the generators to have written them. */
const REQUIRED_FILES = [
  'README.md',
  'README.zh.md',
  'website/.vitepress/config.ts',
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

function write(root: string, relativePath: string, text: string): void {
  const full = join(root, relativePath)
  mkdirSync(dirname(full), { recursive: true })
  writeFileSync(full, text)
}

function fixture(): string {
  const root = mkdtempSync(join(tmpdir(), 'forwork-identity-'))
  onTestFinished(() => rmSync(root, { recursive: true, force: true }))
  for (const file of REQUIRED_FILES) write(root, file, '')
  return root
}

describe('scanForworkIdentity rejections', () => {
  it('rejects the company name in a manifest (CV-04)', () => {
    const root = fixture()
    write(root, 'packages/probe/package.json', JSON.stringify({ name: 'Hangzhou DeepSeek Artificial Intelligence Co., Ltd.' }))
    expect(scanForworkIdentity(root)).toContainEqual(
      expect.objectContaining({ rule: 'deepseek-brand-word', file: 'packages/probe/package.json', line: 1 }),
    )
  })

  it('rejects the lowercase deepseek-harness URL and product token in source', () => {
    const root = fixture()
    write(root, 'packages/probe/src/index.ts', "export const url = 'https://deepseek-harness.github.io/deepseek-harness/'\n")
    expect(scanForworkIdentity(root)).toContainEqual(
      expect.objectContaining({ rule: 'deepseek-harness-token', file: 'packages/probe/src/index.ts', line: 1 }),
    )
  })

  it('rejects a standalone DeepSeek brand word in a README', () => {
    const root = fixture()
    write(root, 'README.md', '# ForWork\n\nForWork is developed by DeepSeek AI.\n')
    expect(scanForworkIdentity(root)).toContainEqual(
      expect.objectContaining({ rule: 'deepseek-brand-word', file: 'README.md', line: 3 }),
    )
  })

  it('rejects a Python pyproject distribution named deepseek-harness-sdk', () => {
    const root = fixture()
    write(root, 'python/probe/pyproject.toml', '[project]\nname = "deepseek-harness-sdk"\nversion = "0.0.1"\n')
    expect(scanForworkIdentity(root)).toContainEqual(
      expect.objectContaining({ rule: 'deepseek-harness-token', file: 'python/probe/pyproject.toml', line: 2 }),
    )
  })

  it('rejects the legacy DEEPSEEK environment names outside release history', () => {
    const root = fixture()
    write(root, 'packages/probe/src/config.ts', "export const keyEnv = 'DEEPSEEK_API_KEY'\n")
    expect(scanForworkIdentity(root)).toContainEqual(
      expect.objectContaining({ rule: 'deepseek-api-key', file: 'packages/probe/src/config.ts', line: 1 }),
    )
  })
})

describe('scanForworkIdentity allowances', () => {
  it('accepts the external LibreOffice kit and a third-party notice entry', () => {
    const root = fixture()
    write(root, 'packages/probe/src/notices.ts', [
      "export const kit = '@deepseek-ai/libreoffice-kit'",
      "export const engine = '@deepseek-ai/libreoffice-kit-wasm'",
      '',
    ].join('\n'))
    write(root, 'THIRD_PARTY_NOTICES.md', '## @deepseek-ai/libreoffice-kit\n\nMPL-2.0\n')
    expect(scanForworkIdentity(root)).toEqual([])
  })

  it('still rejects a different package in the upstream scope', () => {
    const root = fixture()
    write(root, 'packages/probe/src/dependency.ts', "export const name = '@deepseek-ai/unreviewed-kit'\n")
    expect(scanForworkIdentity(root)).toContainEqual(
      expect.objectContaining({ rule: 'deepseek-scope', file: 'packages/probe/src/dependency.ts' }),
    )
  })

  it('admits released upgrade history that records the old command names', () => {
    const root = fixture()
    write(root, 'docs/upgrade-guide/v0.0.1/old-cli/guide.md', '```sh\ndsh --profile web\n```\n')
    expect(scanForworkIdentity(root)).toEqual([])
  })
})

describe('collectForworkIdentityFiles', () => {
  it('covers the product surfaces FR-009 names', () => {
    const root = fixture()
    for (const file of [
      'README.md',
      'README.zh.md',
      'python/sdk/pyproject.toml',
      'docs/architecture.md',
      'website/.vitepress/config.ts',
      'apps/web/config/production.json',
      'apps/desktop/config/macos-packaging/production.env.example',
      'apps/desktop/scripts/desktop-publish.ts',
      'packages/probe/package.json',
      'packages/probe/src/index.ts',
    ]) write(root, file, '')
    const collected = collectForworkIdentityFiles(root).map(path => path.slice(root.length + 1).replaceAll('\\', '/'))
    expect(collected).toEqual(expect.arrayContaining([
      'README.md',
      'README.zh.md',
      'python/sdk/pyproject.toml',
      'docs/architecture.md',
      'website/.vitepress/config.ts',
      'apps/web/config/production.json',
      'apps/desktop/config/macos-packaging/production.env.example',
      'apps/desktop/scripts/desktop-publish.ts',
      'packages/probe/package.json',
      'packages/probe/src/index.ts',
    ]))
  })

  it('excludes archived and generated trees', () => {
    const root = fixture()
    for (const file of [
      '.agents/notes/archived/process/old.md',
      'docs/archived/old.md',
      'website/tests/wordmark.spec.ts',
      'website/.generated/page.md',
    ]) write(root, file, 'deepseek-harness')
    const collected = collectForworkIdentityFiles(root).map(path => path.slice(root.length + 1).replaceAll('\\', '/'))
    expect(collected).not.toContain('.agents/notes/archived/process/old.md')
    expect(collected).not.toContain('docs/archived/old.md')
    expect(collected).not.toContain('website/tests/wordmark.spec.ts')
    expect(collected).not.toContain('website/.generated/page.md')
  })
})
