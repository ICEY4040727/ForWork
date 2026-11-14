/**
 * The client reaches document content only through the `officeDocument`
 * Remote. Source inspection keeps the boundary honest: the hook imports only
 * the generated wire types from the Exchange package, never its server entry,
 * and names no Host service.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'

function sourceOf(relative: string): string {
  return readFileSync(new URL(relative, import.meta.url), 'utf8')
}

function importSpecifiers(source: string): string[] {
  return [...source.matchAll(/from '([^']+)'/gu)].map(match => match[1] ?? '')
}

describe('client runtime boundary', () => {
  it('keeps Univer out of the statically imported adapter graph', () => {
    for (const file of ['../src/client/adapter.ts', '../src/client/editor/load.ts']) {
      expect(sourceOf(file)).not.toMatch(/from '@univerjs/u)
    }
  })
})

describe('useOfficeDocument Remote boundary', () => {
  const source = sourceOf('../src/client/useOfficeDocument.ts')

  it('imports only the generated wire types from the Exchange package', () => {
    const exchangeImports = importSpecifiers(source).filter(specifier =>
      specifier.startsWith('@deepseek-ai/dsh-office-exchange'))
    expect(exchangeImports).toEqual(['@deepseek-ai/dsh-office-exchange/types'])
  })

  it('names no Host service', () => {
    expect(source).not.toMatch(/\bofficeExchange\b/u)
    expect(source).not.toMatch(/\bctx\b/u)
  })

  it('reads and saves through the injected Remote face', () => {
    expect(source).toContain('remote.read(')
    expect(source).toContain('remote.save(')
  })
})
