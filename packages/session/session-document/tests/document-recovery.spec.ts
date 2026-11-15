import { describe, expect, it } from 'vitest'
import { recoveryDecision } from '../src/index.ts'

describe('recoveryDecision', () => {
  it('offers nothing when the document has no draft', () => {
    expect(recoveryDecision(null, 'v1')).toBe('none')
    expect(recoveryDecision(undefined, 'v1')).toBe('none')
  })

  it('offers a draft captured from the current file version', () => {
    expect(recoveryDecision({ contentVersion: 'v1' }, 'v1')).toBe('offer')
  })

  it('marks a draft from any other version stale', () => {
    expect(recoveryDecision({ contentVersion: 'v0' }, 'v1')).toBe('stale')
    expect(recoveryDecision({ contentVersion: 'v2' }, 'v1')).toBe('stale')
  })
})
