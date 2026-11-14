/** The Host companion registers no node-half behavior. */
import { describe, expect, it } from 'vitest'
import { apply } from '../src/index.ts'

describe('ui-office-editor host companion', () => {
  it('mounts nothing from the node face', () => {
    expect(apply).toBeTypeOf('function')
    apply()
  })
})
