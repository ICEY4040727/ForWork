/** Byte application of sorted, non-overlapping splice operations. */
import { strFromU8, strToU8 } from 'fflate'
import { describe, expect, it } from 'vitest'
import { applySplices } from '../src/source/apply-splices.ts'

function apply(source: string, operations: Parameters<typeof applySplices>[1]): string {
  return strFromU8(applySplices(strToU8(source), operations))
}

function replace(start: number, end: number, text: string) {
  return { kind: 'replace' as const, start, end, bytes: strToU8(text) }
}

function insert(at: number, text: string) {
  return { kind: 'insert' as const, at, bytes: strToU8(text) }
}

function remove(start: number, end: number) {
  return { kind: 'delete' as const, start, end }
}

function append(at: number, text: string) {
  return { kind: 'append' as const, at, bytes: strToU8(text) }
}

describe('applySplices', () => {
  it('returns a new equal array for no operations', () => {
    const source = strToU8('abcdef')
    const result = applySplices(source, [])
    expect(result).toEqual(source)
    expect(result).not.toBe(source)
  })

  it('replaces one interior range and copies both sides verbatim', () => {
    expect(apply('abcdef', [replace(2, 4, 'XYZ')])).toBe('abXYZef')
  })

  it('inserts at the start, middle, and end', () => {
    expect(apply('abcdef', [insert(0, '<')])).toBe('<abcdef')
    expect(apply('abcdef', [insert(3, '!')])).toBe('abc!def')
    expect(apply('abcdef', [insert(6, '>')])).toBe('abcdef>')
  })

  it('deletes one range without changing its neighbours', () => {
    expect(apply('abcdef', [remove(2, 4)])).toBe('abef')
  })

  it('applies multiple sorted operations, including touching boundaries', () => {
    expect(apply('abcdef', [
      insert(0, '['),
      replace(1, 3, 'XX'),
      insert(3, '|'),
      remove(4, 5),
      insert(6, ']'),
    ])).toBe('[aXX|df]')
  })

  it('treats an append as an insertion at its offset', () => {
    expect(apply('ab', [append(2, '<x/>')])).toBe('ab<x/>')
    expect(apply('ab', [append(0, '<x/>')])).toBe('<x/>ab')
  })

  it('allows an insertion at a replacement start and end', () => {
    expect(apply('abcd', [insert(1, 'L'), replace(1, 3, 'XY'), insert(3, 'R')])).toBe('aLXYRd')
  })

  it('keeps replacement lengths independent of the consumed range', () => {
    expect(apply('abcdef', [replace(1, 5, '')])).toBe('af')
    expect(apply('abcdef', [replace(1, 2, 'long replacement')])).toBe('along replacementcdef')
  })

  it('returns bytes when the operations cover the whole part', () => {
    expect(apply('abc', [replace(0, 3, 'xyz')])).toBe('xyz')
    expect(apply('abc', [remove(0, 3)])).toBe('')
  })

  it('rejects an operation range outside the part', () => {
    const source = strToU8('abc')
    expect(() => applySplices(source, [replace(1, 4, 'x')])).toThrow(RangeError)
    expect(() => applySplices(source, [remove(0, 1.5)])).toThrow(RangeError)
    expect(() => applySplices(source, [insert(4, 'x')])).toThrow(RangeError)
    expect(() => applySplices(source, [remove(-1, 1)])).toThrow(RangeError)
  })

  it('rejects unsorted operations', () => {
    const source = strToU8('abcdef')
    expect(() => applySplices(source, [remove(4, 5), remove(1, 2)])).toThrow(RangeError)
    expect(() => applySplices(source, [insert(4, 'x'), insert(1, 'y')])).toThrow(RangeError)
  })

  it('rejects overlapping operations', () => {
    const source = strToU8('abcdef')
    expect(() => applySplices(source, [replace(1, 4, 'x'), remove(3, 5)])).toThrow(RangeError)
    expect(() => applySplices(source, [remove(1, 4), insert(2, 'x')])).toThrow(RangeError)
  })

  it('rejects an inverted replacement range', () => {
    const source = strToU8('abcdef')
    expect(() => applySplices(source, [remove(4, 2)])).toThrow(RangeError)
  })
})
