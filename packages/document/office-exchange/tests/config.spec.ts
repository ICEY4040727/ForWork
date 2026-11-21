/** Configuration resolution: defaults, overrides, and loud validation of bounds. */
import { expect, it } from 'vitest'
import type { OfficeExchangeConfig } from '../src/config.ts'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG, resolveConfig } from '../src/config.ts'

const INVALID_BOUNDS: readonly Partial<OfficeExchangeConfig>[] = [
  { maxFileBytes: 0 },
  { maxFileBytes: -1 },
  { maxFileBytes: Number.NaN },
  { maxFileBytes: Number.POSITIVE_INFINITY },
  { maxFileBytes: 1.5 },
  { maxCells: 0 },
  { maxCells: -2_000 },
  { maxCells: Number.NaN },
  { maxCells: Number.POSITIVE_INFINITY },
  { maxCells: 2.25 },
  { maxConversionMs: 0 },
  { maxConversionMs: -1 },
  { maxConversionMs: Number.NaN },
  { maxConversionMs: Number.POSITIVE_INFINITY },
  { maxConversionMs: 0.5 },
  { commandTimeoutMs: 0 },
  { commandTimeoutMs: -1 },
  { commandTimeoutMs: Number.NaN },
  { commandTimeoutMs: Number.POSITIVE_INFINITY },
  { commandTimeoutMs: 0.5 },
  { maxSourceAnchors: 0 },
  { maxSourceAnchors: -1 },
  { maxSourceAnchors: Number.NaN },
  { maxSourceAnchors: Number.POSITIVE_INFINITY },
  { maxSourceAnchors: 1.5 },
  { maxSourceAnchorBytes: 0 },
  { maxSourceAnchorBytes: -1 },
  { maxSourceAnchorBytes: Number.NaN },
  { maxSourceAnchorBytes: Number.POSITIVE_INFINITY },
  { maxSourceAnchorBytes: 0.5 },
]

function messageOf(overrides: Partial<OfficeExchangeConfig>): string {
  try {
    resolveConfig(overrides)
  } catch (error) {
    if (error instanceof RangeError) return error.message
    throw error
  }
  throw new Error('expected resolveConfig to throw')
}

it('resolves the default bounds when nothing is overridden', () => {
  expect(resolveConfig({})).toEqual({
    maxFileBytes: 25 * 1024 * 1024,
    maxCells: 2_000_000,
    maxConversionMs: 60_000,
    commandTimeoutMs: 5_000,
    maxSourceAnchors: 100_000,
    maxSourceAnchorBytes: 8 * 1024 * 1024,
  })
})

it('merges each override over the defaults', () => {
  expect(resolveConfig({ maxCells: 10 })).toEqual({
    maxFileBytes: 25 * 1024 * 1024,
    maxCells: 10,
    maxConversionMs: 60_000,
    commandTimeoutMs: 5_000,
    maxSourceAnchors: 100_000,
    maxSourceAnchorBytes: 8 * 1024 * 1024,
  })
  expect(resolveConfig({ maxFileBytes: 1_024, maxConversionMs: 5 })).toEqual({
    maxFileBytes: 1_024,
    maxCells: 2_000_000,
    maxConversionMs: 5,
    commandTimeoutMs: 5_000,
    maxSourceAnchors: 100_000,
    maxSourceAnchorBytes: 8 * 1024 * 1024,
  })
})

it('accepts positive integer bounds', () => {
  expect(resolveConfig({
    maxFileBytes: 1,
    maxCells: 1,
    maxConversionMs: 1,
    commandTimeoutMs: 1,
    maxSourceAnchors: 1,
    maxSourceAnchorBytes: 1,
  })).toEqual({
    maxFileBytes: 1,
    maxCells: 1,
    maxConversionMs: 1,
    commandTimeoutMs: 1,
    maxSourceAnchors: 1,
    maxSourceAnchorBytes: 1,
  })
})

it('does not mutate the defaults', () => {
  resolveConfig({ maxFileBytes: 1, maxCells: 2, maxConversionMs: 3 })
  expect(DEFAULT_OFFICE_EXCHANGE_CONFIG).toEqual({
    maxFileBytes: 25 * 1024 * 1024,
    maxCells: 2_000_000,
    maxConversionMs: 60_000,
    commandTimeoutMs: 5_000,
    maxSourceAnchors: 100_000,
    maxSourceAnchorBytes: 8 * 1024 * 1024,
  })
})

it('rejects zero, negative, non-finite, and fractional bounds', () => {
  for (const overrides of INVALID_BOUNDS) {
    expect(() => resolveConfig(overrides)).toThrow(RangeError)
  }
})

it('names the invalid field and value in the message', () => {
  expect(messageOf({ maxCells: 0 })).toContain('maxCells')
  expect(messageOf({ maxCells: 0 })).toContain('0')
  expect(messageOf({ maxFileBytes: Number.NaN })).toContain('maxFileBytes')
  expect(messageOf({ maxFileBytes: Number.NaN })).toContain('NaN')
  expect(messageOf({ maxConversionMs: 1.5 })).toContain('maxConversionMs')
  expect(messageOf({ maxConversionMs: 1.5 })).toContain('1.5')
})
