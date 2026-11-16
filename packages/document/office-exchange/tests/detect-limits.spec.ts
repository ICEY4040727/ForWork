/** Detection limits over a mocked unzip pass: expansion bound and non-Error reader failures. */
import { expect, it, vi } from 'vitest'
import { detectUnsupportedFeatures } from '../src/detect.ts'
import { OfficeExchangeError } from '../src/errors.ts'

const COMPRESSED = new Uint8Array([0x50, 0x4b])

const mocks = vi.hoisted(() => ({ unzip: vi.fn() }))

vi.mock('fflate', async (importOriginal) => {
  const actual = await importOriginal<typeof import('fflate')>()
  return {
    ...actual,
    unzipSync: (...args: Parameters<typeof actual.unzipSync>) => mocks.unzip(...args) as ReturnType<typeof actual.unzipSync>,
  }
})

function failureOf(run: () => void): OfficeExchangeError {
  try {
    run()
  } catch (error) {
    if (!(error instanceof OfficeExchangeError)) throw error
    return error
  }
  throw new Error('expected the detection to throw')
}

it('refuses a package whose entries expand beyond the limit', () => {
  mocks.unzip.mockImplementation((_bytes: Uint8Array, options?: {
    filter?: (file: { name: string; size: number; originalSize: number; compression: number }) => boolean
  }) => {
    options?.filter?.({ name: 'word/document.xml', size: 0, originalSize: 150 * 1024 * 1024, compression: 0 })
    options?.filter?.({ name: 'word/media/image.png', size: 0, originalSize: 100 * 1024 * 1024, compression: 0 })
    return {}
  })

  const error = failureOf(() => detectUnsupportedFeatures(COMPRESSED))

  expect(error.code).toBe('too-large')
  expect(error.message).toContain('expands beyond')
})

it('reports a non-Error unzip failure as text', () => {
  mocks.unzip.mockImplementation(() => { throw 'zip exploded' })

  const error = failureOf(() => detectUnsupportedFeatures(COMPRESSED))

  expect(error.code).toBe('conversion-failed')
  expect(error.message).toContain('zip exploded')
})
