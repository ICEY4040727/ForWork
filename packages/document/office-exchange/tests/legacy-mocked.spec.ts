/**
 * Legacy conversion failure paths with a mocked LibreOffice engine: a missing,
 * empty, or unreadable output, engine unavailability, cancellation, and the
 * configured conversion deadline.
 */
import { mkdir, mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { ConversionError } from '@deepseek-ai/libreoffice-kit'
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { DEFAULT_OFFICE_EXCHANGE_CONFIG, resolveConfig } from '../src/config.ts'
import { OfficeExchangeError } from '../src/errors.ts'
import { convertLegacyFile } from '../src/provider/legacy.ts'

const config = DEFAULT_OFFICE_EXCHANGE_CONFIG

const mocks = vi.hoisted(() => ({
  createConverter: vi.fn(),
  statOverride: undefined as undefined | ((path: string) => Promise<Awaited<ReturnType<typeof stat>>>),
}))

vi.mock('@deepseek-ai/libreoffice-kit', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@deepseek-ai/libreoffice-kit')>()
  return { ...actual, createConverter: mocks.createConverter }
})

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return {
    ...actual,
    stat: (path: string) => mocks.statOverride === undefined ? actual.stat(path) : mocks.statOverride(path),
  }
})

const roots: string[] = []

beforeEach(() => {
  mocks.createConverter.mockReset()
  mocks.statOverride = undefined
})

afterEach(async () => {
  for (const root of roots.splice(0)) await rm(root, { recursive: true, force: true })
})

async function workspace(): Promise<{ input: string; output: string }> {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-legacy-mocked-'))
  roots.push(root)
  const input = join(root, 'report.xls')
  await writeFile(input, 'legacy bytes')
  return { input, output: join(root, 'converted') }
}

function engine(convert: (outputPath: string, signal: AbortSignal) => Promise<void> | void): {
  convert: ReturnType<typeof vi.fn>
  dispose: ReturnType<typeof vi.fn>
} {
  return {
    convert: vi.fn(async (request: { outputPath: string }, signal: AbortSignal) => {
      await convert(request.outputPath, signal)
    }),
    dispose: vi.fn(async () => {}),
  }
}

async function failureOf(operation: Promise<unknown>): Promise<OfficeExchangeError> {
  try {
    await operation
  } catch (error) {
    if (!(error instanceof OfficeExchangeError)) throw error
    return error
  }
  throw new Error('expected the conversion to reject')
}

it('reports a conversion that produced no output file', async () => {
  mocks.createConverter.mockResolvedValue(engine(() => {}))
  const { input, output } = await workspace()

  const error = await failureOf(convertLegacyFile(input, output, { config }))

  expect(error.code).toBe('conversion-failed')
  expect(error.message).toContain('missing or empty')
})

it('reports a conversion that produced an empty file', async () => {
  mocks.createConverter.mockResolvedValue(engine(async (outputPath) => { await writeFile(outputPath, '') }))
  const { input, output } = await workspace()

  const error = await failureOf(convertLegacyFile(input, output, { config }))

  expect(error.code).toBe('conversion-failed')
  expect(error.message).toContain('missing or empty')
})

it('reports an output path that is not a regular file', async () => {
  mocks.createConverter.mockResolvedValue(engine(async (outputPath) => { await mkdir(outputPath) }))
  const { input, output } = await workspace()

  const error = await failureOf(convertLegacyFile(input, output, { config }))

  expect(error.code).toBe('conversion-failed')
  expect(error.message).toContain('missing or empty')
})

it('reports an output size check that fails for another reason', async () => {
  mocks.createConverter.mockResolvedValue(engine(async (outputPath) => { await writeFile(outputPath, 'pdf') }))
  mocks.statOverride = async () => { throw Object.assign(new Error('size check failed'), { code: 'EACCES' }) }
  const { input, output } = await workspace()

  const error = await failureOf(convertLegacyFile(input, output, { config }))

  expect(error.code).toBe('conversion-failed')
  expect(error.message).toContain('size check failed')
})

it('keeps an Exchange refusal raised by the engine', async () => {
  mocks.createConverter.mockResolvedValue(engine(() => {
    throw new OfficeExchangeError('conversion-failed', 'engine refused the source')
  }))
  const { input, output } = await workspace()

  const error = await failureOf(convertLegacyFile(input, output, { config }))

  expect(error.message).toBe('engine refused the source')
})

it('reports an unavailable engine', async () => {
  mocks.createConverter.mockRejectedValue(new ConversionError('unavailable', 'no engine'))
  const { input, output } = await workspace()

  const error = await failureOf(convertLegacyFile(input, output, { config }))

  expect(error.code).toBe('conversion-failed')
  expect(error.message).toBe('LibreOffice is not installed')
})

it('reports an engine Error with its message', async () => {
  mocks.createConverter.mockRejectedValue(new Error('engine died'))
  const { input, output } = await workspace()

  const error = await failureOf(convertLegacyFile(input, output, { config }))

  expect(error.message).toContain('engine died')
})

it('reports a non-Error engine rejection as text', async () => {
  mocks.createConverter.mockRejectedValue('engine exploded')
  const { input, output } = await workspace()

  const error = await failureOf(convertLegacyFile(input, output, { config }))

  expect(error.message).toContain('engine exploded')
})

it('honors an abort that arrives after the engine starts', async () => {
  const controller = new AbortController()
  const reason = new Error('cancelled after start')
  mocks.createConverter.mockImplementation(async () => {
    controller.abort(reason)
    return engine(() => {})
  })
  const { input, output } = await workspace()

  await expect(convertLegacyFile(input, output, { config, signal: controller.signal })).rejects.toBe(reason)
})

it('honors an abort raised by the engine conversion', async () => {
  const controller = new AbortController()
  const reason = new Error('cancelled during conversion')
  mocks.createConverter.mockResolvedValue(engine(() => {
    controller.abort(reason)
    throw reason
  }))
  const { input, output } = await workspace()

  await expect(convertLegacyFile(input, output, { config, signal: controller.signal })).rejects.toBe(reason)
})

it('aborts a conversion that outlives maxConversionMs and leaves the source unchanged', async () => {
  mocks.createConverter.mockResolvedValue(engine(async (_outputPath, signal) => {
    await new Promise((_resolve, reject) => {
      signal.addEventListener('abort', () => {
        reject(new Error('engine observed the conversion deadline', { cause: signal.reason }))
      }, { once: true })
    })
  }))
  const { input, output } = await workspace()
  const sourceBefore = await readFile(input)
  const bound = resolveConfig({ maxConversionMs: 20 })

  const error = await failureOf(convertLegacyFile(input, output, { config: bound }))

  expect(error.code).toBe('too-large')
  expect(error.message).toContain('maxConversionMs')
  expect(error.message).toContain('20')
  expect(await readFile(input)).toEqual(sourceBefore)
  expect(mocks.createConverter).toHaveBeenCalledWith({ timeoutMs: 20 })
})

it('converts a small file under the default config', async () => {
  mocks.createConverter.mockResolvedValue(engine(async (outputPath) => { await writeFile(outputPath, 'converted') }))
  const { input, output } = await workspace()

  const converted = await convertLegacyFile(input, output, { config })

  expect(converted).toBe(join(output, 'report.xlsx'))
  expect(await readFile(converted, 'utf8')).toBe('converted')
  expect(mocks.createConverter).toHaveBeenCalledWith({ timeoutMs: config.maxConversionMs })
})
