import { describe, expect, it } from 'vitest'
import { DocumentId, type DocumentEditOperation } from '@deepseek-ai/dsh-session'
import {
  applyChange,
  OfficeDocumentError,
  undoChange,
  type ApplyChangeInput,
  type OfficeDocumentErrorCode,
} from '../src/index.ts'

const documentId = DocumentId('doc-1')

const operation: DocumentEditOperation = [
  { kind: 'cell', sheetId: 'sheet-1', row: 2, column: 1, value: 42 },
]

const inverse: DocumentEditOperation = [
  { kind: 'cell', sheetId: 'sheet-1', row: 2, column: 1, value: 'old' },
]

function request(overrides: Partial<ApplyChangeInput> = {}): ApplyChangeInput {
  return {
    documentId,
    actor: 'agent',
    currentVersion: 3,
    fromVersion: 3,
    operation,
    inverse,
    at: 1_000,
    ...overrides,
  }
}

function errorCode(run: () => void): OfficeDocumentErrorCode | undefined {
  try {
    run()
  } catch (error) {
    if (error instanceof OfficeDocumentError) return error.code
    throw error
  }
  return undefined
}

describe('applyChange', () => {
  it('accepts a change that matches the live version and advances it by one', () => {
    const { record, version } = applyChange(request())

    expect(version).toBe(4)
    expect(record).toEqual({
      changeId: 'doc-1:4',
      documentId,
      actor: 'agent',
      fromVersion: 3,
      toVersion: 4,
      operation,
      inverse,
      at: 1_000,
    })
  })

  it('records the actor and both directions of the edit', () => {
    const { record } = applyChange(request({ actor: 'user' }))

    expect(record.actor).toBe('user')
    expect(record.operation).toBe(operation)
    expect(record.inverse).toBe(inverse)
    expect(record.toVersion).toBe(record.fromVersion + 1)
  })

  it('derives a deterministic change id from the document id and target version', () => {
    const first = applyChange(request({ currentVersion: 0, fromVersion: 0 }))
    const second = applyChange(request({ currentVersion: 0, fromVersion: 0 }))

    expect(first.record.changeId).toBe('doc-1:1')
    expect(second.record.changeId).toBe(first.record.changeId)
  })

  it('keeps an explicit change id', () => {
    const { record } = applyChange(request({ changeId: 'custom-change' }))

    expect(record.changeId).toBe('custom-change')
  })

  it('rejects a change whose fromVersion is not the live version', () => {
    const run = (): void => {
      applyChange(request({ currentVersion: 4, fromVersion: 3 }))
    }

    expect(errorCode(run)).toBe('stale-version')
  })

  it('rejects an empty operation', () => {
    const run = (): void => {
      applyChange(request({ operation: [] }))
    }

    expect(errorCode(run)).toBe('invalid-operation')
  })

  it('rejects an empty inverse', () => {
    const run = (): void => {
      applyChange(request({ inverse: [] }))
    }

    expect(errorCode(run)).toBe('invalid-operation')
  })

  it('rejects a version that is not a non-negative safe integer', () => {
    expect(errorCode(() => { applyChange(request({ currentVersion: 1.5, fromVersion: 1.5 })) })).toBe('invalid-transition')
    expect(errorCode(() => { applyChange(request({ currentVersion: -1, fromVersion: -1 })) })).toBe('invalid-transition')
  })

  it('rejects a change that cannot advance to a safe integer version', () => {
    const run = (): void => {
      applyChange(request({ currentVersion: Number.MAX_SAFE_INTEGER, fromVersion: Number.MAX_SAFE_INTEGER }))
    }

    expect(errorCode(run)).toBe('invalid-transition')
  })

  it('does not mutate its input', () => {
    const input = request()
    const operationSnapshot = [...operation]

    applyChange(input)

    expect(input.fromVersion).toBe(3)
    expect(input.operation).toBe(operation)
    expect([...input.operation]).toEqual(operationSnapshot)
    expect(input.inverse).toBe(inverse)
  })
})

describe('undoChange', () => {
  it('returns the inverse steps and the version before the change', () => {
    const { record } = applyChange(request())

    expect(undoChange(record)).toEqual({ operation: inverse, version: 3 })
  })
})
