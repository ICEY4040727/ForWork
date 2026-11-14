import { describe, expect, it } from 'vitest'
import { DocumentId } from '@deepseek-ai/dsh-session'
import {
  createDocumentIdentity,
  markDocumentMissing,
  rebindDocument,
  resolveDocumentBinding,
  type DocumentIdentityBinding,
} from '../src/index.ts'

const documentId = DocumentId('doc-1')

describe('createDocumentIdentity', () => {
  it('binds a new document id to a resolving path', () => {
    const binding = createDocumentIdentity(documentId, 'C:/work/report.docx')

    expect(binding).toEqual({
      documentId,
      path: 'C:/work/report.docx',
      missing: false,
    })
    expect('reboundAt' in binding).toBe(false)
  })
})

describe('rebindDocument', () => {
  it('preserves the document id and records the new path and time', () => {
    const original = createDocumentIdentity(documentId, 'C:/work/report.docx')

    const rebound = rebindDocument(original, 'C:/work/renamed.docx', 1_000, true)

    expect(rebound).toEqual({
      documentId,
      path: 'C:/work/renamed.docx',
      missing: false,
      reboundAt: 1_000,
    })
    expect(rebound.documentId).toBe(original.documentId)
  })

  it('clears missing only when the caller observes the path existing', () => {
    const missing = markDocumentMissing(createDocumentIdentity(documentId, 'C:/work/report.docx'), 900)

    const rebound = rebindDocument(missing, 'C:/work/renamed.docx', 1_000, true)

    expect(rebound.missing).toBe(false)
  })

  it('keeps a rebind missing when the caller cannot observe the new path', () => {
    const missing = markDocumentMissing(createDocumentIdentity(documentId, 'C:/work/report.docx'), 900)

    const rebound = rebindDocument(missing, 'C:/work/renamed.docx', 1_000, false)

    expect(rebound).toEqual({
      documentId,
      path: 'C:/work/renamed.docx',
      missing: true,
      reboundAt: 1_000,
    })
  })

  it('does not mutate the source binding', () => {
    const original = createDocumentIdentity(documentId, 'C:/work/report.docx')
    const snapshot: DocumentIdentityBinding = { ...original }

    rebindDocument(original, 'C:/work/renamed.docx', 1_000, false)

    expect(original).toEqual(snapshot)
    expect(original.path).toBe('C:/work/report.docx')
    expect(original.missing).toBe(false)
  })
})

describe('markDocumentMissing', () => {
  it('marks the binding missing while preserving identity, path, and rebind time', () => {
    const rebound = rebindDocument(createDocumentIdentity(documentId, 'C:/work/report.docx'), 'C:/work/renamed.docx', 1_000, true)

    const missing = markDocumentMissing(rebound, 2_000)

    expect(missing).toEqual({
      documentId,
      path: 'C:/work/renamed.docx',
      missing: true,
      reboundAt: 1_000,
    })
    expect(missing.documentId).toBe(documentId)
  })
})

describe('resolveDocumentBinding', () => {
  it('finds the binding that carries the document id', () => {
    const first = createDocumentIdentity(DocumentId('doc-1'), 'C:/work/one.docx')
    const second = createDocumentIdentity(DocumentId('doc-2'), 'C:/work/two.docx')

    expect(resolveDocumentBinding([first, second], DocumentId('doc-2'))).toBe(second)
  })

  it('returns undefined when no binding carries the document id', () => {
    const bindings = [createDocumentIdentity(documentId, 'C:/work/report.docx')]

    expect(resolveDocumentBinding(bindings, DocumentId('doc-missing'))).toBeUndefined()
  })
})
