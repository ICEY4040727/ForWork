/** The open-document prompt section leads with an id the document tools resolve. */
import { copyFile, mkdir } from 'node:fs/promises'
import { join } from 'node:path'
import { DocumentId } from '@forwork/session'
import { renderOpenDocuments } from '@forwork/session-document'
import type { DocumentProjectionState } from '@forwork/session-document'
import type { JsonValue } from '@forwork/util-values'
import { describe, expect, it, vi } from 'vitest'
import { documentTarget } from '../src/documents.ts'
import { createHarness, disposeHarness, documentIdFor } from './harness.ts'

const SIGNAL = new AbortController().signal

/** A JSON object view of one tool result value. */
function objectOf(value: JsonValue | undefined): Record<string, JsonValue> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) throw new Error('expected a JSON object')
  return value
}

/** The addressable id of one rendered section line. */
function idOf(line: string): string {
  const id = /^\d+\. (.+) \(/.exec(line)?.[1]
  if (id === undefined) throw new Error(`expected an addressable id in "${line}"`)
  return id
}

describe('open-document prompt ids', () => {
  it('resolves rendered ids verbatim for same-named files in different directories', async () => {
    const h = await createHarness()
    try {
      const name = await h.copy('roundtrip.xlsx')
      await mkdir(join(h.root, 'archive'), { recursive: true })
      await copyFile(join(h.root, name), join(h.root, 'archive', name))
      const paths = [name, `archive/${name}`] as const
      const state: DocumentProjectionState = {
        order: paths.map(path => documentIdFor(h, path)),
        documents: Object.fromEntries(paths.map((path) => {
          const documentId = DocumentId(documentIdFor(h, path))
          return [documentId, {
            documentId,
            path,
            format: 'xlsx',
            contentVersion: 0,
            status: 'open',
            unsupported: [],
          }]
        })),
      }

      const rendered = renderOpenDocuments(state).split('\n').slice(1).map(idOf)

      expect(rendered).toHaveLength(2)
      expect(new Set(rendered).size).toBe(2)
      expect(rendered).toEqual(paths.map(path => documentIdFor(h, path)))

      const read = await h.server.read(h.scope, paths[0], SIGNAL)
      vi.spyOn(h.server, 'dispatchCommand').mockResolvedValue({
        status: 'read',
        snapshot: read.snapshot,
        contentVersion: 0,
      })

      for (const [index, documentId] of rendered.entries()) {
        const path = paths[index]
        if (path === undefined) throw new Error('expected one rendered path per entry')
        const target = documentTarget(documentId, h.agent.session.id, h.root)
        expect(target.path).toBe(path)
        expect(target.absolutePath).toBe(join(h.root, path))

        const result = objectOf((await h.execute('document_read', { documentId }, SIGNAL)).value)
        expect(result['status']).toBe('ok')
        expect(result['documentId']).toBe(documentId)
      }
    } finally {
      await disposeHarness(h)
    }
  })
})
