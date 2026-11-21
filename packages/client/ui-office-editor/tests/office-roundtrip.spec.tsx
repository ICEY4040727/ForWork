// @vitest-environment jsdom
/**
 * FR-014: preservation state survives the full read → wire → editor → save
 * path. A document mounts in the real Univer Docs runtime, receives one text
 * edit, and serializes through the unit's own `save()`; the source sidecar must
 * still be present in the saved snapshot so the next save can splice.
 */
import { afterEach, describe, expect, it } from 'vitest'
import type { OfficeDocumentSnapshot } from '@forwork/office-exchange/types'
import { mount } from '../src/client/editor/docs.ts'

/**
 * jsdom has no 2D canvas implementation. Univer's render engine only requires
 * a context object to construct; drawing calls on a headless surface are
 * received by no-op methods.
 */
/** Structural subset a headless 2D context needs; every other member is a no-op. */
interface CanvasContextStub extends Record<string, unknown> {
  canvas: HTMLCanvasElement
  measureText: (text: string) => TextMetrics
}

function canvasStub(): () => void {
  const original: unknown = Reflect.get(HTMLCanvasElement.prototype, 'getContext')
  const stub = function getContext(
    this: HTMLCanvasElement,
    contextId: string,
  ): CanvasRenderingContext2D | null {
    if (contextId !== '2d') return null
    const target: CanvasContextStub = {
      canvas: this,
      measureText: () => ({ width: 0 }) as TextMetrics,
    }
    const context = new Proxy(target, {
      get(value, property) {
        if (typeof property === 'string' && property in value) return value[property]
        return () => undefined
      },
      set(value, property, next) {
        if (typeof property === 'string') value[property] = next
        return true
      },
    })
    const partial: Partial<CanvasRenderingContext2D> = context
    return partial as CanvasRenderingContext2D
  }
  HTMLCanvasElement.prototype.getContext = stub as typeof HTMLCanvasElement.prototype.getContext
  return () => {
    Reflect.set(HTMLCanvasElement.prototype, 'getContext', original)
  }
}

/** The sidecar an anchored read attaches to `data.officeExchange.source`. */
const SIDECAR = {
  schemaVersion: 1,
  format: 'docx',
  packageDigest: 'a'.repeat(64),
  byteLength: 1024,
  anchorMode: 'per-unit',
  anchors: [
    {
      kind: 'paragraph',
      part: 'word/document.xml',
      key: 'para_1',
      start: 0,
      end: 10,
      digest: 'b'.repeat(64),
    },
  ],
  sourceParts: ['word/document.xml'],
}

const ORIGINAL_TEXT = 'Original paragraph'
const EDITED_TEXT = 'Edited paragraph'

/** One wire document snapshot carrying the sidecar through the editor. */
function documentSnapshot(): OfficeDocumentSnapshot {
  return {
    kind: 'document',
    data: {
      id: 'roundtrip-document',
      locale: 'enUS',
      title: '',
      body: {
        dataStream: `${ORIGINAL_TEXT}\r\n`,
        textRuns: [],
        paragraphs: [{ startIndex: ORIGINAL_TEXT.length, paragraphId: 'para_1' }],
      },
      documentStyle: {},
      officeExchange: { source: SIDECAR },
    },
  }
}

afterEach(() => {
  document.body.replaceChildren()
})

describe('office editor round trip', () => {
  it('keeps data.officeExchange.source through an edit and unit.save()', async () => {
    const restoreCanvas = canvasStub()
    const container = document.createElement('div')
    document.body.appendChild(container)
    const runtime = await mount({ container, kind: 'document', snapshot: documentSnapshot() })
    try {
      expect(runtime.apply([{ kind: 'paragraph', index: 0, text: EDITED_TEXT }]))
        .toEqual({ status: 'applied' })

      const saved = runtime.getSnapshot()
      expect(saved.kind).toBe('document')
      const data = saved.data as Record<string, unknown>
      expect(data['officeExchange']).toEqual({ source: SIDECAR })
      const body = data['body'] as Record<string, unknown>
      expect(body['dataStream']).toContain(EDITED_TEXT)
      expect(body['dataStream']).not.toContain(ORIGINAL_TEXT)
    } finally {
      runtime.dispose()
      restoreCanvas()
    }
  })
})
