/**
 * Byte application of a splice plan's span operations.
 *
 * The applicator is independent of the provider that planned the operations:
 * a caller supplies one part's uncompressed source bytes and the sorted,
 * non-overlapping operations for that part, and receives a complete new byte
 * array. Bytes outside every operation are copied exactly, so an edited unit
 * never disturbs its neighbours.
 *
 * Operations must be sorted by offset and must not overlap. Touching ranges are
 * allowed: an insertion at a replacement's start or end is applied on the
 * corresponding side. Violating that contract is a caller defect, not a user
 * refusal, so it raises {@link RangeError} instead of producing partial output.
 * @module @forwork/office-exchange/source/apply-splices
 */

/** Replace an existing byte range with rendered unit bytes. */
export interface ApplyReplace {
  /** Operation discriminant. */
  readonly kind: 'replace'
  /** Inclusive source byte offset. */
  readonly start: number
  /** Exclusive source byte offset. */
  readonly end: number
  /** Replacement bytes. */
  readonly bytes: Uint8Array
}

/** Insert rendered unit bytes at a source offset. */
export interface ApplyInsert {
  /** Operation discriminant. */
  readonly kind: 'insert'
  /** Source byte offset before which the bytes are inserted. */
  readonly at: number
  /** Inserted bytes. */
  readonly bytes: Uint8Array
}

/** Remove a source byte range. */
export interface ApplyDelete {
  /** Operation discriminant. */
  readonly kind: 'delete'
  /** Inclusive source byte offset. */
  readonly start: number
  /** Exclusive source byte offset. */
  readonly end: number
}

/** Append auxiliary element bytes at a source offset. */
export interface ApplyAppend {
  /** Operation discriminant. */
  readonly kind: 'append'
  /** Source byte offset before which the bytes are inserted. */
  readonly at: number
  /** Appended bytes. */
  readonly bytes: Uint8Array
}

/** One byte-range operation accepted by {@link applySplices}. */
export type ApplySpliceOperation = ApplyReplace | ApplyInsert | ApplyDelete | ApplyAppend

/**
 * Apply one part's span operations to its source bytes.
 *
 * An append is an insertion at its offset: both add bytes without consuming a
 * source range. The remaining source bytes before, between, and after the
 * operations are copied verbatim.
 * @param partBytes - Uncompressed source bytes of the part.
 * @param operations - Sorted, non-overlapping operations for this part only.
 * @returns A new byte array holding the spliced part.
 * @throws {RangeError} When an operation's range is outside the part, is
 * unsorted, or overlaps an earlier operation.
 */
export function applySplices(partBytes: Uint8Array, operations: readonly ApplySpliceOperation[]): Uint8Array {
  if (operations.length === 0) return partBytes.slice()
  const chunks: Uint8Array[] = []
  let cursor = 0
  for (const operation of operations) {
    switch (operation.kind) {
      case 'replace': {
        assertSorted(partBytes, operation.start, cursor)
        assertRange(partBytes, operation.start, operation.end)
        chunks.push(partBytes.subarray(cursor, operation.start), operation.bytes)
        cursor = operation.end
        break
      }
      case 'delete': {
        assertSorted(partBytes, operation.start, cursor)
        assertRange(partBytes, operation.start, operation.end)
        chunks.push(partBytes.subarray(cursor, operation.start))
        cursor = operation.end
        break
      }
      case 'insert':
      case 'append': {
        assertSorted(partBytes, operation.at, cursor)
        chunks.push(partBytes.subarray(cursor, operation.at), operation.bytes)
        cursor = operation.at
        break
      }
    }
  }
  chunks.push(partBytes.subarray(cursor))
  return concatChunks(chunks)
}

function assertSorted(partBytes: Uint8Array, start: number, cursor: number): void {
  if (!Number.isSafeInteger(start) || start < cursor || start > partBytes.byteLength) {
    throw new RangeError(`Splice operation at byte ${String(start)} is out of order or overlaps an earlier operation`)
  }
}

function assertRange(partBytes: Uint8Array, start: number, end: number): void {
  if (!Number.isSafeInteger(end) || end < start || end > partBytes.byteLength) {
    throw new RangeError(`Splice range ${String(start)}-${String(end)} is outside the ${String(partBytes.byteLength)} byte part`)
  }
}

function concatChunks(chunks: readonly Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0)
  const result = new Uint8Array(total)
  let position = 0
  for (const chunk of chunks) {
    result.set(chunk, position)
    position += chunk.byteLength
  }
  return result
}
