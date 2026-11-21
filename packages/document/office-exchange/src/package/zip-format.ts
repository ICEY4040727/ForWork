/** ZIP container constants, DOS timestamps, and bounds-checked field access. */
import type { DosDateTime, ZipFailure } from './types.ts'

/** Local file header signature. */
export const ZIP_LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50

/** Central directory header signature. */
export const ZIP_CENTRAL_DIRECTORY_HEADER_SIGNATURE = 0x02014b50

/** End of central directory signature. */
export const ZIP_END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50

/** Optional data descriptor signature. */
export const ZIP_DATA_DESCRIPTOR_SIGNATURE = 0x08074b50

/** ZIP64 end of central directory locator signature. */
export const ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR_SIGNATURE = 0x07064b50

/** ZIP64 end of central directory signature. */
export const ZIP64_END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06064b50

/** Compound File Binary signature used by OLE wrappers. */
export const CFB_SIGNATURE: readonly number[] = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]

/** Fixed local file header length before its name and extra fields. */
export const ZIP_LOCAL_HEADER_LENGTH = 30

/** Fixed central directory header length before its variable fields. */
export const ZIP_CENTRAL_HEADER_LENGTH = 46

/** Fixed end of central directory length before its comment. */
export const ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH = 22

/** Signed data descriptor length. */
export const ZIP_DATA_DESCRIPTOR_LENGTH = 16

/** Signatureless data descriptor length. */
export const ZIP_DATA_DESCRIPTOR_UNSIGNED_LENGTH = 12

/** General-purpose bit 0: the entry is encrypted. */
export const ZIP_FLAG_ENCRYPTED = 0x0001

/** General-purpose bit 3: sizes and CRC follow the data as a descriptor. */
export const ZIP_FLAG_DATA_DESCRIPTOR = 0x0008

/** General-purpose bit 6: strong encryption is in use. */
export const ZIP_FLAG_STRONG_ENCRYPTION = 0x0040

/** General-purpose bit 11: the entry name is UTF-8. */
export const ZIP_FLAG_UTF8 = 0x0800

/** General-purpose bit 13: local header values are masked. */
export const ZIP_FLAG_MASKED_HEADERS = 0x2000

/** Stored (uncompressed) method. */
export const ZIP_METHOD_STORE = 0

/** Deflate compression method. */
export const ZIP_METHOD_DEFLATE = 8

/** Sentinel stored in a 32-bit size or offset field when ZIP64 is in use. */
export const ZIP_UINT32_MAX = 0xffffffff

/** Sentinel stored in a 16-bit entry-count field when ZIP64 is in use. */
export const ZIP_UINT16_MAX = 0xffff

/** Local header flags field offset. */
export const ZIP_LOCAL_FLAGS_OFFSET = 6

/** Local header CRC-32 field offset. */
export const ZIP_LOCAL_CRC32_OFFSET = 14

/** Local header compressed-size field offset. */
export const ZIP_LOCAL_COMPRESSED_SIZE_OFFSET = 18

/** Local header uncompressed-size field offset. */
export const ZIP_LOCAL_UNCOMPRESSED_SIZE_OFFSET = 22

/** Central header flags field offset. */
export const ZIP_CENTRAL_FLAGS_OFFSET = 8

/** Central header CRC-32 field offset. */
export const ZIP_CENTRAL_CRC32_OFFSET = 16

/** Central header compressed-size field offset. */
export const ZIP_CENTRAL_COMPRESSED_SIZE_OFFSET = 20

/** Central header uncompressed-size field offset. */
export const ZIP_CENTRAL_UNCOMPRESSED_SIZE_OFFSET = 24

/** Central header local-offset field offset. */
export const ZIP_CENTRAL_LOCAL_OFFSET_OFFSET = 42

/** Fixed DOS timestamp used when a caller provides none. */
export const DEFAULT_DOS_DATE_TIME: DosDateTime = { time: 0, date: 0x0021 }

/** A typed ZIP container failure carried by {@link ZipError}. */
export class ZipError extends Error {
  /**
   * @param failure - Structured reason the container was refused.
   */
  constructor(readonly failure: ZipFailure) {
    super(failure.detail)
    this.name = 'ZipError'
  }
}

/**
 * Throw a typed ZIP failure.
 * @param failure - Structured reason the operation cannot continue.
 * @returns Never; the function always throws.
 */
export function failZip(failure: ZipFailure): never {
  throw new ZipError(failure)
}

/**
 * Continue only when a container invariant holds.
 * @param condition - Invariant that must be true.
 * @param failure - Structured reason reported when it is false.
 * @returns Nothing when the condition holds.
 */
export function assertZip(condition: boolean, failure: ZipFailure): asserts condition {
  if (!condition) throw new ZipError(failure)
}

/**
 * Whether a byte sequence starts with a fixed signature.
 * @param bytes - Buffer to inspect.
 * @param offset - Offset to compare at.
 * @param signature - Expected byte values.
 * @returns `true` when every signature byte matches at the offset.
 */
export function hasSignature(bytes: Uint8Array, offset: number, signature: readonly number[]): boolean {
  if (offset < 0 || offset + signature.length > bytes.byteLength) return false
  for (let index = 0; index < signature.length; index += 1) {
    if (bytes[offset + index] !== signature[index]) return false
  }
  return true
}

/**
 * Whether a numeric compression method is supported.
 * @param method - Method value read from a header.
 * @returns `true` for store and deflate.
 */
export function isZipMethod(method: number): method is 0 | 8 {
  return method === ZIP_METHOD_STORE || method === ZIP_METHOD_DEFLATE
}

/**
 * Throw when a byte range is not inside the buffer.
 * @param bytes - Buffer being read.
 * @param offset - Start of the requested range.
 * @param length - Requested range length.
 * @param context - Field or record being read, used in the failure detail.
 * @returns Nothing when the range is in bounds.
 */
export function assertRange(bytes: Uint8Array, offset: number, length: number, context: string): void {
  if (offset < 0 || offset + length > bytes.byteLength) {
    throw new ZipError({
      type: 'out-of-bounds',
      detail: `${context} at ${offset} with length ${length} exceeds the ${bytes.byteLength}-byte archive`,
    })
  }
}

/**
 * Read a little-endian 16-bit field.
 * @param bytes - Buffer to read.
 * @param offset - Field offset.
 * @param context - Field being read, used in the failure detail.
 * @returns The unsigned field value.
 */
export function readUint16(bytes: Uint8Array, offset: number, context: string): number {
  assertRange(bytes, offset, 2, context)
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint16(offset, true)
}

/**
 * Read a little-endian 32-bit field.
 * @param bytes - Buffer to read.
 * @param offset - Field offset.
 * @param context - Field being read, used in the failure detail.
 * @returns The unsigned field value.
 */
export function readUint32(bytes: Uint8Array, offset: number, context: string): number {
  assertRange(bytes, offset, 4, context)
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(offset, true)
}

/**
 * Write a little-endian 16-bit field into an already-sized buffer.
 * @param target - Buffer to write.
 * @param offset - Field offset.
 * @param value - Value to encode; only the low 16 bits are used.
 * @returns Nothing.
 */
export function writeUint16(target: Uint8Array, offset: number, value: number): void {
  target[offset] = value & 0xff
  target[offset + 1] = (value >>> 8) & 0xff
}

/**
 * Write a little-endian 32-bit field into an already-sized buffer.
 * @param target - Buffer to write.
 * @param offset - Field offset.
 * @param value - Value to encode; only the low 32 bits are used.
 * @returns Nothing.
 */
export function writeUint32(target: Uint8Array, offset: number, value: number): void {
  target[offset] = value & 0xff
  target[offset + 1] = (value >>> 8) & 0xff
  target[offset + 2] = (value >>> 16) & 0xff
  target[offset + 3] = (value >>> 24) & 0xff
}

/**
 * Compare two byte sequences.
 * @param left - First sequence.
 * @param right - Second sequence.
 * @returns `true` when both sequences have equal length and bytes.
 */
export function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false
  for (let index = 0; index < left.byteLength; index += 1) {
    if (left[index] !== right[index]) return false
  }
  return true
}

/**
 * Build a lookup key that folds ASCII `A`-`Z` to `a`-`z`.
 * @param name - Raw entry-name bytes.
 * @returns A one-character-per-byte key whose bytes are ASCII case-folded.
 */
export function asciiFoldKey(name: Uint8Array): string {
  let key = ''
  for (let index = 0; index < name.byteLength; index += 1) {
    const value = name[index] as number
    key += String.fromCharCode(value >= 0x41 && value <= 0x5a ? value + 0x20 : value)
  }
  return key
}

/**
 * Decode a DOS date and time pair as UTC.
 * @param dos - DOS date and time fields.
 * @returns The equivalent UTC instant.
 */
export function dosDateTimeToDate(dos: DosDateTime): Date {
  const day = dos.date & 0x1f
  const month = (dos.date >>> 5) & 0x0f
  const year = 1980 + ((dos.date >>> 9) & 0x7f)
  const seconds = (dos.time & 0x1f) * 2
  const minutes = (dos.time >>> 5) & 0x3f
  const hours = (dos.time >>> 11) & 0x1f
  return new Date(Date.UTC(year, month - 1, day, hours, minutes, seconds))
}

/**
 * Encode a UTC instant as a DOS date and time pair.
 * @param date - Instant to encode; years outside 1980-2107 are clamped.
 * @returns The DOS date and time fields.
 */
export function dateToDosDateTime(date: Date): DosDateTime {
  const year = Math.min(Math.max(date.getUTCFullYear() - 1980, 0), 127)
  const month = date.getUTCMonth() + 1
  const day = date.getUTCDate()
  const hours = date.getUTCHours()
  const minutes = date.getUTCMinutes()
  const seconds = Math.floor(date.getUTCSeconds() / 2)
  return {
    date: (year << 9) | (month << 5) | day,
    time: (hours << 11) | (minutes << 5) | seconds,
  }
}
