/** ZIP reader behavior: raw record indexing and every refused container branch. */
import { crc32, deflateRawSync } from 'node:zlib'
import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import {
  bytesEqual,
  ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR_SIGNATURE,
  ZIP_CENTRAL_DIRECTORY_HEADER_SIGNATURE,
  ZIP_CENTRAL_HEADER_LENGTH,
  ZIP_DATA_DESCRIPTOR_LENGTH,
  ZIP_DATA_DESCRIPTOR_SIGNATURE,
  ZIP_DATA_DESCRIPTOR_UNSIGNED_LENGTH,
  ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH,
  ZIP_END_OF_CENTRAL_DIRECTORY_SIGNATURE,
  ZIP_FLAG_DATA_DESCRIPTOR,
  ZIP_FLAG_ENCRYPTED,
  ZIP_FLAG_UTF8,
  ZIP_LOCAL_FILE_HEADER_SIGNATURE,
  ZIP_LOCAL_HEADER_LENGTH,
  ZipError,
  readUint32,
  writeUint16,
  writeUint32,
} from '../src/package/zip-format.ts'
import { readZipArchive } from '../src/package/zip-reader.ts'
import type { ZipMethod } from '../src/package/types.ts'

const LIMITS = { maxEntries: 100 }

interface RawSpec {
  readonly name: string
  readonly data: Uint8Array
  readonly method?: ZipMethod
  readonly flags?: number
  readonly signedDescriptor?: boolean
}

interface RawEntry {
  readonly name: string
  readonly localOffset: number
  readonly headerLength: number
  readonly dataOffset: number
  readonly compressedLength: number
  readonly uncompressedLength: number
  readonly crc32: number
  readonly flags: number
  readonly method: ZipMethod
  centralOffset: number
  centralLength: number
}

interface BuiltArchive {
  readonly bytes: Uint8Array
  readonly entries: readonly RawEntry[]
  /** Offset of the first central record. */
  readonly centralOffset: number
  /** Offset of the end of central directory record. */
  readonly eocdOffset: number
}

const TEXT = new TextEncoder()

function concat(parts: readonly Uint8Array[]): Uint8Array {
  let length = 0
  for (const part of parts) length += part.byteLength
  const result = new Uint8Array(length)
  let cursor = 0
  for (const part of parts) {
    result.set(part, cursor)
    cursor += part.byteLength
  }
  return result
}

function insert(bytes: Uint8Array, at: number, inserted: Uint8Array): Uint8Array {
  return concat([bytes.subarray(0, at), inserted, bytes.subarray(at)])
}

function buildArchive(specs: readonly RawSpec[], comment = ''): BuiltArchive {
  const locals: Uint8Array[] = []
  const entries: RawEntry[] = []
  let offset = 0

  for (const spec of specs) {
    const nameBytes = TEXT.encode(spec.name)
    const method = spec.method ?? 0
    const compressed = method === 8 ? deflateRawSync(spec.data) : spec.data
    const checksum = crc32(spec.data)
    const flags = spec.flags ?? 0
    const descriptor = (flags & ZIP_FLAG_DATA_DESCRIPTOR) !== 0
    const headerLength = ZIP_LOCAL_HEADER_LENGTH + nameBytes.byteLength
    const header = new Uint8Array(headerLength)
    writeUint32(header, 0, ZIP_LOCAL_FILE_HEADER_SIGNATURE)
    writeUint16(header, 4, 20)
    writeUint16(header, 6, flags)
    writeUint16(header, 8, method)
    writeUint16(header, 10, 0)
    writeUint16(header, 12, 0x21)
    writeUint32(header, 14, descriptor ? 0 : checksum)
    writeUint32(header, 18, descriptor ? 0 : compressed.byteLength)
    writeUint32(header, 22, descriptor ? 0 : spec.data.byteLength)
    writeUint16(header, 26, nameBytes.byteLength)
    writeUint16(header, 28, 0)
    header.set(nameBytes, ZIP_LOCAL_HEADER_LENGTH)

    const dataOffset = offset + headerLength
    const pieces = [header, compressed]
    let recordLength = headerLength + compressed.byteLength
    if (descriptor) {
      const signed = spec.signedDescriptor ?? true
      const descriptorBytes = new Uint8Array(
        signed ? ZIP_DATA_DESCRIPTOR_LENGTH : ZIP_DATA_DESCRIPTOR_UNSIGNED_LENGTH,
      )
      let at = 0
      if (signed) {
        writeUint32(descriptorBytes, 0, ZIP_DATA_DESCRIPTOR_SIGNATURE)
        at = 4
      }
      writeUint32(descriptorBytes, at, checksum)
      writeUint32(descriptorBytes, at + 4, compressed.byteLength)
      writeUint32(descriptorBytes, at + 8, spec.data.byteLength)
      pieces.push(descriptorBytes)
      recordLength += descriptorBytes.byteLength
    }

    entries.push({
      name: spec.name,
      localOffset: offset,
      headerLength,
      dataOffset,
      compressedLength: compressed.byteLength,
      uncompressedLength: spec.data.byteLength,
      crc32: checksum,
      flags,
      method,
      centralOffset: 0,
      centralLength: 0,
    })
    locals.push(...pieces)
    offset += recordLength
  }

  const centralOffset = offset
  const centrals: Uint8Array[] = []
  let centralLength = 0
  for (const entry of entries) {
    const nameBytes = TEXT.encode(entry.name)
    const central = new Uint8Array(ZIP_CENTRAL_HEADER_LENGTH + nameBytes.byteLength)
    writeUint32(central, 0, ZIP_CENTRAL_DIRECTORY_HEADER_SIGNATURE)
    writeUint16(central, 4, 20)
    writeUint16(central, 6, 20)
    writeUint16(central, 8, entry.flags)
    writeUint16(central, 10, entry.method)
    writeUint16(central, 12, 0)
    writeUint16(central, 14, 0x21)
    writeUint32(central, 16, entry.crc32)
    writeUint32(central, 20, entry.compressedLength)
    writeUint32(central, 24, entry.uncompressedLength)
    writeUint16(central, 28, nameBytes.byteLength)
    writeUint32(central, 42, entry.localOffset)
    central.set(nameBytes, ZIP_CENTRAL_HEADER_LENGTH)
    entry.centralOffset = centralOffset + centralLength
    entry.centralLength = central.byteLength
    centrals.push(central)
    centralLength += central.byteLength
  }

  const commentBytes = TEXT.encode(comment)
  const eocdOffset = centralOffset + centralLength
  const eocd = new Uint8Array(ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH + commentBytes.byteLength)
  writeUint32(eocd, 0, ZIP_END_OF_CENTRAL_DIRECTORY_SIGNATURE)
  writeUint16(eocd, 8, entries.length)
  writeUint16(eocd, 10, entries.length)
  writeUint32(eocd, 12, centralLength)
  writeUint32(eocd, 16, centralOffset)
  writeUint16(eocd, 20, commentBytes.byteLength)
  eocd.set(commentBytes, ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH)

  return {
    bytes: concat([...locals, ...centrals, eocd]),
    entries,
    centralOffset,
    eocdOffset,
  }
}

function failureOf(run: () => unknown): ZipError {
  try {
    run()
  } catch (error) {
    if (error instanceof ZipError) return error
    throw error
  }
  throw new Error('expected the reader to throw a ZipError')
}

function failureTypeOf(run: () => unknown): string {
  return failureOf(run).failure.type
}

describe('record indexing', () => {
  it('indexes a fflate archive in central and local order', () => {
    const bytes = zipSync({ 'b.txt': strToU8('bbb'), 'a.txt': strToU8('aaa') })
    const archive = readZipArchive(bytes, LIMITS)

    expect(archive.entries.map(entry => entry.name)).toEqual(['b.txt', 'a.txt'])
    expect(archive.localOrder.map(entry => entry.name)).toEqual(['b.txt', 'a.txt'])
    expect(archive.comment.byteLength).toBe(0)

    for (const entry of archive.entries) {
      expect(entry.localRecord.byteLength)
        .toBe(entry.localHeaderLength + entry.compressedSize + entry.descriptorLength)
      expect(bytesEqual(
        entry.localRecord,
        bytes.subarray(entry.localOffset, entry.localOffset + entry.localRecord.byteLength),
      )).toBe(true)
      expect(readUint32(entry.centralRecord, 0, 'central record signature'))
        .toBe(ZIP_CENTRAL_DIRECTORY_HEADER_SIGNATURE)
    }
  })

  it('indexes a hand-built stored entry', () => {
    const data = Uint8Array.of(1, 2, 3, 4)
    const built = buildArchive([{ name: 'stored.bin', data, method: 0 }])
    const archive = readZipArchive(built.bytes, LIMITS)
    const entry = archive.entries[0]!

    expect(entry.name).toBe('stored.bin')
    expect(entry.method).toBe(0)
    expect(entry.crc32).toBe(crc32(data))
    expect(entry.compressedSize).toBe(4)
    expect(entry.uncompressedSize).toBe(4)
    expect(entry.localHeaderLength).toBe(ZIP_LOCAL_HEADER_LENGTH + 'stored.bin'.length)
    expect(entry.dataOffset).toBe(built.entries[0]!.dataOffset)
    expect(entry.descriptorLength).toBe(0)
    expect(entry.localRecord.byteLength).toBe(entry.localHeaderLength + 4)
  })

  it('preserves an archive comment and scans past comment bytes', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }], 'kept comment')
    const archive = readZipArchive(built.bytes, LIMITS)

    expect(new TextDecoder().decode(archive.comment)).toBe('kept comment')
    expect(archive.entries).toHaveLength(1)
  })

  it('carries a directory entry without changing the record', () => {
    const bytes = zipSync({ 'word/': new Uint8Array(0), 'word/document.xml': strToU8('<w:document/>') })
    const archive = readZipArchive(bytes, LIMITS)

    expect(archive.entries.map(entry => entry.name)).toContain('word/')
  })
})

describe('data descriptors', () => {
  it('measures a signed descriptor', () => {
    const data = strToU8('descriptor payload')
    const built = buildArchive([{ name: 'd.bin', data, method: 8, flags: ZIP_FLAG_DATA_DESCRIPTOR }])
    const archive = readZipArchive(built.bytes, LIMITS)
    const entry = archive.entries[0]!

    expect(entry.descriptorLength).toBe(ZIP_DATA_DESCRIPTOR_LENGTH)
    expect(entry.localRecord.byteLength).toBe(
      entry.localHeaderLength + entry.compressedSize + ZIP_DATA_DESCRIPTOR_LENGTH,
    )
  })

  it('measures a signatureless descriptor', () => {
    const data = strToU8('descriptor payload')
    const built = buildArchive([{
      name: 'd.bin',
      data,
      method: 8,
      flags: ZIP_FLAG_DATA_DESCRIPTOR,
      signedDescriptor: false,
    }])
    const archive = readZipArchive(built.bytes, LIMITS)

    expect(archive.entries[0]!.descriptorLength).toBe(ZIP_DATA_DESCRIPTOR_UNSIGNED_LENGTH)
  })

  it('rejects a descriptor that disagrees with the central directory', () => {
    const built = buildArchive([{
      name: 'd.bin',
      data: Uint8Array.of(1, 2, 3),
      method: 0,
      flags: ZIP_FLAG_DATA_DESCRIPTOR,
      signedDescriptor: false,
    }])
    writeUint32(built.bytes, built.entries[0]!.dataOffset + built.entries[0]!.compressedLength, 0xdeadbeef)

    expect(failureOf(() => readZipArchive(built.bytes, LIMITS)).failure).toMatchObject({
      type: 'invalid-archive',
    })
  })

  it('rejects a descriptor whose first word is both signature and CRC', () => {
    const built = buildArchive([{ name: 'd.bin', data: Uint8Array.of(1, 2, 3), method: 0, flags: ZIP_FLAG_DATA_DESCRIPTOR }])
    writeUint32(built.bytes, built.entries[0]!.centralOffset + 16, ZIP_DATA_DESCRIPTOR_SIGNATURE)

    expect(failureOf(() => readZipArchive(built.bytes, LIMITS)).failure).toMatchObject({
      type: 'unsupported-container',
      container: 'ambiguous-descriptor',
    })
  })

  it('rejects a truncated signed descriptor', () => {
    const built = buildArchive([{
      name: 'd.bin',
      data: Uint8Array.of(1, 2, 3),
      method: 0,
      flags: ZIP_FLAG_DATA_DESCRIPTOR,
      signedDescriptor: false,
    }])
    const entry = built.entries[0]!
    writeUint32(built.bytes, entry.dataOffset + entry.compressedLength, ZIP_DATA_DESCRIPTOR_SIGNATURE)
    writeUint32(built.bytes, entry.centralOffset + 16, 0x12345678)

    expect(failureOf(() => readZipArchive(built.bytes, LIMITS)).failure).toMatchObject({
      type: 'out-of-bounds',
    })
  })
})

describe('names', () => {
  it('decodes a UTF-8 name when the flag is set', () => {
    const built = buildArchive([{ name: 'héllo.xml', data: Uint8Array.of(1), flags: ZIP_FLAG_UTF8 }])
    const archive = readZipArchive(built.bytes, LIMITS)

    expect(archive.entries[0]!.name).toBe('héllo.xml')
  })

  it('rejects an invalid UTF-8 name when the flag is set', () => {
    const built = buildArchive([{ name: 'a', data: Uint8Array.of(1), flags: ZIP_FLAG_UTF8 }])
    const entry = built.entries[0]!
    built.bytes[entry.centralOffset + ZIP_CENTRAL_HEADER_LENGTH] = 0xff

    expect(failureTypeOf(() => readZipArchive(built.bytes, LIMITS))).toBe('invalid-archive')
  })

  it('rejects duplicate and ASCII case-equivalent names', () => {
    const duplicate = buildArchive([
      { name: 'word/document.xml', data: Uint8Array.of(1) },
      { name: 'word/document.xml', data: Uint8Array.of(2) },
    ])
    expect(failureTypeOf(() => readZipArchive(duplicate.bytes, LIMITS))).toBe('duplicate-name')

    const folded = buildArchive([
      { name: 'word/document.xml', data: Uint8Array.of(1) },
      { name: 'WORD/Document.XML', data: Uint8Array.of(2) },
    ])
    expect(failureTypeOf(() => readZipArchive(folded.bytes, LIMITS))).toBe('duplicate-name')
  })
})

describe('container refusals', () => {
  it('rejects a CFB wrapper', () => {
    const bytes = Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0)
    expect(failureOf(() => readZipArchive(bytes, LIMITS)).failure).toMatchObject({
      type: 'unsupported-container',
      container: 'cfb',
    })
  })

  it('rejects ZIP64 sentinel fields', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    writeUint16(built.bytes, built.eocdOffset + 8, 0xffff)
    writeUint16(built.bytes, built.eocdOffset + 10, 0xffff)

    expect(failureOf(() => readZipArchive(built.bytes, LIMITS)).failure).toMatchObject({
      type: 'unsupported-container',
      container: 'zip64',
    })
  })

  it('rejects a ZIP64 locator', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    const locator = new Uint8Array(20)
    writeUint32(locator, 0, ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR_SIGNATURE)
    const bytes = insert(built.bytes, built.eocdOffset, locator)

    expect(failureOf(() => readZipArchive(bytes, LIMITS)).failure).toMatchObject({
      type: 'unsupported-container',
      container: 'zip64',
    })
  })

  it('rejects a multi-disk archive', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    writeUint16(built.bytes, built.eocdOffset + 4, 1)

    expect(failureOf(() => readZipArchive(built.bytes, LIMITS)).failure).toMatchObject({
      type: 'unsupported-container',
      container: 'multi-disk',
    })
  })

  it('rejects an encrypted entry', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    const flags = built.entries[0]!.flags | ZIP_FLAG_ENCRYPTED
    writeUint16(built.bytes, built.entries[0]!.centralOffset + 8, flags)

    expect(failureOf(() => readZipArchive(built.bytes, LIMITS)).failure).toMatchObject({
      type: 'unsupported-container',
      container: 'encrypted',
    })
  })

  it('rejects an unsupported compression method', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    writeUint16(built.bytes, built.entries[0]!.centralOffset + 10, 12)

    expect(failureOf(() => readZipArchive(built.bytes, LIMITS)).failure).toMatchObject({
      type: 'unsupported-container',
      container: 'unsupported-method',
    })
  })

  it('rejects an archive over the entry bound', () => {
    const built = buildArchive([
      { name: 'a.bin', data: Uint8Array.of(1) },
      { name: 'b.bin', data: Uint8Array.of(2) },
    ])
    const error = failureOf(() => readZipArchive(built.bytes, { maxEntries: 1 }))

    expect(error.failure).toMatchObject({ type: 'entry-limit' })
  })
})

describe('structural refusals', () => {
  it('rejects a file shorter than an end record and a missing end record', () => {
    expect(failureTypeOf(() => readZipArchive(Uint8Array.of(1, 2, 3), LIMITS))).toBe('invalid-archive')

    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    expect(failureTypeOf(() => readZipArchive(built.bytes.subarray(0, built.bytes.byteLength - 4), LIMITS)))
      .toBe('invalid-archive')
  })

  it('ignores an end signature whose comment length does not reach the file end', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    writeUint16(built.bytes, built.eocdOffset + 20, 3)
    const bytes = concat([built.bytes, Uint8Array.of(0, 0, 0, 0)])

    expect(failureTypeOf(() => readZipArchive(bytes, LIMITS))).toBe('invalid-archive')
  })

  it('rejects a central directory that overlaps the end record', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    writeUint32(built.bytes, built.eocdOffset + 16, built.bytes.byteLength)

    expect(failureOf(() => readZipArchive(built.bytes, LIMITS)).failure.type).toBe('out-of-bounds')
  })

  it('rejects a central record with the wrong signature', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    built.bytes[built.centralOffset] = 0

    expect(failureTypeOf(() => readZipArchive(built.bytes, LIMITS))).toBe('invalid-archive')
  })

  it('rejects a central size that does not end after the parsed records', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    writeUint32(built.bytes, built.eocdOffset + 12, built.entries[0]!.centralLength - 1)

    expect(failureTypeOf(() => readZipArchive(built.bytes, LIMITS))).toBe('invalid-archive')
  })

  it('rejects a local offset inside the central directory', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    writeUint32(built.bytes, built.entries[0]!.centralOffset + 42, built.centralOffset)

    expect(failureTypeOf(() => readZipArchive(built.bytes, LIMITS))).toBe('out-of-bounds')
  })

  it('rejects a local offset that does not point at a local header', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    writeUint32(built.bytes, built.entries[0]!.centralOffset + 42, built.centralOffset - 5)

    expect(failureTypeOf(() => readZipArchive(built.bytes, LIMITS))).toBe('invalid-archive')
  })

  it('rejects a local header whose name and extra fields exceed the file', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    writeUint16(built.bytes, built.entries[0]!.localOffset + 26, 0xffff)

    expect(failureTypeOf(() => readZipArchive(built.bytes, LIMITS))).toBe('out-of-bounds')
  })

  it('rejects data that runs into the central directory', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    const entry = built.entries[0]!
    writeUint32(built.bytes, entry.localOffset + 18, 0x1000)
    writeUint32(built.bytes, entry.centralOffset + 20, 0x1000)

    expect(failureTypeOf(() => readZipArchive(built.bytes, LIMITS))).toBe('out-of-bounds')
  })
})

describe('local and central agreement', () => {
  it('rejects a local method that differs from the central method', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    writeUint16(built.bytes, built.entries[0]!.localOffset + 8, 8)

    expect(failureTypeOf(() => readZipArchive(built.bytes, LIMITS))).toBe('invalid-archive')
  })

  it('rejects local flags that differ from the central flags', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    writeUint16(built.bytes, built.entries[0]!.localOffset + 6, 0x0800)

    expect(failureTypeOf(() => readZipArchive(built.bytes, LIMITS))).toBe('invalid-archive')
  })

  it('rejects a local name that differs from the central name', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    built.bytes[built.entries[0]!.localOffset + ZIP_LOCAL_HEADER_LENGTH] = 0x62

    expect(failureTypeOf(() => readZipArchive(built.bytes, LIMITS))).toBe('invalid-archive')
  })

  it('rejects local sizes and CRC that differ from the central record', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    writeUint32(built.bytes, built.entries[0]!.localOffset + 14, 0)

    expect(failureTypeOf(() => readZipArchive(built.bytes, LIMITS))).toBe('invalid-archive')
  })

  it('rejects a stored entry whose declared sizes differ', () => {
    const built = buildArchive([{ name: 'a.bin', data: Uint8Array.of(1) }])
    const entry = built.entries[0]!
    writeUint32(built.bytes, entry.localOffset + 22, 2)
    writeUint32(built.bytes, entry.centralOffset + 24, 2)

    expect(failureTypeOf(() => readZipArchive(built.bytes, LIMITS))).toBe('invalid-archive')
  })
})
