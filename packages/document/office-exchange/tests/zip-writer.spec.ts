/** ZIP writer behavior: no-op identity, carried records, changed records, additions. */
import { crc32 } from 'node:zlib'
import { strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import {
  bytesEqual,
  dateToDosDateTime,
  dosDateTimeToDate,
  readUint16,
  ZIP_CENTRAL_DIRECTORY_HEADER_SIGNATURE,
  ZIP_DATA_DESCRIPTOR_LENGTH,
  ZIP_DATA_DESCRIPTOR_SIGNATURE,
  ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH,
  ZIP_END_OF_CENTRAL_DIRECTORY_SIGNATURE,
  ZIP_FLAG_DATA_DESCRIPTOR,
  ZIP_FLAG_UTF8,
  ZIP_LOCAL_FILE_HEADER_SIGNATURE,
  ZipError,
  writeUint16,
  writeUint32,
} from '../src/package/zip-format.ts'
import { readZipArchive } from '../src/package/zip-reader.ts'
import { writeZipArchive } from '../src/package/zip-writer.ts'
import type { ZipArchive, ZipEntryRecord } from '../src/package/types.ts'

const LIMITS = { maxEntries: 200 }
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

function withComment(bytes: Uint8Array, comment: string): Uint8Array {
  const commentBytes = TEXT.encode(comment)
  const result = new Uint8Array(bytes.byteLength + commentBytes.byteLength)
  result.set(bytes)
  writeUint16(result, bytes.byteLength - ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH + 20, commentBytes.byteLength)
  result.set(commentBytes, bytes.byteLength)
  return result
}

function buildDescriptorArchive(name: string, data: Uint8Array): Uint8Array {
  const nameBytes = TEXT.encode(name)
  const checksum = crc32(data)
  const header = new Uint8Array(30 + nameBytes.byteLength)
  writeUint32(header, 0, ZIP_LOCAL_FILE_HEADER_SIGNATURE)
  writeUint16(header, 4, 20)
  writeUint16(header, 6, ZIP_FLAG_DATA_DESCRIPTOR)
  writeUint16(header, 8, 0)
  writeUint16(header, 10, 0)
  writeUint16(header, 12, 0x21)
  writeUint16(header, 26, nameBytes.byteLength)
  header.set(nameBytes, 30)

  const descriptor = new Uint8Array(ZIP_DATA_DESCRIPTOR_LENGTH)
  writeUint32(descriptor, 0, ZIP_DATA_DESCRIPTOR_SIGNATURE)
  writeUint32(descriptor, 4, checksum)
  writeUint32(descriptor, 8, data.byteLength)
  writeUint32(descriptor, 12, data.byteLength)

  const central = new Uint8Array(46 + nameBytes.byteLength)
  writeUint32(central, 0, ZIP_CENTRAL_DIRECTORY_HEADER_SIGNATURE)
  writeUint16(central, 4, 20)
  writeUint16(central, 6, 20)
  writeUint16(central, 8, ZIP_FLAG_DATA_DESCRIPTOR)
  writeUint16(central, 10, 0)
  writeUint16(central, 12, 0)
  writeUint16(central, 14, 0x21)
  writeUint32(central, 16, checksum)
  writeUint32(central, 20, data.byteLength)
  writeUint32(central, 24, data.byteLength)
  writeUint16(central, 28, nameBytes.byteLength)
  writeUint32(central, 42, 0)
  central.set(nameBytes, 46)

  const centralOffset = header.byteLength + data.byteLength + descriptor.byteLength
  const eocd = new Uint8Array(ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH)
  writeUint32(eocd, 0, ZIP_END_OF_CENTRAL_DIRECTORY_SIGNATURE)
  writeUint16(eocd, 8, 1)
  writeUint16(eocd, 10, 1)
  writeUint32(eocd, 12, central.byteLength)
  writeUint32(eocd, 16, centralOffset)
  return concat([header, data, descriptor, central, eocd])
}

function failureOf(run: () => unknown): ZipError {
  try {
    run()
  } catch (error) {
    if (error instanceof ZipError) return error
    throw error
  }
  throw new Error('expected the writer to throw a ZipError')
}

function entryByName(archive: ZipArchive, name: string): ZipEntryRecord {
  const entry = archive.entries.find(candidate => candidate.name === name)
  if (entry === undefined) throw new Error(`missing entry ${name}`)
  return entry
}

describe('identity and carried records', () => {
  it('returns the exact source bytes when nothing changes', () => {
    const source = zipSync({ 'a.txt': strToU8('aaa'), 'b.txt': strToU8('bbb') })
    const archive = readZipArchive(source, LIMITS)

    expect(writeZipArchive(archive, {})).toBe(source)
    expect(writeZipArchive(archive, { replacements: [], additions: [], removals: [] })).toBe(source)
  })

  it('copies untouched local records and rewrites only the changed entry', () => {
    const source = zipSync({
      'a.txt': strToU8('aaa'),
      'word/document.xml': strToU8('<w:document>old</w:document>'),
      'z.bin': [Uint8Array.of(9, 9, 9), { level: 0 }],
    })
    const archive = readZipArchive(source, LIMITS)
    const before = new Map(archive.entries.map(entry => [entry.name, entry]))
    const replacement = strToU8('<w:document>a much longer replacement payload</w:document>')

    const output = writeZipArchive(archive, {
      replacements: [{ name: 'word/document.xml', bytes: replacement }],
    })
    const written = readZipArchive(output, LIMITS)

    for (const entry of written.entries) {
      if (entry.name === 'word/document.xml') continue
      expect(bytesEqual(entry.localRecord, before.get(entry.name)!.localRecord)).toBe(true)
    }

    const changed = entryByName(written, 'word/document.xml')
    const original = before.get('word/document.xml')!
    expect(changed.method).toBe(original.method)
    expect(changed.localHeaderLength).toBe(original.localHeaderLength)
    expect(changed.crc32).toBe(crc32(replacement))
    expect(changed.uncompressedSize).toBe(replacement.byteLength)
    expect(changed.flags & ZIP_FLAG_DATA_DESCRIPTOR).toBe(0)
    expect(changed.localRecord.byteLength)
      .toBe(changed.localHeaderLength + changed.compressedSize)
    expect(unzipSync(output)['word/document.xml']).toEqual(replacement)
    expect(entryByName(written, 'z.bin').localOffset)
      .not.toBe(before.get('z.bin')!.localOffset)
    expect(Array.from(entryByName(written, 'a.txt').localRecord))
      .toEqual(Array.from(entryByName(archive, 'a.txt').localRecord))
  })

  it('clears the descriptor flag and drops a source descriptor', () => {
    const name = 'word/document.xml'
    const source = buildDescriptorArchive(name, strToU8('descriptor body'))
    const archive = readZipArchive(source, LIMITS)
    expect(archive.entries[0]!.descriptorLength).toBe(ZIP_DATA_DESCRIPTOR_LENGTH)

    const replacement = strToU8('rewritten')
    const output = writeZipArchive(archive, {
      replacements: [{ name, bytes: replacement }],
    })
    const written = readZipArchive(output, LIMITS)
    const entry = written.entries[0]!

    expect(entry.flags & ZIP_FLAG_DATA_DESCRIPTOR).toBe(0)
    expect(entry.descriptorLength).toBe(0)
    expect(entry.localHeaderLength).toBe(archive.entries[0]!.localHeaderLength)
    expect(entry.localRecord.byteLength).toBe(entry.localHeaderLength + entry.compressedSize)
    expect(unzipSync(output)[name]).toEqual(replacement)
  })

  it('preserves the archive comment across a change', () => {
    const source = withComment(zipSync({ 'a.txt': strToU8('aaa') }), 'archive comment')
    const archive = readZipArchive(source, LIMITS)
    expect(new TextDecoder().decode(archive.comment)).toBe('archive comment')

    const output = writeZipArchive(archive, {
      replacements: [{ name: 'a.txt', bytes: strToU8('changed') }],
    })
    const written = readZipArchive(output, LIMITS)

    expect(new TextDecoder().decode(written.comment)).toBe('archive comment')
    expect(unzipSync(output)['a.txt']).toEqual(strToU8('changed'))
  })

  it('keeps a stored entry stored when it changes', () => {
    const source = zipSync({ 'stored.bin': [Uint8Array.of(1, 2, 3), { level: 0 }] })
    const archive = readZipArchive(source, LIMITS)
    const replacement = Uint8Array.of(4, 5, 6, 7)

    const output = writeZipArchive(archive, {
      replacements: [{ name: 'stored.bin', bytes: replacement }],
    })
    const entry = readZipArchive(output, LIMITS).entries[0]!

    expect(entry.method).toBe(0)
    expect(entry.compressedSize).toBe(replacement.byteLength)
    expect(entry.uncompressedSize).toBe(replacement.byteLength)
    expect(entry.crc32).toBe(crc32(replacement))
    expect(Array.from(output.subarray(entry.dataOffset, entry.dataOffset + entry.compressedSize)))
      .toEqual(Array.from(replacement))
  })
})

describe('additions and removals', () => {
  it('appends additions in caller order with deterministic timestamps', () => {
    const source = zipSync({ 'a.txt': strToU8('aaa') })
    const archive = readZipArchive(source, LIMITS)
    const dos = { time: 0x4a5f, date: 0x5a6b }
    const changes = {
      additions: [
        { name: 'word/new.bin', bytes: Uint8Array.of(1, 2, 3), method: 0 as const },
        { name: 'héllo.txt', bytes: strToU8('hello') },
      ],
    }

    const output = writeZipArchive(archive, changes, {
      deflateLevel: 1,
      addedEntryDosDateTime: dos,
    })
    const written = readZipArchive(output, LIMITS)

    expect(written.entries.map(entry => entry.name))
      .toEqual(['a.txt', 'word/new.bin', 'héllo.txt'])
    const stored = entryByName(written, 'word/new.bin')
    const deflated = entryByName(written, 'héllo.txt')
    expect(stored.method).toBe(0)
    expect(stored.flags & ZIP_FLAG_UTF8).toBe(0)
    expect(deflated.method).toBe(8)
    expect(deflated.flags & ZIP_FLAG_UTF8).toBe(ZIP_FLAG_UTF8)
    expect(readUint16(output, stored.localOffset + 10, 'stored time')).toBe(dos.time)
    expect(readUint16(output, stored.localOffset + 12, 'stored date')).toBe(dos.date)
    expect(readUint16(deflated.centralRecord, 12, 'central time')).toBe(dos.time)
    expect(readUint16(deflated.centralRecord, 14, 'central date')).toBe(dos.date)
    expect(unzipSync(output)['word/new.bin']).toEqual(Uint8Array.of(1, 2, 3))
    expect(unzipSync(output)['héllo.txt']).toEqual(strToU8('hello'))
    expect(bytesEqual(
      output,
      writeZipArchive(archive, changes, { deflateLevel: 1, addedEntryDosDateTime: dos }),
    )).toBe(true)
  })

  it('uses the fixed default timestamp when none is provided', () => {
    const archive = readZipArchive(zipSync({ 'a.txt': strToU8('aaa') }), LIMITS)
    const output = writeZipArchive(archive, {
      additions: [{ name: 'new.bin', bytes: Uint8Array.of(1) }],
    })
    const added = entryByName(readZipArchive(output, LIMITS), 'new.bin')

    expect(readUint16(output, added.localOffset + 10, 'local time')).toBe(0)
    expect(readUint16(output, added.localOffset + 12, 'local date')).toBe(0x21)
  })

  it('removes entries from both sections', () => {
    const source = zipSync({
      'a.txt': strToU8('aaa'),
      'b.txt': strToU8('bbb'),
    })
    const archive = readZipArchive(source, LIMITS)

    const output = writeZipArchive(archive, { removals: ['a.txt'] })
    const written = readZipArchive(output, LIMITS)

    expect(written.entries.map(entry => entry.name)).toEqual(['b.txt'])
    expect(Object.keys(unzipSync(output))).toEqual(['b.txt'])
  })

  it('rejects unknown and duplicate changes', () => {
    const archive = readZipArchive(zipSync({
      'a.txt': strToU8('aaa'),
      'word/document.xml': strToU8('doc'),
    }), LIMITS)

    expect(failureOf(() => writeZipArchive(archive, {
      replacements: [{ name: 'missing.xml', bytes: Uint8Array.of(1) }],
    })).failure.type).toBe('unknown-entry')
    expect(failureOf(() => writeZipArchive(archive, { removals: ['missing.xml'] })).failure.type)
      .toBe('unknown-entry')
    expect(failureOf(() => writeZipArchive(archive, {
      replacements: [{ name: 'a.txt', bytes: Uint8Array.of(1) }],
      removals: ['a.txt'],
    })).failure.type).toBe('unknown-entry')
    expect(failureOf(() => writeZipArchive(archive, {
      replacements: [
        { name: 'a.txt', bytes: Uint8Array.of(1) },
        { name: 'a.txt', bytes: Uint8Array.of(2) },
      ],
    })).failure.type).toBe('duplicate-name')
  })

  it('rejects addition names that collide under ASCII case folding', () => {
    const archive = readZipArchive(zipSync({ 'word/document.xml': strToU8('doc') }), LIMITS)

    expect(failureOf(() => writeZipArchive(archive, {
      additions: [{ name: 'WORD/Document.XML', bytes: Uint8Array.of(1) }],
    })).failure.type).toBe('name-collision')
    expect(failureOf(() => writeZipArchive(archive, {
      additions: [
        { name: 'new.bin', bytes: Uint8Array.of(1) },
        { name: 'NEW.BIN', bytes: Uint8Array.of(2) },
      ],
    })).failure.type).toBe('name-collision')
  })

  it('rejects an output that needs ZIP64 entry counts', () => {
    const entries: ZipEntryRecord[] = []
    for (let index = 0; index < 0xffff; index += 1) {
      const name = `e${index}`
      entries.push({
        name,
        nameBytes: TEXT.encode(name),
        method: 0,
        flags: 0,
        crc32: 0,
        compressedSize: 0,
        uncompressedSize: 0,
        localOffset: index,
        localHeaderLength: 30,
        dataOffset: 30,
        descriptorLength: 0,
        localRecord: new Uint8Array(0),
        centralRecord: new Uint8Array(46),
      })
    }
    const archive: ZipArchive = {
      bytes: new Uint8Array(0),
      entries,
      localOrder: entries,
      centralOrder: entries,
      comment: new Uint8Array(0),
    }

    expect(failureOf(() => writeZipArchive(archive, {
      replacements: [{ name: 'e0', bytes: new Uint8Array(0) }],
    })).failure.type).toBe('entry-limit')
  })
})

describe('DOS timestamps', () => {
  it('round-trips a UTC instant', () => {
    const instant = new Date(Date.UTC(2026, 10, 18, 9, 47, 36))
    const dos = dateToDosDateTime(instant)

    expect(dosDateTimeToDate(dos).toISOString()).toBe('2026-11-18T09:47:36.000Z')
  })
})
