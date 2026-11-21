/** OfficePackage behavior: part map, dirty tracking, no-op save, and verification. */
import { createHash } from 'node:crypto'
import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { assertOpcGraph, open } from '../src/package/ooxml-package.ts'
import {
  bytesEqual,
  readUint16,
  readUint32,
  ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH,
  ZipError,
  writeUint16,
  writeUint32,
} from '../src/package/zip-format.ts'
import { readZipArchive } from '../src/package/zip-reader.ts'
import type { ZipEntryRecord } from '../src/package/types.ts'

const LIMITS = { maxEntries: 100 }
const TEXT = new TextEncoder()

function sourceOf(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8('<Types/>'),
    '_rels/.rels': strToU8('<Relationships/>'),
    'word/document.xml': strToU8('<w:document>Hello from the source package</w:document>'),
    'word/media/image1.bin': [Uint8Array.of(1, 2, 3, 4), { level: 0 }],
    'word/': new Uint8Array(0),
  })
}

function withComment(bytes: Uint8Array, comment: string): Uint8Array {
  const commentBytes = TEXT.encode(comment)
  const result = new Uint8Array(bytes.byteLength + commentBytes.byteLength)
  result.set(bytes)
  writeUint16(result, bytes.byteLength - ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH + 20, commentBytes.byteLength)
  result.set(commentBytes, bytes.byteLength)
  return result
}

interface CentralLocation {
  readonly offset: number
  readonly localOffset: number
}

function centralLocations(bytes: Uint8Array): Map<string, CentralLocation> {
  const eocdOffset = bytes.byteLength - ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH
  const count = readUint16(bytes, eocdOffset + 10, 'entry count')
  let cursor = readUint32(bytes, eocdOffset + 16, 'central offset')
  const locations = new Map<string, CentralLocation>()
  for (let index = 0; index < count; index += 1) {
    const nameLength = readUint16(bytes, cursor + 28, 'name length')
    const extraLength = readUint16(bytes, cursor + 30, 'extra length')
    const commentLength = readUint16(bytes, cursor + 32, 'comment length')
    const name = new TextDecoder().decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength))
    locations.set(name, {
      offset: cursor,
      localOffset: readUint32(bytes, cursor + 42, 'local offset'),
    })
    cursor += 46 + nameLength + extraLength + commentLength
  }
  return locations
}

function entryByName(bytes: Uint8Array, name: string): ZipEntryRecord {
  const entry = readZipArchive(bytes, LIMITS).entries.find(candidate => candidate.name === name)
  if (entry === undefined) throw new Error(`missing entry ${name}`)
  return entry
}

function failureOf(run: () => unknown): ZipError {
  try {
    run()
  } catch (error) {
    if (error instanceof ZipError) return error
    throw error
  }
  throw new Error('expected the package to throw a ZipError')
}

function corruptEntryCrc(bytes: Uint8Array, name: string): Uint8Array {
  const location = centralLocations(bytes).get(name)
  if (location === undefined) throw new Error(`package has no entry ${name}`)
  const checksum = readUint32(bytes, location.localOffset + 14, 'local CRC')
  writeUint32(bytes, location.localOffset + 14, checksum ^ 0xff)
  writeUint32(bytes, location.offset + 16, checksum ^ 0xff)
  return bytes
}

describe('opening and reading', () => {
  it('reads parts and records the source digest', () => {
    const source = sourceOf()
    const pkg = open(source, LIMITS)

    expect(pkg.source.bytes).toBe(source)
    expect(pkg.source.digest).toBe(createHash('sha256').update(source).digest('hex'))
    expect(pkg.readPart('word/document.xml'))
      .toEqual(strToU8('<w:document>Hello from the source package</w:document>'))
    expect(pkg.readPart('WORD/DOCUMENT.XML')).toEqual(pkg.readPart('word/document.xml'))
  })

  it('keeps directory entries out of the part map', () => {
    const pkg = open(sourceOf(), LIMITS)

    expect(pkg.source.entries.map(entry => entry.name)).toContain('word/')
    expect(failureOf(() => pkg.readPart('word/')).failure.type).toBe('unknown-entry')
    expect(failureOf(() => pkg.readPart('word/missing.xml')).failure.type).toBe('unknown-entry')
  })

  it('rejects duplicate case-equivalent source names', () => {
    const source = zipSync({
      'word/document.xml': strToU8('one'),
      'WORD/Document.XML': strToU8('two'),
    })

    expect(failureOf(() => open(source, LIMITS)).failure.type).toBe('duplicate-name')
  })

  it('rejects out-of-bounds source offsets', () => {
    const source = sourceOf()
    const location = centralLocations(source).get('word/document.xml')!
    writeUint32(source, location.offset + 42, source.byteLength)

    expect(failureOf(() => open(source, LIMITS)).failure.type).toBe('out-of-bounds')
  })

  it('applies the entry bound', () => {
    expect(failureOf(() => open(sourceOf(), { maxEntries: 1 })).failure.type).toBe('entry-limit')
  })
})

describe('no-op save', () => {
  it('returns the exact source bytes when nothing changes', () => {
    const source = sourceOf()
    const pkg = open(source, LIMITS)

    expect(pkg.save()).toBe(source)
  })

  it('reads pending additions and replacements before saving', () => {
    const pkg = open(sourceOf(), LIMITS)
    const replacement = strToU8('<w:document>Pending replacement</w:document>')

    pkg.addPart('word/new.xml', strToU8('<new/>'))
    pkg.replacePart('word/document.xml', replacement)

    expect(pkg.readPart('word/new.xml')).toEqual(strToU8('<new/>'))
    expect(pkg.readPart('WORD/NEW.XML')).toEqual(strToU8('<new/>'))
    expect(pkg.readPart('word/document.xml')).toEqual(replacement)
  })

  it('clears the dirty mark when replacement bytes equal the source', () => {
    const source = sourceOf()
    const pkg = open(source, LIMITS)
    const current = pkg.readPart('word/document.xml')

    pkg.replacePart('word/document.xml', current)
    expect(pkg.save()).toBe(source)
  })

  it('clears the dirty mark when a replacement is reverted', () => {
    const source = sourceOf()
    const pkg = open(source, LIMITS)
    const current = pkg.readPart('word/document.xml')

    pkg.replacePart('word/document.xml', strToU8('changed'))
    pkg.replacePart('word/document.xml', current)
    expect(pkg.save()).toBe(source)
  })

  it('cancels a pending addition when it is removed', () => {
    const source = sourceOf()
    const pkg = open(source, LIMITS)

    pkg.addPart('word/extra.xml', strToU8('<extra/>'))
    pkg.removePart('word/extra.xml')
    expect(pkg.save()).toBe(source)
  })
})

describe('editing and saving', () => {
  it('replaces one part and copies every other local record verbatim', () => {
    const source = sourceOf()
    const pkg = open(source, LIMITS)
    const replacement = strToU8('<w:document>Edited body</w:document>')
    const before = new Map(readZipArchive(source, LIMITS).entries.map(entry => [entry.name, entry]))

    pkg.replacePart('word/document.xml', replacement)
    const output = pkg.save()
    const reopened = open(output, LIMITS)
    const after = new Map(readZipArchive(output, LIMITS).entries.map(entry => [entry.name, entry]))

    expect(reopened.readPart('word/document.xml')).toEqual(replacement)
    for (const entry of after.values()) {
      if (entry.name === 'word/document.xml') continue
      expect(bytesEqual(entry.localRecord, before.get(entry.name)!.localRecord)).toBe(true)
    }
  })

  it('adds parts at the end of the archive', () => {
    const source = sourceOf()
    const pkg = open(source, LIMITS)

    pkg.addPart('word/new.xml', strToU8('<new/>'))
    pkg.addPart('word/stored.bin', Uint8Array.of(7, 8, 9), { method: 0 })
    const output = pkg.save()
    const reopened = open(output, LIMITS)

    expect(reopened.readPart('word/new.xml')).toEqual(strToU8('<new/>'))
    expect(reopened.readPart('word/stored.bin')).toEqual(Uint8Array.of(7, 8, 9))
    expect(entryByName(output, 'word/stored.bin').method).toBe(0)
    const names = reopened.source.entries.map(entry => entry.name)
    expect(names.slice(-2)).toEqual(['word/new.xml', 'word/stored.bin'])
  })

  it('rejects duplicate and case-equivalent added names', () => {
    const pkg = open(sourceOf(), LIMITS)

    pkg.addPart('word/new.xml', strToU8('<new/>'))
    expect(failureOf(() => { pkg.addPart('word/new.xml', strToU8('<again/>')) }).failure.type)
      .toBe('name-collision')
    expect(failureOf(() => { pkg.addPart('WORD/New.XML', strToU8('<again/>')) }).failure.type)
      .toBe('name-collision')
    expect(failureOf(() => { pkg.addPart('WORD/Document.XML', strToU8('<again/>')) }).failure.type)
      .toBe('name-collision')
  })

  it('replaces the bytes of a pending addition', () => {
    const pkg = open(sourceOf(), LIMITS)

    pkg.addPart('word/new.xml', strToU8('<new/>'))
    pkg.replacePart('word/new.xml', strToU8('<updated/>'))

    expect(open(pkg.save(), LIMITS).readPart('word/new.xml')).toEqual(strToU8('<updated/>'))
  })

  it('replaces a part without inflating the source when the caller knows it differs', () => {
    const source = sourceOf()
    const location = centralLocations(source).get('word/document.xml')!
    // A source part that cannot inflate proves the known-different path skips it.
    writeUint32(source, location.localOffset + 22, 1)
    writeUint32(source, location.offset + 24, 1)
    const pkg = open(source, LIMITS)
    const replacement = strToU8('<w:document>Known different</w:document>')

    expect(failureOf(() => pkg.readPart('word/document.xml')).failure.type).toBe('invalid-archive')
    pkg.replacePartKnownDifferent('WORD/DOCUMENT.XML', replacement)

    expect(open(pkg.save(), LIMITS).readPart('word/document.xml')).toEqual(replacement)
  })

  it('replaces a pending addition and rejects unknown parts on the known-different path', () => {
    const pkg = open(sourceOf(), LIMITS)

    pkg.addPart('word/new.xml', strToU8('<new/>'))
    pkg.replacePartKnownDifferent('word/new.xml', strToU8('<updated/>'))
    expect(failureOf(() => { pkg.replacePartKnownDifferent('word/missing.xml', Uint8Array.of(1)) }).failure.type)
      .toBe('unknown-entry')
    const removed = open(sourceOf(), LIMITS)
    removed.removePart('word/document.xml')
    expect(failureOf(() => { removed.replacePartKnownDifferent('word/document.xml', Uint8Array.of(1)) }).failure.type)
      .toBe('unknown-entry')

    expect(open(pkg.save(), LIMITS).readPart('word/new.xml')).toEqual(strToU8('<updated/>'))
  })

  it('removes a part from the saved archive', () => {
    const pkg = open(sourceOf(), LIMITS)

    pkg.removePart('word/media/image1.bin')
    const output = pkg.save()
    const reopened = open(output, LIMITS)

    expect(reopened.source.entries.map(entry => entry.name)).not.toContain('word/media/image1.bin')
    expect(failureOf(() => reopened.readPart('word/media/image1.bin')).failure.type).toBe('unknown-entry')
  })

  it('rejects unknown or repeated removals and replacements', () => {
    const pkg = open(sourceOf(), LIMITS)

    expect(failureOf(() => { pkg.removePart('word/missing.xml') }).failure.type).toBe('unknown-entry')
    expect(failureOf(() => { pkg.replacePart('word/missing.xml', Uint8Array.of(1)) }).failure.type)
      .toBe('unknown-entry')
    pkg.removePart('word/document.xml')
    expect(failureOf(() => { pkg.removePart('word/document.xml') }).failure.type).toBe('unknown-entry')
    expect(failureOf(() => { pkg.replacePart('word/document.xml', Uint8Array.of(1)) }).failure.type)
      .toBe('unknown-entry')
  })

  it('allows re-adding a removed part under the same name', () => {
    const source = sourceOf()
    const pkg = open(source, LIMITS)
    const replacement = strToU8('<w:document>Re-added</w:document>')

    pkg.removePart('word/document.xml')
    pkg.addPart('word/document.xml', replacement)
    const output = pkg.save()
    const reopened = open(output, LIMITS)

    expect(reopened.readPart('word/document.xml')).toEqual(replacement)
    expect(reopened.source.entries.filter(entry => entry.name === 'word/document.xml')).toHaveLength(1)
  })

  it('keeps a stored part stored when it changes', () => {
    const pkg = open(sourceOf(), LIMITS)
    const replacement = Uint8Array.of(11, 12, 13)

    pkg.replacePart('word/media/image1.bin', replacement)
    const output = pkg.save()
    const entry = entryByName(output, 'word/media/image1.bin')

    expect(entry.method).toBe(0)
    expect(entry.compressedSize).toBe(replacement.byteLength)
    expect(Array.from(output.subarray(entry.dataOffset, entry.dataOffset + entry.compressedSize)))
      .toEqual(Array.from(replacement))
  })

  it('preserves the archive comment through a package save', () => {
    const source = withComment(sourceOf(), 'kept archive comment')
    const pkg = open(source, LIMITS)

    pkg.replacePart('word/document.xml', strToU8('changed'))
    const output = pkg.save()

    expect(new TextDecoder().decode(readZipArchive(output, LIMITS).comment))
      .toBe('kept archive comment')
  })
})

describe('inflation bounds', () => {
  it('refuses a package whose declared expansion exceeds the total limit', () => {
    const source = zipSync({
      'word/a.xml': new Uint8Array(600),
      'word/b.xml': new Uint8Array(600),
    })

    const error = failureOf(() => open(source, { maxEntries: 100, maxTotalUncompressedBytes: 1_000 }))

    expect(error.failure.type).toBe('too-large')
    expect(error.message).toContain('uncompressed bytes')
  })

  it('refuses an entry that declares more than the per-entry limit without inflating it', () => {
    const source = sourceOf()
    const location = centralLocations(source).get('word/document.xml')!
    const declared = 65 * 1024 * 1024
    writeUint32(source, location.localOffset + 22, declared)
    writeUint32(source, location.offset + 24, declared)
    const pkg = open(source, {
      maxEntries: 100,
      maxEntryUncompressedBytes: 1_024,
      maxTotalUncompressedBytes: 128 * 1024 * 1024,
    })

    const error = failureOf(() => pkg.readPart('word/document.xml'))

    expect(error.failure.type).toBe('too-large')
    expect(error.message).toContain('per-entry limit')
  })

  it('caps inflation at the declared size when a stream expands past it', () => {
    const source = zipSync({ 'word/bomb.bin': new Uint8Array(1024 * 1024) })
    const location = centralLocations(source).get('word/bomb.bin')!
    writeUint32(source, location.localOffset + 22, 16)
    writeUint32(source, location.offset + 24, 16)
    const pkg = open(source, { maxEntries: 100 })

    const error = failureOf(() => pkg.readPart('word/bomb.bin'))

    expect(error.failure.type).toBe('invalid-archive')
  })
})

describe('part verification', () => {
  it('rejects a part whose CRC disagrees with its bytes', () => {
    const source = sourceOf()
    const location = centralLocations(source).get('word/document.xml')!
    const checksum = readUint32(source, location.localOffset + 14, 'local CRC')
    writeUint32(source, location.localOffset + 14, checksum ^ 0xff)
    writeUint32(source, location.offset + 16, checksum ^ 0xff)
    const pkg = open(source, LIMITS)

    expect(failureOf(() => pkg.readPart('word/document.xml')).failure.type).toBe('crc-mismatch')
  })

  it('rejects a part whose inflated length disagrees with its header', () => {
    const source = sourceOf()
    const location = centralLocations(source).get('word/document.xml')!
    const declared = readUint32(source, location.offset + 24, 'uncompressed size')
    writeUint32(source, location.localOffset + 22, declared + 1)
    writeUint32(source, location.offset + 24, declared + 1)
    const pkg = open(source, LIMITS)

    expect(failureOf(() => pkg.readPart('word/document.xml')).failure.type).toBe('invalid-archive')
  })

  it('reports an inflate failure as an invalid archive', () => {
    const source = sourceOf()
    const location = centralLocations(source).get('word/document.xml')!
    writeUint32(source, location.localOffset + 22, 1)
    writeUint32(source, location.offset + 24, 1)
    const pkg = open(source, LIMITS)
    const error = failureOf(() => pkg.readPart('word/document.xml'))

    expect(error.failure.type).toBe('invalid-archive')
    expect(error.failure.detail).toContain('cannot be inflated')
  })
})

const CONTENT_TYPES_NAMESPACE = 'http://schemas.openxmlformats.org/package/2006/content-types'
const PACKAGE_RELATIONSHIPS_NAMESPACE = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OFFICE_DOCUMENT_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

/** A minimal package whose content types, relationships, and parts agree. */
function opcSource(overrides: Readonly<Record<string, string | Uint8Array>> = {}): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">`
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '</Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0"?>'
      + `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">`
      + `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_RELATIONSHIP}" Target="word/document.xml"/>`
      + '</Relationships>',
    ),
    'word/document.xml': strToU8('<document/>'),
    ...overrides,
  })
}

function graphFailureOf(pkg: ReturnType<typeof open>): ZipError {
  return failureOf(() => {
    assertOpcGraph(pkg)
  })
}

describe('OPC graph validation', () => {
  it('accepts a package whose declarations, targets, and parts agree', () => {
    expect(() => { assertOpcGraph(open(opcSource(), LIMITS)) }).not.toThrow()
  })

  it('accepts absolute, query, fragment, and external relationship targets', () => {
    const source = opcSource({
      '_rels/.rels': strToU8(
        '<?xml version="1.0"?>'
        + `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">`
        + `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_RELATIONSHIP}" Target="/word/document.xml?x=1#part"/>`
        + '<Relationship Id="rId2" Type="http://example.invalid/external" Target="https://example.invalid/x" TargetMode="External"/>'
        + '</Relationships>',
      ),
    })

    expect(() => { assertOpcGraph(open(source, LIMITS)) }).not.toThrow()
  })

  it('matches percent-encoded declarations and targets to literal part names', () => {
    const pkg = open(opcSource({
      '[Content_Types].xml': strToU8(
        `<?xml version="1.0"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">`
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        + '<Override PartName="/word/part%201.xml" ContentType="application/xml"/>'
        + '</Types>',
      ),
      'word/_rels/document.xml.rels': strToU8(
        '<?xml version="1.0"?>'
        + `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">`
        + `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_RELATIONSHIP}" Target="part%201.xml"/>`
        + '</Relationships>',
      ),
      'word/part 1.xml': strToU8('<part/>'),
    }), LIMITS)

    expect(() => { assertOpcGraph(pkg) }).not.toThrow()
  })

  it('rejects a relationship whose internal target is missing', () => {
    const pkg = open(opcSource({
      '_rels/.rels': strToU8(
        '<?xml version="1.0"?>'
        + `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">`
        + `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_RELATIONSHIP}" Target="word/missing.xml"/>`
        + '</Relationships>',
      ),
    }), LIMITS)
    const error = graphFailureOf(pkg)

    expect(error.failure.type).toBe('invalid-archive')
    expect(error.message).toContain('word/missing.xml')
  })

  it('rejects a relative target that escapes the package root', () => {
    const pkg = open(opcSource({
      'word/_rels/document.xml.rels': strToU8(
        '<?xml version="1.0"?>'
        + `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">`
        + `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_RELATIONSHIP}" Target="../../outside.xml"/>`
        + '</Relationships>',
      ),
    }), LIMITS)

    expect(graphFailureOf(pkg).failure.type).toBe('invalid-archive')
  })

  it('rejects a relationship declaration without a target', () => {
    const pkg = open(opcSource({
      '_rels/.rels': strToU8(
        '<?xml version="1.0"?>'
        + `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">`
        + `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_RELATIONSHIP}"/>`
        + '</Relationships>',
      ),
    }), LIMITS)

    expect(graphFailureOf(pkg).message).toContain('no Target attribute')
  })

  it('rejects a part with no content-type declaration', () => {
    const pkg = open(opcSource({ 'word/media/image1.bin': Uint8Array.of(1, 2, 3) }), LIMITS)
    const error = graphFailureOf(pkg)

    expect(error.failure.type).toBe('invalid-archive')
    expect(error.message).toContain('word/media/image1.bin')
  })

  it('rejects a declaration that names a missing part', () => {
    const pkg = open(opcSource({
      '[Content_Types].xml': strToU8(
        `<?xml version="1.0"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">`
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        + '<Override PartName="/word/missing.xml" ContentType="application/xml"/>'
        + '</Types>',
      ),
    }), LIMITS)
    const error = graphFailureOf(pkg)

    expect(error.failure.type).toBe('invalid-archive')
    expect(error.message).toContain('word/missing.xml')
  })

  it('rejects a package without a readable content-types part', () => {
    expect(graphFailureOf(open(zipSync({ 'word/document.xml': strToU8('<document/>') }), LIMITS)).message)
      .toContain('[Content_Types].xml')

    const malformed = open(opcSource({
      '[Content_Types].xml': strToU8('<Types'),
      '_rels/.rels': strToU8(''),
    }), LIMITS)
    expect(graphFailureOf(malformed).failure.type).toBe('invalid-archive')
  })

  it('validates pending additions, removals, and replacements', () => {
    const pkg = open(opcSource(), LIMITS)
    pkg.addPart('word/media/image1.bin', Uint8Array.of(1))
    expect(graphFailureOf(pkg).message).toContain('word/media/image1.bin')

    pkg.removePart('word/media/image1.bin')
    pkg.replacePart('word/document.xml', strToU8('<document edited="true"/>'))
    expect(() => { assertOpcGraph(pkg) }).not.toThrow()

    pkg.removePart('word/document.xml')
    expect(graphFailureOf(pkg).message).toContain('word/document.xml')
  })

  it('lists the current file parts without directory entries', () => {
    const pkg = open(zipSync({
      '[Content_Types].xml': strToU8('<Types/>'),
      '_rels/.rels': strToU8('<Relationships/>'),
      'word/document.xml': strToU8('<document/>'),
      'word/': new Uint8Array(0),
    }), LIMITS)
    pkg.addPart('word/new.xml', strToU8('<new/>'))
    pkg.removePart('word/document.xml')

    expect(pkg.partNames()).toEqual(['[Content_Types].xml', '_rels/.rels', 'word/new.xml'])
  })

  it('propagates a ZIP failure from the content-types part', () => {
    const source = corruptEntryCrc(opcSource(), '[Content_Types].xml')

    expect(graphFailureOf(open(source, LIMITS)).failure.type).toBe('crc-mismatch')
  })

  it('propagates a ZIP failure from a relationships part', () => {
    const source = corruptEntryCrc(opcSource(), '_rels/.rels')

    expect(graphFailureOf(open(source, LIMITS)).failure.type).toBe('crc-mismatch')
  })

  it('rejects a content-types part without a Types element', () => {
    const pkg = open(opcSource({ '[Content_Types].xml': strToU8('<Other/>') }), LIMITS)

    expect(graphFailureOf(pkg).message).toContain('has no Types element')
  })

  it('rejects a Default declaration missing its extension or content type', () => {
    const pkg = open(opcSource({
      '[Content_Types].xml': strToU8(
        `<?xml version="1.0"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">`
        + '<Default ContentType="application/xml"/>'
        + '</Types>',
      ),
    }), LIMITS)

    expect(graphFailureOf(pkg).message).toContain('Default without an extension or content type')
  })

  it('rejects an Override declaration missing its part name or content type', () => {
    const pkg = open(opcSource({
      '[Content_Types].xml': strToU8(
        `<?xml version="1.0"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">`
        + '<Override ContentType="application/xml"/>'
        + '</Types>',
      ),
    }), LIMITS)

    expect(graphFailureOf(pkg).message).toContain('Override without a part name or content type')
  })

  it('accepts declarations with no Default or Override collection at all', () => {
    const noDefaults = open(opcSource({
      '[Content_Types].xml': strToU8(
        `<?xml version="1.0"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">`
        + '<Override PartName="/word/document.xml" ContentType="application/xml"/>'
        + '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '</Types>',
      ),
    }), LIMITS)
    const noOverrides = open(opcSource({
      '[Content_Types].xml': strToU8(
        `<?xml version="1.0"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">`
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '</Types>',
      ),
    }), LIMITS)

    expect(() => { assertOpcGraph(noDefaults) }).not.toThrow()
    expect(() => { assertOpcGraph(noOverrides) }).not.toThrow()
  })

  it('rejects a relationships part without a Relationships element', () => {
    const pkg = open(opcSource({ '_rels/.rels': strToU8('<Other/>') }), LIMITS)

    expect(graphFailureOf(pkg).message).toContain('has no Relationships element')
  })

  it('rejects a relationships part that cannot be parsed', () => {
    const pkg = open(opcSource({ '_rels/.rels': strToU8('<Relationships') }), LIMITS)

    expect(graphFailureOf(pkg).message).toContain('_rels/.rels cannot be parsed')
  })

  it('accepts a Relationships element with no Relationship declarations', () => {
    const pkg = open(zipSync({
      '[Content_Types].xml': strToU8(
        `<?xml version="1.0"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">`
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + '</Types>',
      ),
      '_rels/.rels': strToU8(
        `<?xml version="1.0"?><Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}"><Other/></Relationships>`,
      ),
      'word/document.xml': strToU8('<document/>'),
    }), LIMITS)

    expect(() => { assertOpcGraph(pkg) }).not.toThrow()
  })

  it('accepts a relationship with no Id attribute', () => {
    const pkg = open(opcSource({
      '_rels/.rels': strToU8(
        '<?xml version="1.0"?>'
        + `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">`
        + `<Relationship Type="${OFFICE_DOCUMENT_RELATIONSHIP}" Target="word/document.xml"/>`
        + '</Relationships>',
      ),
    }), LIMITS)

    expect(() => { assertOpcGraph(pkg) }).not.toThrow()
  })

  it('cannot resolve the source part of a relationships file outside _rels', () => {
    const pkg = open(opcSource({
      '[Content_Types].xml': strToU8(
        `<?xml version="1.0"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">`
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + '<Override PartName="/xl/notes.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '</Types>',
      ),
      'xl/notes.rels': strToU8(
        '<?xml version="1.0"?>'
        + `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">`
        + '<Relationship Id="rId1" Type="http://example.invalid/x" Target="missing.xml"/>'
        + '</Relationships>',
      ),
    }), LIMITS)

    expect(() => { assertOpcGraph(pkg) }).not.toThrow()
  })

  it('resolves targets declared by a root-level _rels part', () => {
    const pkg = open(opcSource({
      '_rels/other.rels': strToU8(
        '<?xml version="1.0"?>'
        + `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">`
        + `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_RELATIONSHIP}" Target="word/document.xml"/>`
        + '</Relationships>',
      ),
    }), LIMITS)

    expect(() => { assertOpcGraph(pkg) }).not.toThrow()
  })

  it('keeps a part name with a malformed percent escape literal', () => {
    const pkg = open(opcSource({
      '[Content_Types].xml': strToU8(
        `<?xml version="1.0"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">`
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + '<Override PartName="/word/%ZZ.xml" ContentType="application/xml"/>'
        + '</Types>',
      ),
      'word/%ZZ.xml': strToU8('<part/>'),
    }), LIMITS)

    expect(() => { assertOpcGraph(pkg) }).not.toThrow()
  })

  it('rejects a part with no extension and no Override declaration', () => {
    const pkg = open(opcSource({ 'word/README': strToU8('readme') }), LIMITS)

    expect(graphFailureOf(pkg).message).toContain('word/README')
  })
})
