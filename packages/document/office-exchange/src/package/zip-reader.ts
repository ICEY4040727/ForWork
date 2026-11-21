/** Bounded ZIP reader that indexes raw local and central records without decoding them. */
import {
  asciiFoldKey,
  assertRange,
  assertZip,
  bytesEqual,
  CFB_SIGNATURE,
  failZip,
  hasSignature,
  isZipMethod,
  readUint16,
  readUint32,
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
  ZIP_FLAG_MASKED_HEADERS,
  ZIP_FLAG_STRONG_ENCRYPTION,
  ZIP_FLAG_UTF8,
  ZIP_LOCAL_FILE_HEADER_SIGNATURE,
  ZIP_LOCAL_HEADER_LENGTH,
  ZIP_METHOD_STORE,
  ZIP_UINT16_MAX,
  ZIP_UINT32_MAX,
} from './zip-format.ts'
import type { ZipArchive, ZipEntryRecord, ZipMethod, ZipReadLimits } from './types.ts'

/** ZIP64 locator length immediately before an EOCD that has one. */
const ZIP64_LOCATOR_LENGTH = 20

/** Every general-purpose bit that signals encryption. */
const ENCRYPTION_FLAGS = ZIP_FLAG_ENCRYPTED | ZIP_FLAG_STRONG_ENCRYPTION | ZIP_FLAG_MASKED_HEADERS

/** End of central directory fields needed to locate the entry table. */
interface EndOfCentralDirectory {
  readonly offset: number
  readonly diskNumber: number
  readonly centralDisk: number
  readonly entriesOnDisk: number
  readonly entryCount: number
  readonly centralSize: number
  readonly centralOffset: number
  readonly commentLength: number
}

/** Central-directory fields that both headers must agree on. */
interface CentralRecord {
  readonly name: string
  readonly nameBytes: Uint8Array
  readonly method: ZipMethod
  readonly flags: number
  readonly crc32: number
  readonly compressedSize: number
  readonly uncompressedSize: number
  readonly localOffset: number
  readonly centralRecord: Uint8Array
}

/**
 * Index a ZIP archive that the in-place writer can preserve.
 * @param bytes - Complete archive bytes.
 * @param limits - Entry-count and uncompressed-byte bounds applied before records are parsed.
 * @returns Records in central order plus the local-order projection and comment.
 * @throws {ZipError} When the container is unsupported, inconsistent, or over a bound.
 */
export function readZipArchive(bytes: Uint8Array, limits: ZipReadLimits): ZipArchive {
  assertZip(!hasSignature(bytes, 0, CFB_SIGNATURE), {
    type: 'unsupported-container',
    container: 'cfb',
    detail: 'The file is a Compound File Binary (OLE) container, not a ZIP package',
  })

  const eocd = findEndOfCentralDirectory(bytes)
  assertZip(
    eocd.diskNumber === 0 && eocd.centralDisk === 0 && eocd.entriesOnDisk === eocd.entryCount,
    {
      type: 'unsupported-container',
      container: 'multi-disk',
      detail: 'Split and multi-disk ZIP archives are not supported',
    },
  )

  const zip64LocatorOffset = eocd.offset - ZIP64_LOCATOR_LENGTH
  assertZip(
    zip64LocatorOffset < 0
      || readUint32(bytes, zip64LocatorOffset, 'ZIP64 locator') !== ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR_SIGNATURE,
    {
      type: 'unsupported-container',
      container: 'zip64',
      detail: 'ZIP64 archives are not supported',
    },
  )
  assertZip(
    eocd.entriesOnDisk !== ZIP_UINT16_MAX
      && eocd.entryCount !== ZIP_UINT16_MAX
      && eocd.centralSize !== ZIP_UINT32_MAX
      && eocd.centralOffset !== ZIP_UINT32_MAX,
    {
      type: 'unsupported-container',
      container: 'zip64',
      detail: 'ZIP64 sentinel fields are not supported',
    },
  )

  const centralEnd = eocd.centralOffset + eocd.centralSize
  assertZip(centralEnd <= eocd.offset, {
    type: 'out-of-bounds',
    detail: `Central directory [${eocd.centralOffset}, ${centralEnd}) overlaps the end record at ${eocd.offset}`,
  })
  assertZip(eocd.entryCount <= limits.maxEntries, {
    type: 'entry-limit',
    detail: `Archive contains ${eocd.entryCount} entries; the limit is ${limits.maxEntries}`,
  })

  const centralRecords = readCentralDirectory(bytes, eocd)
  assertDeclaredExpansion(centralRecords, limits)
  const entries = attachLocalRecords(bytes, centralRecords, eocd)
  const localOrder = [...entries].sort((left, right) => left.localOffset - right.localOffset)
  let previousEnd = 0
  for (const entry of localOrder) {
    assertZip(entry.localOffset >= previousEnd, {
      type: 'invalid-archive',
      detail: `Entry ${entry.name} at ${entry.localOffset} overlaps the previous local record ending at ${previousEnd}`,
    })
    previousEnd = entry.localOffset + entry.localRecord.byteLength
  }

  return {
    bytes,
    entries,
    localOrder,
    centralOrder: entries,
    comment: bytes.slice(eocd.offset + ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH),
  }
}

/**
 * Locate an EOCD whose comment length reaches the end of the file.
 * @param bytes - Complete archive bytes.
 * @returns The parsed EOCD fields.
 */
function findEndOfCentralDirectory(bytes: Uint8Array): EndOfCentralDirectory {
  if (bytes.byteLength < ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH) {
    failZip({
      type: 'invalid-archive',
      detail: `The ${bytes.byteLength}-byte file is shorter than an end of central directory record`,
    })
  }

  for (let offset = bytes.byteLength - ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH; offset >= 0; offset -= 1) {
    if (readUint32(bytes, offset, 'end of central directory signature') !== ZIP_END_OF_CENTRAL_DIRECTORY_SIGNATURE) {
      continue
    }
    const commentLength = readUint16(bytes, offset + 20, 'end of central directory comment length')
    if (offset + ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH + commentLength !== bytes.byteLength) {
      continue
    }
    return {
      offset,
      diskNumber: readUint16(bytes, offset + 4, 'end of central directory disk number'),
      centralDisk: readUint16(bytes, offset + 6, 'end of central directory central disk'),
      entriesOnDisk: readUint16(bytes, offset + 8, 'end of central directory entries on disk'),
      entryCount: readUint16(bytes, offset + 10, 'end of central directory entry count'),
      centralSize: readUint32(bytes, offset + 12, 'end of central directory size'),
      centralOffset: readUint32(bytes, offset + 16, 'end of central directory offset'),
      commentLength,
    }
  }

  return failZip({
    type: 'invalid-archive',
    detail: 'No end of central directory record with a consistent comment length was found',
  })
}

/**
 * Parse every central-directory record in order.
 * @param bytes - Complete archive bytes.
 * @param eocd - Located end record.
 * @returns Central records with their raw views.
 */
function readCentralDirectory(bytes: Uint8Array, eocd: EndOfCentralDirectory): readonly CentralRecord[] {
  const centralEnd = eocd.centralOffset + eocd.centralSize
  const records: CentralRecord[] = []
  const seenNames = new Set<string>()
  let cursor = eocd.centralOffset

  for (let index = 0; index < eocd.entryCount; index += 1) {
    assertRange(bytes, cursor, ZIP_CENTRAL_HEADER_LENGTH, 'central directory header')
    assertZip(readUint32(bytes, cursor, 'central directory signature') === ZIP_CENTRAL_DIRECTORY_HEADER_SIGNATURE, {
      type: 'invalid-archive',
      detail: `Central directory record ${index} at ${cursor} has the wrong signature`,
    })

    const flags = readUint16(bytes, cursor + 8, 'central directory flags')
    const method = readUint16(bytes, cursor + 10, 'central directory method')
    const crc32 = readUint32(bytes, cursor + 16, 'central directory CRC-32')
    const compressedSize = readUint32(bytes, cursor + 20, 'central directory compressed size')
    const uncompressedSize = readUint32(bytes, cursor + 24, 'central directory uncompressed size')
    const nameLength = readUint16(bytes, cursor + 28, 'central directory name length')
    const extraLength = readUint16(bytes, cursor + 30, 'central directory extra length')
    const commentLength = readUint16(bytes, cursor + 32, 'central directory comment length')
    const diskStart = readUint16(bytes, cursor + 34, 'central directory disk start')
    const localOffset = readUint32(bytes, cursor + 42, 'central directory local offset')
    const recordLength = ZIP_CENTRAL_HEADER_LENGTH + nameLength + extraLength + commentLength
    assertRange(bytes, cursor, recordLength, 'central directory record')

    assertZip(diskStart === 0, {
      type: 'unsupported-container',
      container: 'multi-disk',
      detail: `Entry at central directory offset ${cursor} starts on disk ${diskStart}`,
    })
    assertZip(isZipMethod(method), {
      type: 'unsupported-container',
      container: 'unsupported-method',
      detail: `Entry at central directory offset ${cursor} uses compression method ${method}`,
    })
    assertZip((flags & ENCRYPTION_FLAGS) === 0, {
      type: 'unsupported-container',
      container: 'encrypted',
      detail: `Entry at central directory offset ${cursor} is encrypted`,
    })
    assertZip(
      compressedSize !== ZIP_UINT32_MAX && uncompressedSize !== ZIP_UINT32_MAX && localOffset !== ZIP_UINT32_MAX,
      {
        type: 'unsupported-container',
        container: 'zip64',
        detail: `Entry at central directory offset ${cursor} uses ZIP64 sizes`,
      },
    )

    const nameBytes = bytes.subarray(cursor + ZIP_CENTRAL_HEADER_LENGTH, cursor + ZIP_CENTRAL_HEADER_LENGTH + nameLength)
    const name = decodeEntryName(nameBytes, flags)
    const foldedName = asciiFoldKey(nameBytes)
    assertZip(!seenNames.has(foldedName), {
      type: 'duplicate-name',
      detail: `Entry name ${name} duplicates an earlier name under ASCII case folding`,
    })
    seenNames.add(foldedName)

    records.push({
      name,
      nameBytes,
      method,
      flags,
      crc32,
      compressedSize,
      uncompressedSize,
      localOffset,
      centralRecord: bytes.subarray(cursor, cursor + recordLength),
    })
    cursor += recordLength
  }

  assertZip(cursor === centralEnd, {
    type: 'invalid-archive',
    detail: `Central directory ended at ${cursor}, but its header declares the end at ${centralEnd}`,
  })
  return records
}

/**
 * Refuse an archive whose declared uncompressed sizes exceed the total bound.
 *
 * The check runs before any entry is inflated, so a container that declares an
 * expansion bomb is refused without allocating its declared size.
 * @param records - Central records with their declared sizes.
 * @param limits - Bounds whose total limit may be omitted.
 * @throws {ZipError} `too-large` when the declared total exceeds the limit.
 */
function assertDeclaredExpansion(records: readonly CentralRecord[], limits: ZipReadLimits): void {
  if (limits.maxTotalUncompressedBytes === undefined) return
  let total = 0
  for (const record of records) total += record.uncompressedSize
  assertZip(total <= limits.maxTotalUncompressedBytes, {
    type: 'too-large',
    detail: `Archive declares ${total} uncompressed bytes; the limit is ${limits.maxTotalUncompressedBytes}`,
  })
}

/**
 * Decode an entry name using the header's UTF-8 flag.
 * @param nameBytes - Raw name bytes.
 * @param flags - General-purpose flags from the same header.
 * @returns The decoded entry name.
 */
function decodeEntryName(nameBytes: Uint8Array, flags: number): string {
  const decoder = new TextDecoder('utf-8', { fatal: (flags & ZIP_FLAG_UTF8) !== 0 })
  try {
    return decoder.decode(nameBytes)
  } catch (cause) {
    return failZip({
      type: 'invalid-archive',
      detail: `Entry name is not valid UTF-8: ${String(cause)}`,
    })
  }
}

/**
 * Attach each central record's local header, data range, and descriptor.
 * @param bytes - Complete archive bytes.
 * @param centralRecords - Records parsed from the central directory.
 * @param eocd - Located end record bounding the local section.
 * @returns Complete entry records in central order.
 */
function attachLocalRecords(
  bytes: Uint8Array,
  centralRecords: readonly CentralRecord[],
  eocd: EndOfCentralDirectory,
): readonly ZipEntryRecord[] {
  const entries: ZipEntryRecord[] = []
  for (const record of centralRecords) {
    const localOffset = record.localOffset
    assertZip(localOffset < eocd.centralOffset, {
      type: 'out-of-bounds',
      detail: `Entry ${record.name} has local offset ${localOffset} inside the central directory`,
    })
    assertRange(bytes, localOffset, ZIP_LOCAL_HEADER_LENGTH, 'local file header')
    assertZip(readUint32(bytes, localOffset, 'local file header signature') === ZIP_LOCAL_FILE_HEADER_SIGNATURE, {
      type: 'invalid-archive',
      detail: `Entry ${record.name} at ${localOffset} has the wrong local signature`,
    })

    const flags = readUint16(bytes, localOffset + 6, 'local file header flags')
    const method = readUint16(bytes, localOffset + 8, 'local file header method')
    const crc32 = readUint32(bytes, localOffset + 14, 'local file header CRC-32')
    const compressedSize = readUint32(bytes, localOffset + 18, 'local file header compressed size')
    const uncompressedSize = readUint32(bytes, localOffset + 22, 'local file header uncompressed size')
    const nameLength = readUint16(bytes, localOffset + 26, 'local file header name length')
    const extraLength = readUint16(bytes, localOffset + 28, 'local file header extra length')
    const localHeaderLength = ZIP_LOCAL_HEADER_LENGTH + nameLength + extraLength
    assertRange(bytes, localOffset, localHeaderLength, 'local file header name and extra fields')

    assertZip(method === record.method, {
      type: 'invalid-archive',
      detail: `Entry ${record.name} stores method ${method} locally and ${record.method} centrally`,
    })
    assertZip(flags === record.flags, {
      type: 'invalid-archive',
      detail: `Entry ${record.name} stores flags ${flags} locally and ${record.flags} centrally`,
    })
    const localName = bytes.subarray(localOffset + ZIP_LOCAL_HEADER_LENGTH, localOffset + ZIP_LOCAL_HEADER_LENGTH + nameLength)
    assertZip(bytesEqual(localName, record.nameBytes), {
      type: 'invalid-archive',
      detail: `Entry ${record.name} has a different local name`,
    })

    const dataEnd = localOffset + localHeaderLength + record.compressedSize
    assertZip(dataEnd <= eocd.centralOffset, {
      type: 'out-of-bounds',
      detail: `Entry ${record.name} data ends at ${dataEnd}, beyond the central directory at ${eocd.centralOffset}`,
    })

    let descriptorLength = 0
    if ((record.flags & ZIP_FLAG_DATA_DESCRIPTOR) !== 0) {
      descriptorLength = readDescriptorLength(bytes, dataEnd, eocd.centralOffset, record)
    } else {
      assertZip(
        crc32 === record.crc32
          && compressedSize === record.compressedSize
          && uncompressedSize === record.uncompressedSize,
        {
          type: 'invalid-archive',
          detail: `Entry ${record.name} local sizes or CRC disagree with the central directory`,
        },
      )
    }

    assertZip(record.method !== ZIP_METHOD_STORE || record.compressedSize === record.uncompressedSize, {
      type: 'invalid-archive',
      detail: `Stored entry ${record.name} has different compressed and uncompressed sizes`,
    })

    entries.push({
      name: record.name,
      nameBytes: record.nameBytes,
      method: record.method,
      flags: record.flags,
      crc32: record.crc32,
      compressedSize: record.compressedSize,
      uncompressedSize: record.uncompressedSize,
      localOffset,
      localHeaderLength,
      dataOffset: localOffset + localHeaderLength,
      descriptorLength,
      localRecord: bytes.subarray(localOffset, dataEnd + descriptorLength),
      centralRecord: record.centralRecord,
    })
  }
  return entries
}

/**
 * Measure and verify a data descriptor that follows compressed data.
 * @param bytes - Complete archive bytes.
 * @param dataEnd - Offset immediately after the compressed data.
 * @param centralOffset - Start of the central directory, bounding the record.
 * @param record - Central record whose CRC and sizes the descriptor must repeat.
 * @returns Descriptor length in bytes, including an optional signature.
 */
function readDescriptorLength(
  bytes: Uint8Array,
  dataEnd: number,
  centralOffset: number,
  record: CentralRecord,
): number {
  assertZip(dataEnd + ZIP_DATA_DESCRIPTOR_UNSIGNED_LENGTH <= centralOffset, {
    type: 'out-of-bounds',
    detail: `Entry ${record.name} has no room for a data descriptor after ${dataEnd}`,
  })
  const first = readUint32(bytes, dataEnd, 'data descriptor field')
  if (first === ZIP_DATA_DESCRIPTOR_SIGNATURE) {
    assertZip(record.crc32 !== ZIP_DATA_DESCRIPTOR_SIGNATURE, {
      type: 'unsupported-container',
      container: 'ambiguous-descriptor',
      detail: `Entry ${record.name} has a descriptor whose first word is both the signature and the CRC`,
    })
    assertZip(dataEnd + ZIP_DATA_DESCRIPTOR_LENGTH <= centralOffset, {
      type: 'out-of-bounds',
      detail: `Entry ${record.name} has a truncated signed data descriptor after ${dataEnd}`,
    })
    assertZip(
      readUint32(bytes, dataEnd + 4, 'data descriptor CRC-32') === record.crc32
        && readUint32(bytes, dataEnd + 8, 'data descriptor compressed size') === record.compressedSize
        && readUint32(bytes, dataEnd + 12, 'data descriptor uncompressed size') === record.uncompressedSize,
      {
        type: 'invalid-archive',
        detail: `Entry ${record.name} signed data descriptor disagrees with the central directory`,
      },
    )
    return ZIP_DATA_DESCRIPTOR_LENGTH
  }

  assertZip(
    first === record.crc32
      && readUint32(bytes, dataEnd + 4, 'data descriptor compressed size') === record.compressedSize
      && readUint32(bytes, dataEnd + 8, 'data descriptor uncompressed size') === record.uncompressedSize,
    {
      type: 'invalid-archive',
      detail: `Entry ${record.name} data descriptor disagrees with the central directory`,
    },
  )
  return ZIP_DATA_DESCRIPTOR_UNSIGNED_LENGTH
}
