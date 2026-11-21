/** ZIP writer that copies carried records and rewrites only changed entries. */
import { crc32, deflateRawSync } from 'node:zlib'
import {
  asciiFoldKey,
  assertZip,
  DEFAULT_DOS_DATE_TIME,
  ZIP_CENTRAL_CRC32_OFFSET,
  ZIP_CENTRAL_DIRECTORY_HEADER_SIGNATURE,
  ZIP_CENTRAL_FLAGS_OFFSET,
  ZIP_CENTRAL_COMPRESSED_SIZE_OFFSET,
  ZIP_CENTRAL_HEADER_LENGTH,
  ZIP_CENTRAL_LOCAL_OFFSET_OFFSET,
  ZIP_CENTRAL_UNCOMPRESSED_SIZE_OFFSET,
  ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH,
  ZIP_END_OF_CENTRAL_DIRECTORY_SIGNATURE,
  ZIP_FLAG_DATA_DESCRIPTOR,
  ZIP_FLAG_UTF8,
  ZIP_LOCAL_CRC32_OFFSET,
  ZIP_LOCAL_FILE_HEADER_SIGNATURE,
  ZIP_LOCAL_FLAGS_OFFSET,
  ZIP_LOCAL_COMPRESSED_SIZE_OFFSET,
  ZIP_LOCAL_HEADER_LENGTH,
  ZIP_LOCAL_UNCOMPRESSED_SIZE_OFFSET,
  ZIP_METHOD_DEFLATE,
  ZIP_METHOD_STORE,
  ZIP_UINT16_MAX,
  writeUint16,
  writeUint32,
} from './zip-format.ts'
import type {
  DosDateTime,
  ZipArchive,
  ZipChanges,
  ZipEntryRecord,
  ZipMethod,
  ZipPartAddition,
  ZipWriteOptions,
} from './types.ts'

/** One entry prepared for both sections of the output archive. */
interface OutputEntry {
  readonly name: string
  readonly local: Uint8Array
  readonly central: Uint8Array
  readonly localOffset: number
}

/** A rewritten local record and the values its central record must carry. */
interface RewrittenRecord {
  readonly local: Uint8Array
  readonly crc32: number
  readonly compressedSize: number
  readonly uncompressedSize: number
}

/** A newly built local record plus header fields shared with its central record. */
interface BuiltAddition {
  readonly local: Uint8Array
  readonly nameBytes: Uint8Array
  readonly flags: number
  readonly method: ZipMethod
  readonly crc32: number
  readonly compressedSize: number
  readonly uncompressedSize: number
  readonly dos: DosDateTime
}

/** Shared UTF-8 encoder for entry names. */
const TEXT_ENCODER = new TextEncoder()

/**
 * Write an archive with the requested entry-level changes.
 *
 * With no changes the source bytes are returned unchanged. Otherwise every
 * carried local record is copied verbatim, changed entries are rebuilt from
 * their original local header, and the central directory is re-emitted with
 * updated offsets.
 *
 * @param archive - Source archive indexed by the reader.
 * @param changes - Replacements, additions, and removals to apply.
 * @param options - Determinism inputs for records this writer creates.
 * @returns The complete output archive.
 * @throws {ZipError} When a change names an unknown entry or a colliding name.
 */
export function writeZipArchive(
  archive: ZipArchive,
  changes: ZipChanges,
  options: ZipWriteOptions = {},
): Uint8Array {
  const replacements = changes.replacements ?? []
  const additions = changes.additions ?? []
  const removals = changes.removals ?? []
  if (replacements.length === 0 && additions.length === 0 && removals.length === 0) return archive.bytes

  const entryByName = new Map<string, ZipEntryRecord>()
  for (const entry of archive.centralOrder) entryByName.set(entry.name, entry)

  const removalSet = new Set<string>()
  for (const name of removals) {
    assertZip(entryByName.has(name), {
      type: 'unknown-entry',
      detail: `Cannot remove ${name}: the source archive has no such entry`,
    })
    removalSet.add(name)
  }

  const replacementByName = new Map<string, Uint8Array>()
  for (const replacement of replacements) {
    const entry = entryByName.get(replacement.name)
    assertZip(entry !== undefined, {
      type: 'unknown-entry',
      detail: `Cannot replace ${replacement.name}: the source archive has no such entry`,
    })
    assertZip(!removalSet.has(replacement.name), {
      type: 'unknown-entry',
      detail: `Cannot replace ${replacement.name}: the entry is also removed`,
    })
    assertZip(!replacementByName.has(replacement.name), {
      type: 'duplicate-name',
      detail: `Entry ${replacement.name} is replaced more than once`,
    })
    replacementByName.set(replacement.name, replacement.bytes)
  }

  const occupiedNames = new Set<string>()
  for (const entry of archive.centralOrder) {
    if (!removalSet.has(entry.name)) occupiedNames.add(asciiFoldKey(entry.nameBytes))
  }
  const additionNameBytes = new Map<string, Uint8Array>()
  for (const addition of additions) {
    const nameBytes = TEXT_ENCODER.encode(addition.name)
    const foldedName = asciiFoldKey(nameBytes)
    assertZip(!occupiedNames.has(foldedName) && !additionNameBytes.has(foldedName), {
      type: 'name-collision',
      detail: `Cannot add ${addition.name}: the name collides under ASCII case folding`,
    })
    occupiedNames.add(foldedName)
    additionNameBytes.set(addition.name, nameBytes)
  }

  const outputByName = new Map<string, OutputEntry>()
  const localOrder: OutputEntry[] = []
  let localLength = 0
  for (const entry of archive.localOrder) {
    if (removalSet.has(entry.name)) continue
    const replacement = replacementByName.get(entry.name)
    if (replacement === undefined) {
      const output: OutputEntry = {
        name: entry.name,
        local: entry.localRecord,
        central: carryCentralRecord(entry, localLength),
        localOffset: localLength,
      }
      localOrder.push(output)
      outputByName.set(entry.name, output)
      localLength += entry.localRecord.byteLength
      continue
    }
    const rewritten = rewriteLocalRecord(entry, replacement, options.deflateLevel)
    const output: OutputEntry = {
      name: entry.name,
      local: rewritten.local,
      central: rewriteCentralRecord(entry, localLength, rewritten),
      localOffset: localLength,
    }
    localOrder.push(output)
    outputByName.set(entry.name, output)
    localLength += rewritten.local.byteLength
  }

  for (const addition of additions) {
    const nameBytes = additionNameBytes.get(addition.name)
    assertZip(nameBytes !== undefined, {
      type: 'invalid-archive',
      detail: `Added entry ${addition.name} lost its encoded name`,
    })
    const built = buildLocalAddition(addition, nameBytes, options)
    const output: OutputEntry = {
      name: addition.name,
      local: built.local,
      central: buildCentralAddition(built, localLength),
      localOffset: localLength,
    }
    localOrder.push(output)
    outputByName.set(addition.name, output)
    localLength += built.local.byteLength
  }

  const centralOrder: OutputEntry[] = []
  for (const entry of archive.centralOrder) {
    if (removalSet.has(entry.name)) continue
    const output = outputByName.get(entry.name)
    assertZip(output !== undefined, {
      type: 'invalid-archive',
      detail: `Carried entry ${entry.name} has no prepared output record`,
    })
    centralOrder.push(output)
  }
  for (const addition of additions) {
    const output = outputByName.get(addition.name)
    assertZip(output !== undefined, {
      type: 'invalid-archive',
      detail: `Added entry ${addition.name} has no prepared output record`,
    })
    centralOrder.push(output)
  }

  assertZip(centralOrder.length < ZIP_UINT16_MAX, {
    type: 'entry-limit',
    detail: `The output would contain ${centralOrder.length} entries, which needs ZIP64`,
  })

  const centralOffset = localLength
  let outputLength = centralOffset
  for (const output of centralOrder) outputLength += output.central.byteLength
  outputLength += ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH + archive.comment.byteLength
  const bytes = new Uint8Array(outputLength)

  let cursor = 0
  for (const output of localOrder) {
    bytes.set(output.local, cursor)
    cursor += output.local.byteLength
  }
  for (const output of centralOrder) {
    bytes.set(output.central, cursor)
    cursor += output.central.byteLength
  }

  writeUint32(bytes, cursor, ZIP_END_OF_CENTRAL_DIRECTORY_SIGNATURE)
  writeUint16(bytes, cursor + 4, 0)
  writeUint16(bytes, cursor + 6, 0)
  writeUint16(bytes, cursor + 8, centralOrder.length)
  writeUint16(bytes, cursor + 10, centralOrder.length)
  writeUint32(bytes, cursor + 12, cursor - centralOffset)
  writeUint32(bytes, cursor + 16, centralOffset)
  writeUint16(bytes, cursor + 20, archive.comment.byteLength)
  bytes.set(archive.comment, cursor + ZIP_END_OF_CENTRAL_DIRECTORY_LENGTH)
  return bytes
}

/**
 * Copy a carried central record and patch its local offset.
 * @param entry - Source entry.
 * @param localOffset - Offset of the copied local record in the output.
 * @returns The patched central record bytes.
 */
function carryCentralRecord(entry: ZipEntryRecord, localOffset: number): Uint8Array {
  const central = entry.centralRecord.slice()
  writeUint32(central, ZIP_CENTRAL_LOCAL_OFFSET_OFFSET, localOffset)
  return central
}

/**
 * Rebuild a changed entry's local record from its original header.
 * @param entry - Source entry whose header, name, and extra fields are reused.
 * @param replacement - Uncompressed replacement bytes.
 * @param deflateLevel - Optional deflate level.
 * @returns The new local record and data values.
 */
function rewriteLocalRecord(
  entry: ZipEntryRecord,
  replacement: Uint8Array,
  deflateLevel: number | undefined,
): RewrittenRecord {
  const compressed = compress(replacement, entry.method, deflateLevel)
  const checksum = crc32(replacement)
  const local = new Uint8Array(entry.localHeaderLength + compressed.byteLength)
  local.set(entry.localRecord.subarray(0, entry.localHeaderLength))
  writeUint16(local, ZIP_LOCAL_FLAGS_OFFSET, entry.flags & ~ZIP_FLAG_DATA_DESCRIPTOR)
  writeUint32(local, ZIP_LOCAL_CRC32_OFFSET, checksum)
  writeUint32(local, ZIP_LOCAL_COMPRESSED_SIZE_OFFSET, compressed.byteLength)
  writeUint32(local, ZIP_LOCAL_UNCOMPRESSED_SIZE_OFFSET, replacement.byteLength)
  local.set(compressed, entry.localHeaderLength)
  return {
    local,
    crc32: checksum,
    compressedSize: compressed.byteLength,
    uncompressedSize: replacement.byteLength,
  }
}

/**
 * Patch a changed entry's central record.
 * @param entry - Source entry whose central record is reused.
 * @param localOffset - Offset of the rewritten local record.
 * @param rewritten - New CRC and size values.
 * @returns The patched central record bytes.
 */
function rewriteCentralRecord(
  entry: ZipEntryRecord,
  localOffset: number,
  rewritten: RewrittenRecord,
): Uint8Array {
  const central = entry.centralRecord.slice()
  writeUint16(central, ZIP_CENTRAL_FLAGS_OFFSET, entry.flags & ~ZIP_FLAG_DATA_DESCRIPTOR)
  writeUint32(central, ZIP_CENTRAL_CRC32_OFFSET, rewritten.crc32)
  writeUint32(central, ZIP_CENTRAL_COMPRESSED_SIZE_OFFSET, rewritten.compressedSize)
  writeUint32(central, ZIP_CENTRAL_UNCOMPRESSED_SIZE_OFFSET, rewritten.uncompressedSize)
  writeUint32(central, ZIP_CENTRAL_LOCAL_OFFSET_OFFSET, localOffset)
  return central
}

/**
 * Build a new entry's local file header and compressed data.
 * @param addition - Caller-provided entry.
 * @param nameBytes - Encoded entry name.
 * @param options - Deflate level and DOS timestamp inputs.
 * @returns The local record and header fields needed by the central record.
 */
function buildLocalAddition(
  addition: ZipPartAddition,
  nameBytes: Uint8Array,
  options: ZipWriteOptions,
): BuiltAddition {
  const method = addition.method ?? ZIP_METHOD_DEFLATE
  const compressed = compress(addition.bytes, method, options.deflateLevel)
  const checksum = crc32(addition.bytes)
  const flags = hasNonAscii(nameBytes) ? ZIP_FLAG_UTF8 : 0
  const dos = options.addedEntryDosDateTime ?? DEFAULT_DOS_DATE_TIME
  const local = new Uint8Array(ZIP_LOCAL_HEADER_LENGTH + nameBytes.byteLength + compressed.byteLength)
  writeUint32(local, 0, ZIP_LOCAL_FILE_HEADER_SIGNATURE)
  writeUint16(local, 4, 20)
  writeUint16(local, ZIP_LOCAL_FLAGS_OFFSET, flags)
  writeUint16(local, 8, method)
  writeUint16(local, 10, dos.time)
  writeUint16(local, 12, dos.date)
  writeUint32(local, ZIP_LOCAL_CRC32_OFFSET, checksum)
  writeUint32(local, ZIP_LOCAL_COMPRESSED_SIZE_OFFSET, compressed.byteLength)
  writeUint32(local, ZIP_LOCAL_UNCOMPRESSED_SIZE_OFFSET, addition.bytes.byteLength)
  writeUint16(local, 26, nameBytes.byteLength)
  writeUint16(local, 28, 0)
  local.set(nameBytes, ZIP_LOCAL_HEADER_LENGTH)
  local.set(compressed, ZIP_LOCAL_HEADER_LENGTH + nameBytes.byteLength)
  return {
    local,
    nameBytes,
    flags,
    method,
    crc32: checksum,
    compressedSize: compressed.byteLength,
    uncompressedSize: addition.bytes.byteLength,
    dos,
  }
}

/**
 * Build a new entry's central directory record.
 * @param built - Local record fields.
 * @param localOffset - Offset of the new local record.
 * @returns The central record bytes.
 */
function buildCentralAddition(built: BuiltAddition, localOffset: number): Uint8Array {
  const central = new Uint8Array(ZIP_CENTRAL_HEADER_LENGTH + built.nameBytes.byteLength)
  writeUint32(central, 0, ZIP_CENTRAL_DIRECTORY_HEADER_SIGNATURE)
  writeUint16(central, 4, 20)
  writeUint16(central, 6, 20)
  writeUint16(central, ZIP_CENTRAL_FLAGS_OFFSET, built.flags)
  writeUint16(central, 10, built.method)
  writeUint16(central, 12, built.dos.time)
  writeUint16(central, 14, built.dos.date)
  writeUint32(central, ZIP_CENTRAL_CRC32_OFFSET, built.crc32)
  writeUint32(central, ZIP_CENTRAL_COMPRESSED_SIZE_OFFSET, built.compressedSize)
  writeUint32(central, ZIP_CENTRAL_UNCOMPRESSED_SIZE_OFFSET, built.uncompressedSize)
  writeUint16(central, 28, built.nameBytes.byteLength)
  writeUint16(central, 30, 0)
  writeUint16(central, 32, 0)
  writeUint16(central, 34, 0)
  writeUint16(central, 36, 0)
  writeUint32(central, 38, 0)
  writeUint32(central, ZIP_CENTRAL_LOCAL_OFFSET_OFFSET, localOffset)
  central.set(built.nameBytes, ZIP_CENTRAL_HEADER_LENGTH)
  return central
}

/**
 * Compress entry bytes with the entry's method.
 * @param bytes - Uncompressed bytes.
 * @param method - Store or deflate.
 * @param deflateLevel - Optional deflate level.
 * @returns The bytes that belong after the local header.
 */
function compress(bytes: Uint8Array, method: ZipMethod, deflateLevel: number | undefined): Uint8Array {
  if (method === ZIP_METHOD_STORE) return bytes
  if (deflateLevel === undefined) return deflateRawSync(bytes)
  return deflateRawSync(bytes, { level: deflateLevel })
}

/**
 * Whether an encoded name needs the UTF-8 flag.
 * @param nameBytes - Encoded entry name.
 * @returns `true` when any byte is outside ASCII.
 */
function hasNonAscii(nameBytes: Uint8Array): boolean {
  for (const value of nameBytes) {
    if (value > 0x7f) return true
  }
  return false
}
