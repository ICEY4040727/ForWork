/** Shared types for the in-place OOXML package layer. */

/** Compression methods this layer reads and writes. */
export type ZipMethod = 0 | 8

/** A DOS date and time pair stored in a ZIP header as two 16-bit fields. */
export interface DosDateTime {
  /** DOS time field: hours, minutes, and two-second units. */
  readonly time: number
  /** DOS date field: year offset, month, and day. */
  readonly date: number
}

/**
 * One central-directory entry with raw views into the source bytes.
 *
 * `localRecord` and `centralRecord` are contiguous views of the source
 * archive. `localRecord` starts at `localOffset` and ends at the descriptor
 * end, so a carried entry can be copied without decoding it.
 */
export interface ZipEntryRecord {
  /** Decoded entry name as stored in the central directory. */
  readonly name: string
  /** Raw name bytes, used for exact comparisons and ASCII case folding. */
  readonly nameBytes: Uint8Array
  /** Compression method recorded in both headers. */
  readonly method: ZipMethod
  /** General-purpose flag bits recorded in both headers. */
  readonly flags: number
  /** CRC-32 of the uncompressed bytes. */
  readonly crc32: number
  /** Compressed byte length from the central directory. */
  readonly compressedSize: number
  /** Uncompressed byte length from the central directory. */
  readonly uncompressedSize: number
  /** Offset of the local file header in the source archive. */
  readonly localOffset: number
  /** Length of the local header, including its name and extra fields. */
  readonly localHeaderLength: number
  /** Offset of the compressed data in the source archive. */
  readonly dataOffset: number
  /** Length of the optional data descriptor, or zero when bit 3 is clear. */
  readonly descriptorLength: number
  /** Contiguous source view from `localOffset` through the descriptor end. */
  readonly localRecord: Uint8Array
  /** Contiguous source view of the central-directory record. */
  readonly centralRecord: Uint8Array
}

/** An indexed ZIP archive whose records are views into the source bytes. */
export interface ZipArchive {
  /** The complete source archive. */
  readonly bytes: Uint8Array
  /** Records in central-directory order. */
  readonly entries: readonly ZipEntryRecord[]
  /** The same records ordered by local header offset. */
  readonly localOrder: readonly ZipEntryRecord[]
  /** The same records in the order their central records appear. */
  readonly centralOrder: readonly ZipEntryRecord[]
  /** Archive comment bytes, preserved verbatim. */
  readonly comment: Uint8Array
}

/** The package as opened, including its content digest. */
export interface SourcePackage extends ZipArchive {
  /** SHA-256 hex digest of `bytes`. */
  readonly digest: string
}

/** Replacement bytes for an existing entry. */
export interface ZipPartReplacement {
  /** Exact entry name from the central directory. */
  readonly name: string
  /** Uncompressed replacement bytes. */
  readonly bytes: Uint8Array
}

/** A new entry appended after the source entries. */
export interface ZipPartAddition {
  /** New entry name; must not collide with a surviving entry. */
  readonly name: string
  /** Uncompressed entry bytes. */
  readonly bytes: Uint8Array
  /** Compression method; omitted uses deflate. */
  readonly method?: ZipMethod
}

/** The entry-level changes applied to a source archive. */
export interface ZipChanges {
  /** Existing entries whose uncompressed bytes change. */
  readonly replacements?: readonly ZipPartReplacement[]
  /** New entries in the caller's deterministic order. */
  readonly additions?: readonly ZipPartAddition[]
  /** Existing entries removed from the archive. */
  readonly removals?: readonly string[]
}

/** Writer determinism inputs for records this layer creates. */
export interface ZipWriteOptions {
  /** Deflate level for changed and added deflated entries. */
  readonly deflateLevel?: number
  /** DOS timestamp for added entries; omitted uses 1980-01-01 00:00. */
  readonly addedEntryDosDateTime?: DosDateTime
}

/** Bounds applied before an archive is indexed or one entry is inflated. */
export interface ZipReadLimits {
  /** Maximum number of central-directory entries accepted. */
  readonly maxEntries: number
  /** Maximum uncompressed bytes one entry may inflate to; omitted uses the package default. */
  readonly maxEntryUncompressedBytes?: number
  /** Maximum declared uncompressed bytes across all entries; omitted uses the package default. */
  readonly maxTotalUncompressedBytes?: number
}

/** Container structures the reader refuses before indexing entries. */
export type UnsupportedContainer =
  | 'cfb'
  | 'zip64'
  | 'multi-disk'
  | 'encrypted'
  | 'unsupported-method'
  | 'ambiguous-descriptor'

/** Typed reason a ZIP container cannot be read or written. */
export type ZipFailure =
  | {
    readonly type: 'unsupported-container'
    readonly container: UnsupportedContainer
    readonly detail: string
  }
  | { readonly type: 'invalid-archive'; readonly detail: string }
  | { readonly type: 'out-of-bounds'; readonly detail: string }
  | { readonly type: 'duplicate-name'; readonly detail: string }
  | { readonly type: 'entry-limit'; readonly detail: string }
  | { readonly type: 'too-large'; readonly detail: string }
  | { readonly type: 'unknown-entry'; readonly detail: string }
  | { readonly type: 'name-collision'; readonly detail: string }
  | { readonly type: 'crc-mismatch'; readonly detail: string }
