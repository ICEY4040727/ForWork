/** Editable OPC part map over an indexed ZIP archive with a no-op save fast path. */
import { createHash } from 'node:crypto'
import { posix } from 'node:path'
import { crc32, inflateRawSync } from 'node:zlib'
import { XMLParser } from 'fast-xml-parser'
import { asciiFoldKey, bytesEqual, ZipError, ZIP_METHOD_STORE } from './zip-format.ts'
import { readZipArchive } from './zip-reader.ts'
import { writeZipArchive } from './zip-writer.ts'
import type {
  SourcePackage,
  ZipArchive,
  ZipEntryRecord,
  ZipMethod,
  ZipPartAddition,
  ZipPartReplacement,
  ZipReadLimits,
  ZipWriteOptions,
} from './types.ts'

/** Bounds applied when a package is opened. */
export type OfficePackageLimits = ZipReadLimits

/** Creation options for a new part. */
export interface OfficePartOptions {
  /** Compression method; omitted uses deflate. */
  readonly method?: ZipMethod
}

/** One pending addition, keyed by its ASCII-folded name. */
interface PendingAddition {
  readonly name: string
  bytes: Uint8Array
  readonly method: ZipMethod | undefined
}

/** Shared UTF-8 encoder for lookup keys. */
const TEXT_ENCODER = new TextEncoder()

/** Default uncompressed-byte bound for one package entry. */
const DEFAULT_MAX_ENTRY_UNCOMPRESSED_BYTES = 64 * 1024 * 1024

/** Default declared-uncompressed-byte bound for a whole package. */
const DEFAULT_MAX_TOTAL_UNCOMPRESSED_BYTES = 256 * 1024 * 1024

/**
 * An OPC package whose parts can be read, replaced, added, and removed.
 *
 * `save` returns the source bytes unchanged when no part differs from the
 * source, so an untouched package keeps every original record.
 */
export class OfficePackage {
  /** The package as opened, including its SHA-256 digest. */
  readonly source: SourcePackage

  private readonly archive: ZipArchive

  private readonly parts = new Map<string, ZipEntryRecord>()

  private readonly occupiedNames = new Set<string>()

  private readonly additions = new Map<string, PendingAddition>()

  private readonly replacementBytes = new Map<string, Uint8Array>()

  private readonly removedKeys = new Set<string>()

  private readonly maxEntryUncompressedBytes: number

  /**
   * Index a package and enforce its container bounds.
   * @param source - Complete source archive bytes.
   * @param limits - Entry-count and uncompressed-byte bounds.
   * @throws {ZipError} When the archive is unsupported, malformed, or over a bound.
   */
  constructor(source: Uint8Array, limits: OfficePackageLimits) {
    const maxEntryUncompressedBytes
      = limits.maxEntryUncompressedBytes ?? DEFAULT_MAX_ENTRY_UNCOMPRESSED_BYTES
    const archive = readZipArchive(source, {
      maxEntries: limits.maxEntries,
      maxEntryUncompressedBytes,
      maxTotalUncompressedBytes: limits.maxTotalUncompressedBytes ?? DEFAULT_MAX_TOTAL_UNCOMPRESSED_BYTES,
    })
    this.maxEntryUncompressedBytes = maxEntryUncompressedBytes
    this.archive = archive
    this.source = {
      ...archive,
      digest: createHash('sha256').update(source).digest('hex'),
    }
    for (const entry of archive.centralOrder) {
      const key = asciiFoldKey(entry.nameBytes)
      this.occupiedNames.add(key)
      if (!entry.name.endsWith('/')) this.parts.set(key, entry)
    }
  }

  /**
   * Read a part's uncompressed bytes.
   * @param name - Part name; matched under ASCII case folding.
   * @returns A copy of the current uncompressed part bytes.
   * @throws {ZipError} When the part is absent, removed, or fails its CRC check.
   */
  readPart(name: string): Uint8Array {
    const key = foldName(name)
    const addition = this.additions.get(key)
    if (addition !== undefined) return addition.bytes.slice()
    const entry = this.parts.get(key)
    if (entry === undefined || this.removedKeys.has(key)) {
      throw new ZipError({
        type: 'unknown-entry',
        detail: `The package has no part named ${name}`,
      })
    }
    const replacement = this.replacementBytes.get(entry.name)
    if (replacement !== undefined) return replacement.slice()
    return this.inflateEntry(entry)
  }

  /**
   * Replace an existing part's bytes.
   *
   * Bytes equal to the current part clear the dirty mark, so a later save
   * returns the source archive unchanged. The comparison inflates the source
   * part; a caller that already established the bytes differ should use
   * {@link replacePartKnownDifferent} instead.
   * @param name - Part name; matched under ASCII case folding.
   * @param bytes - Uncompressed replacement bytes.
   * @throws {ZipError} When the part is absent or removed.
   */
  replacePart(name: string, bytes: Uint8Array): void {
    this.setReplacement(name, bytes, true)
  }

  /**
   * Replace a part whose bytes the caller already established differ.
   *
   * Unlike {@link replacePart}, this does not inflate the source part to check
   * for equality, so a writer that already compared its rendered bytes keeps
   * the save from decoding an unchanged source part again.
   * @param name - Part name; matched under ASCII case folding.
   * @param bytes - Uncompressed replacement bytes known to differ from the part.
   * @throws {ZipError} When the part is absent or removed.
   */
  replacePartKnownDifferent(name: string, bytes: Uint8Array): void {
    this.setReplacement(name, bytes, false)
  }

  /**
   * Store replacement bytes for an existing part or a pending addition.
   * @param name - Part name; matched under ASCII case folding.
   * @param bytes - Uncompressed replacement bytes.
   * @param compare - Whether to inflate the source part and clear the dirty
   * mark when the replacement equals it.
   * @throws {ZipError} When the part is absent or removed.
   */
  private setReplacement(name: string, bytes: Uint8Array, compare: boolean): void {
    const key = foldName(name)
    const addition = this.additions.get(key)
    if (addition !== undefined) {
      addition.bytes = bytes.slice()
      return
    }
    const entry = this.parts.get(key)
    if (entry === undefined || this.removedKeys.has(key)) {
      throw new ZipError({
        type: 'unknown-entry',
        detail: `Cannot replace ${name}: the package has no such part`,
      })
    }
    if (compare && bytesEqual(this.inflateEntry(entry), bytes)) {
      this.replacementBytes.delete(entry.name)
      return
    }
    this.replacementBytes.set(entry.name, bytes.slice())
  }

  /**
   * Add a new part.
   * @param name - New part name.
   * @param bytes - Uncompressed part bytes.
   * @param options - Compression method for the added entry.
   * @throws {ZipError} When the name collides under ASCII case folding.
   */
  addPart(name: string, bytes: Uint8Array, options: OfficePartOptions = {}): void {
    const key = foldName(name)
    if (this.additions.has(key) || (!this.removedKeys.has(key) && this.occupiedNames.has(key))) {
      throw new ZipError({
        type: 'name-collision',
        detail: `Cannot add ${name}: another entry has the same name under ASCII case folding`,
      })
    }
    this.additions.set(key, { name, bytes: bytes.slice(), method: options.method })
  }

  /**
   * Remove a part or cancel a pending addition.
   * @param name - Part name; matched under ASCII case folding.
   * @throws {ZipError} When the part is absent or already removed.
   */
  removePart(name: string): void {
    const key = foldName(name)
    if (this.additions.delete(key)) return
    const entry = this.parts.get(key)
    if (entry === undefined || this.removedKeys.has(key)) {
      throw new ZipError({
        type: 'unknown-entry',
        detail: `Cannot remove ${name}: the package has no such part`,
      })
    }
    this.removedKeys.add(key)
    this.replacementBytes.delete(entry.name)
  }

  /**
   * Serialize the package with every pending change.
   * @param options - Determinism inputs for records the writer creates.
   * @returns The original bytes when nothing changed, otherwise the new archive.
   */
  save(options: ZipWriteOptions = {}): Uint8Array {
    const replacements: ZipPartReplacement[] = []
    const removals: string[] = []
    for (const entry of this.archive.centralOrder) {
      const key = asciiFoldKey(entry.nameBytes)
      if (this.removedKeys.has(key)) {
        removals.push(entry.name)
        continue
      }
      const bytes = this.replacementBytes.get(entry.name)
      if (bytes !== undefined) replacements.push({ name: entry.name, bytes })
    }
    const additions: ZipPartAddition[] = []
    for (const addition of this.additions.values()) {
      additions.push(addition.method === undefined
        ? { name: addition.name, bytes: addition.bytes }
        : { name: addition.name, bytes: addition.bytes, method: addition.method })
    }
    return writeZipArchive(this.archive, { replacements, additions, removals }, options)
  }

  /**
   * List the file parts the pending output would carry, in archive order.
   * @returns Surviving source parts followed by pending additions, without directory entries.
   */
  partNames(): readonly string[] {
    const names: string[] = []
    for (const entry of this.archive.centralOrder) {
      const key = asciiFoldKey(entry.nameBytes)
      if (entry.name.endsWith('/') || this.removedKeys.has(key)) continue
      names.push(entry.name)
    }
    for (const addition of this.additions.values()) names.push(addition.name)
    return names
  }

  /**
   * Inflate one source entry and verify its declared length and CRC.
   * @param entry - Source entry to decode.
   * @returns The verified uncompressed bytes.
   * @throws {ZipError} `too-large` when the declared size exceeds the per-entry
   * bound, `invalid-archive` when inflation fails or the bytes disagree with the
   * header, and `crc-mismatch` when the bytes do not match the stored CRC.
   */
  private inflateEntry(entry: ZipEntryRecord): Uint8Array {
    if (entry.uncompressedSize > this.maxEntryUncompressedBytes) {
      throw new ZipError({
        type: 'too-large',
        detail: `Part ${entry.name} declares ${String(entry.uncompressedSize)} uncompressed bytes; `
          + `the per-entry limit is ${String(this.maxEntryUncompressedBytes)}`,
      })
    }
    const compressed = this.archive.bytes.subarray(entry.dataOffset, entry.dataOffset + entry.compressedSize)
    let output: Uint8Array
    if (entry.method === ZIP_METHOD_STORE) {
      output = compressed.slice()
    } else {
      try {
        const inflated = inflateRawSync(compressed, {
          maxOutputLength: Math.min(entry.uncompressedSize, this.maxEntryUncompressedBytes) + 1,
        })
        output = new Uint8Array(inflated.buffer, inflated.byteOffset, inflated.byteLength)
      } catch (cause) {
        throw new ZipError({
          type: 'invalid-archive',
          detail: `Part ${entry.name} cannot be inflated: ${String(cause)}`,
        })
      }
    }
    if (output.byteLength !== entry.uncompressedSize) {
      throw new ZipError({
        type: 'invalid-archive',
        detail: `Part ${entry.name} inflated to ${output.byteLength} bytes; the header declares ${entry.uncompressedSize}`,
      })
    }
    if (crc32(output) !== entry.crc32) {
      throw new ZipError({
        type: 'crc-mismatch',
        detail: `Part ${entry.name} does not match its stored CRC-32`,
      })
    }
    return output
  }
}

/**
 * Open an OPC package for in-place editing.
 *
 * Omitted uncompressed-byte bounds use the package defaults; a caller can lower
 * them, and the per-entry bound also caps what one part may inflate to.
 * @param source - Complete source archive bytes.
 * @param limits - Entry-count and uncompressed-byte bounds.
 * @returns The indexed package.
 * @throws {ZipError} When the archive is unsupported, malformed, or over a bound.
 */
export function open(source: Uint8Array, limits: OfficePackageLimits): OfficePackage {
  return new OfficePackage(source, limits)
}

/** Content-types part that declares every other part's media type. */
const CONTENT_TYPES_PART = '[Content_Types].xml'

/** Package-relative relationships part. */
const ROOT_RELATIONSHIPS_PART = '_rels/.rels'

/** Suffix shared by every relationship part. */
const RELATIONSHIPS_SUFFIX = '.rels'

/** Decode part bytes as UTF-8 for the declaration parsers. */
const TEXT_DECODER = new TextDecoder()

/** Parser for the two OPC declaration parts; both are plain attribute XML. */
const OPC_PARSER = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  parseAttributeValue: false,
  isArray: name => name === 'Default' || name === 'Override' || name === 'Relationship',
})

/** One `Default` or `Override` declaration in `[Content_Types].xml`. */
interface ContentTypeDeclaration {
  /** Declared file extension, for a `Default`. */
  readonly '@_Extension'?: string
  /** Declared media type. */
  readonly '@_ContentType'?: string
  /** Declared part name, for an `Override`. */
  readonly '@_PartName'?: string
}

/** Parsed `[Content_Types].xml`. */
interface ContentTypesDocument {
  /** Root element carrying the declarations. */
  readonly Types?: {
    /** Extension-wide declarations. */
    readonly Default?: readonly ContentTypeDeclaration[]
    /** Part-specific declarations. */
    readonly Override?: readonly ContentTypeDeclaration[]
  }
}

/** One relationship declaration. */
interface RelationshipDeclaration {
  /** Relationship id. */
  readonly '@_Id'?: string
  /** Target part name or external URI. */
  readonly '@_Target'?: string
  /** `External` when the target is not a package part. */
  readonly '@_TargetMode'?: string
}

/** Content-type declarations indexed for part lookup. */
interface DeclaredContentTypes {
  /** Lowercased extension mapped to its media type. */
  readonly defaults: ReadonlyMap<string, string>
  /** Lowercased part name mapped to its declared part name and media type. */
  readonly overrides: ReadonlyMap<string, { readonly partName: string; readonly contentType: string }>
}

/**
 * Require a package's OPC graph to be closed: every internal relationship
 * target resolves to an existing part, every part has a content-type
 * declaration, and no declaration names a missing part.
 * @param packageFile - Package whose pending parts are the prospective output.
 * @throws {ZipError} When the package is not a closed OPC graph.
 */
export function assertOpcGraph(packageFile: OfficePackage): void {
  const parts = new Map<string, string>()
  for (const name of packageFile.partNames()) parts.set(name.toLowerCase(), name)
  const contentTypes = declaredContentTypes(packageFile, parts)
  for (const name of parts.values()) {
    if (name === CONTENT_TYPES_PART) continue
    if (contentTypeOf(name, contentTypes) !== undefined) continue
    invalidOpc(`The package declares no content type for ${name}`)
  }
  for (const [partName, declaration] of contentTypes.overrides) {
    if (parts.has(partName)) continue
    invalidOpc(`[Content_Types].xml has an Override for ${declaration.partName}, which is not a package part`)
  }
  for (const name of parts.values()) {
    if (!name.toLowerCase().endsWith(RELATIONSHIPS_SUFFIX)) continue
    assertRelationshipTargets(packageFile, name, parts)
  }
}

/**
 * Parse `[Content_Types].xml` and index its declarations.
 * @param packageFile - Package whose pending parts are the prospective output.
 * @param parts - Lowercased part names to actual names.
 * @returns Declarations indexed by lowercased extension and part name.
 * @throws {ZipError} When the part is missing or its declarations cannot be read.
 */
function declaredContentTypes(
  packageFile: OfficePackage,
  parts: ReadonlyMap<string, string>,
): DeclaredContentTypes {
  const actualName = parts.get(CONTENT_TYPES_PART.toLowerCase())
  if (actualName === undefined) invalidOpc('The package has no [Content_Types].xml part')
  let parsed: unknown
  try {
    parsed = OPC_PARSER.parse(TEXT_DECODER.decode(packageFile.readPart(actualName)))
  } catch (error) {
    if (error instanceof ZipError) throw error
    invalidOpc(`[Content_Types].xml cannot be parsed: ${messageOf(error)}`)
  }
  const types = contentTypesElement(parsed)
  if (types === undefined) invalidOpc('[Content_Types].xml has no Types element')
  const defaults = new Map<string, string>()
  for (const node of types.Default ?? []) {
    const extension = node['@_Extension']
    const contentType = node['@_ContentType']
    if (typeof extension !== 'string' || typeof contentType !== 'string') {
      invalidOpc('[Content_Types].xml has a Default without an extension or content type')
    }
    defaults.set(extension.toLowerCase(), contentType)
  }
  const overrides = new Map<string, { readonly partName: string; readonly contentType: string }>()
  for (const node of types.Override ?? []) {
    const partName = node['@_PartName']
    const contentType = node['@_ContentType']
    if (typeof partName !== 'string' || typeof contentType !== 'string') {
      invalidOpc('[Content_Types].xml has an Override without a part name or content type')
    }
    const normalized = normalizePartName(partName)
    overrides.set(normalized.toLowerCase(), { partName: normalized, contentType })
  }
  return { defaults, overrides }
}

/**
 * Narrow the parser output to a `Types` element.
 * @param value - Untrusted `XMLParser` output.
 * @returns The declarations container, or `undefined` when the part has none.
 */
function contentTypesElement(value: unknown): NonNullable<ContentTypesDocument['Types']> | undefined {
  // The OPC parser always returns an object; only its Types field can be absent or non-object.
  const types = (value as ContentTypesDocument).Types
  if (typeof types !== 'object') return undefined
  return types
}

/**
 * Resolve one part's declared content type.
 * @param name - Package part name.
 * @param contentTypes - Declarations read from the package.
 * @returns The declared media type, or `undefined` when the part is undeclared.
 */
function contentTypeOf(name: string, contentTypes: DeclaredContentTypes): string | undefined {
  const override = contentTypes.overrides.get(name.toLowerCase())
  if (override !== undefined) return override.contentType
  const dot = name.lastIndexOf('.')
  if (dot < 0) return undefined
  return contentTypes.defaults.get(name.slice(dot + 1).toLowerCase())
}

/**
 * Require every internal target one relationship part declares to resolve.
 * @param packageFile - Package whose pending parts are the prospective output.
 * @param relsName - Relationship part name.
 * @param parts - Lowercased part names to actual names.
 * @throws {ZipError} When a target is missing or names no package part.
 */
function assertRelationshipTargets(
  packageFile: OfficePackage,
  relsName: string,
  parts: ReadonlyMap<string, string>,
): void {
  const sourcePart = relationshipsSourcePart(relsName)
  if (sourcePart === undefined) return
  let parsed: unknown
  try {
    parsed = OPC_PARSER.parse(TEXT_DECODER.decode(packageFile.readPart(relsName)))
  } catch (error) {
    if (error instanceof ZipError) throw error
    invalidOpc(`${relsName} cannot be parsed: ${messageOf(error)}`)
  }
  const relationships = relationshipDeclarations(parsed)
  if (relationships === undefined) invalidOpc(`${relsName} has no Relationships element`)
  for (const relationship of relationships) {
    if (relationship['@_TargetMode'] === 'External') continue
    const id = relationship['@_Id'] ?? ''
    const target = relationship['@_Target']
    if (target === undefined) invalidOpc(`Relationship ${id} in ${relsName} has no Target attribute`)
    const resolved = resolveRelationshipTarget(sourcePart, target)
    if (resolved !== undefined && parts.has(resolved.toLowerCase())) continue
    invalidOpc(`Relationship ${id} in ${relsName} targets ${target}, which resolves to no package part`)
  }
}

/**
 * Narrow the parser output to a relationship part's declarations.
 * @param value - Untrusted `XMLParser` output.
 * @returns Declarations, or `undefined` when the part has no `Relationships` element.
 */
function relationshipDeclarations(value: unknown): readonly RelationshipDeclaration[] | undefined {
  // The OPC parser always returns an object; only the Relationships field and its
  // declarations are optional, and `isArray` guarantees a present Relationship is an array.
  const root = (value as { readonly Relationships?: unknown }).Relationships
  if (typeof root !== 'object' || root === null) return undefined
  const declarations = (root as { readonly Relationship?: unknown }).Relationship
  return declarations === undefined ? [] : declarations as readonly RelationshipDeclaration[]
}

/** The part a relationship part belongs to; the package root for `_rels/.rels`. */
function relationshipsSourcePart(relsName: string): string | undefined {
  if (relsName === ROOT_RELATIONSHIPS_PART) return ''
  const match = /^(?:(.*)\/)?_rels\/([^/]+)\.rels$/u.exec(relsName)
  if (match === null || match[2] === undefined) return undefined
  const directory = match[1] ?? ''
  return directory === '' ? match[2] : `${directory}/${match[2]}`
}

/** Resolve a relationship target against the part that declares it. */
function resolveRelationshipTarget(sourcePart: string, target: string): string | undefined {
  const fragment = target.indexOf('#')
  const withoutFragment = fragment < 0 ? target : target.slice(0, fragment)
  const query = withoutFragment.indexOf('?')
  const withoutQuery = query < 0 ? withoutFragment : withoutFragment.slice(0, query)
  let decoded = withoutQuery
  try {
    decoded = decodeURIComponent(withoutQuery)
  } catch {
    // A malformed percent escape cannot name a package part; the literal target is reported instead.
  }
  const absolute = decoded.startsWith('/')
  const baseDirectory = sourcePart === '' ? '' : posix.dirname(sourcePart)
  const resolved = absolute
    ? posix.normalize(decoded.slice(1))
    : posix.normalize(posix.join(baseDirectory, decoded))
  // `posix.normalize` never returns an empty string, so only an escaping parent segment is refused.
  if (resolved.startsWith('../') || resolved === '..') return undefined
  return resolved.replace(/^\/+/u, '')
}

/** Strip the absolute prefix OPC part names carry and decode URI escapes. */
function normalizePartName(partName: string): string {
  const stripped = partName.replace(/^\/+/u, '')
  try {
    return decodeURIComponent(stripped)
  } catch {
    // A malformed escape cannot name a package part; the literal name is reported instead.
    return stripped
  }
}

/** Throw a typed invalid-archive failure for a malformed OPC graph. */
function invalidOpc(detail: string): never {
  throw new ZipError({ type: 'invalid-archive', detail })
}

/** Render a non-Error parser rejection as a diagnostic message. */
function messageOf(error: unknown): string {
  /* v8 ignore next -- both OPC parsers throw Error instances */
  return error instanceof Error ? error.message : String(error)
}

/**
 * Build the ASCII-folded lookup key for a part name.
 * @param name - Part name supplied by a caller.
 * @returns The folded key used by the part map.
 */
function foldName(name: string): string {
  return asciiFoldKey(TEXT_ENCODER.encode(name))
}
