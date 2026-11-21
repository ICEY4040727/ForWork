/** OPC validation and byte-diff evidence helpers for OOXML preservation tests. */
import { createHash } from 'node:crypto'
import { posix } from 'node:path'
import { XMLParser } from 'fast-xml-parser'
import { SyntaxValidator } from 'fast-xml-validator'
import { strFromU8, unzipSync } from 'fflate'

const CONTENT_TYPES_PART = '[Content_Types].xml'
const ROOT_RELATIONSHIPS_PART = '_rels/.rels'
const RELATIONSHIP_NAMESPACES: ReadonlySet<string> = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
])
const RELATIONSHIP_REFERENCE_ATTRIBUTES: ReadonlySet<string> = new Set(['id', 'embed', 'link'])
const XML_PART_EXTENSIONS: readonly string[] = ['.xml', '.rels', '.vml']

/** One ZIP entry's relationship to its source or written package. */
export type EntryStatus = 'identical' | 'changed' | 'added' | 'removed'

/** Byte-level comparison for one entry name present in either package. */
export interface EntryDiff {
  /** Entry name as stored in the archive. */
  readonly name: string
  /** Whether the entry is identical, changed, added, or removed. */
  readonly status: EntryStatus
  /** Source decompressed byte length, or null when the entry is added. */
  readonly sourceSize: number | null
  /** Written decompressed byte length, or null when the entry is removed. */
  readonly writtenSize: number | null
  /** SHA-256 hex of the source entry, or null when the entry is added. */
  readonly sourceSha256: string | null
  /** SHA-256 hex of the written entry, or null when the entry is removed. */
  readonly writtenSha256: string | null
  /** First differing decompressed byte offset, or null for identical entries. */
  readonly firstDifferenceOffset: number | null
}

/** Set comparison for relationship declarations or content-type declarations. */
export interface PackageSetDiff {
  /** Canonical declarations in the source package. */
  readonly source: readonly string[]
  /** Canonical declarations in the written package. */
  readonly written: readonly string[]
  /** Declarations present only in the written package. */
  readonly added: readonly string[]
  /** Declarations present only in the source package. */
  readonly removed: readonly string[]
}

/** Entry-name expectations for a write that is allowed to change the package. */
export interface PackageDiffExpectation {
  /** Entries allowed to differ from the source. */
  readonly changed?: readonly string[]
  /** Entries the write may add. */
  readonly added?: readonly string[]
  /** Entries the write may remove. */
  readonly removed?: readonly string[]
}

/** Complete byte-diff report comparing a source package with written bytes. */
export interface PackageDiff {
  /** True when every entry status matches the expectation. */
  readonly ok: boolean
  /** One comparison per entry name in source order, then written-only names. */
  readonly entries: readonly EntryDiff[]
  /** Relationship declarations compared as canonical strings. */
  readonly relationships: PackageSetDiff
  /** Content-type declarations compared as canonical strings. */
  readonly contentTypes: PackageSetDiff
  /** Human-readable expectation and parse failures. */
  readonly failures: readonly string[]
}

/** One structural OPC violation found in a package. */
export interface OpcViolation {
  /** Stable violation category. */
  readonly code:
    | 'missing-content-types-part'
    | 'missing-content-type'
    | 'content-type-override-missing'
    | 'relationship-target-missing'
    | 'duplicate-relationship-id'
    | 'dangling-relationship-reference'
    | 'malformed-xml'
  /** Part the violation is attributed to. */
  readonly part: string
  /** Violation detail, including the missing target or relationship id. */
  readonly detail: string
}

interface RelationshipNode {
  readonly '@_Id'?: string
  readonly '@_Type'?: string
  readonly '@_Target'?: string
  readonly '@_TargetMode'?: string
}

interface RelationshipsDocument {
  readonly Relationships?: {
    readonly Relationship?: readonly RelationshipNode[]
  }
}

interface ContentTypeDefaultNode {
  readonly '@_Extension'?: string
  readonly '@_ContentType'?: string
}

interface ContentTypeOverrideNode {
  readonly '@_PartName'?: string
  readonly '@_ContentType'?: string
}

interface ContentTypesDocument {
  readonly Types?: {
    readonly Default?: readonly ContentTypeDefaultNode[]
    readonly Override?: readonly ContentTypeOverrideNode[]
  }
}

interface ParsedContentTypes {
  readonly defaults: ReadonlyMap<string, string>
  readonly overrides: ReadonlyMap<string, ContentTypeOverrideNode>
}

interface StartTagAttribute {
  readonly name: string
  readonly value: string
}

/**
 * Decompress every archive entry, preserving central-directory order.
 * @param bytes - Complete ZIP package bytes.
 * @returns Entry payloads keyed by their stored names.
 */
export function unzipEntries(bytes: Uint8Array): Map<string, Uint8Array> {
  return new Map(Object.entries(unzipSync(bytes)))
}

/**
 * Compute the lowercase SHA-256 hex digest of a byte sequence.
 * @param bytes - Bytes to hash.
 * @returns 64 lowercase hexadecimal characters.
 */
export function sha256Hex(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

/**
 * Compare two OOXML packages entry by entry and report every difference.
 * @param source - Package bytes read from the source file.
 * @param written - Package bytes produced by the writer under test.
 * @param expectation - Entry statuses the write is allowed to produce; omitted means no change is allowed.
 * @returns Entry diffs, relationship and content-type sets, and expectation failures.
 */
export function comparePackages(
  source: Uint8Array,
  written: Uint8Array,
  expectation: PackageDiffExpectation = {},
): PackageDiff {
  const sourceEntries = unzipEntries(source)
  const writtenEntries = unzipEntries(written)
  const names = [...sourceEntries.keys()]
  for (const name of writtenEntries.keys()) {
    if (!sourceEntries.has(name)) names.push(name)
  }

  const failures: string[] = []
  const expectedChanged = new Set(expectation.changed ?? [])
  const expectedAdded = new Set(expectation.added ?? [])
  const expectedRemoved = new Set(expectation.removed ?? [])
  const entries = names.map((name): EntryDiff => {
    const sourceBytes = sourceEntries.get(name)
    const writtenBytes = writtenEntries.get(name)
    if (sourceBytes === undefined) {
      if (!expectedAdded.has(name)) failures.push(`unexpected added entry ${name}`)
      return entryDiff(name, 'added', undefined, writtenBytes)
    }
    if (writtenBytes === undefined) {
      if (!expectedRemoved.has(name)) failures.push(`unexpected removed entry ${name}`)
      return entryDiff(name, 'removed', sourceBytes, undefined)
    }
    const offset = firstDifferenceOffset(sourceBytes, writtenBytes)
    if (offset === null) {
      if (expectedChanged.has(name)) failures.push(`expected ${name} to change but it is identical`)
      if (expectedAdded.has(name)) failures.push(`expected added entry ${name} but it already exists`)
      if (expectedRemoved.has(name)) failures.push(`expected removed entry ${name} but it exists`)
      return entryDiff(name, 'identical', sourceBytes, writtenBytes)
    }
    if (!expectedChanged.has(name)) {
      failures.push(`unexpected change in ${name} at offset ${offset}`)
    }
    return entryDiff(name, 'changed', sourceBytes, writtenBytes, offset)
  })

  for (const name of expectedChanged) {
    const entry = entries.find(candidate => candidate.name === name)
    if (entry?.status !== 'changed') failures.push(`expected ${name} to change but it is ${entry?.status ?? 'absent'}`)
  }
  for (const name of expectedAdded) {
    const entry = entries.find(candidate => candidate.name === name)
    if (entry?.status !== 'added') failures.push(`expected added entry ${name} but it is ${entry?.status ?? 'absent'}`)
  }
  for (const name of expectedRemoved) {
    const entry = entries.find(candidate => candidate.name === name)
    if (entry?.status !== 'removed') failures.push(`expected removed entry ${name} but it is ${entry?.status ?? 'absent'}`)
  }

  return {
    ok: failures.length === 0,
    entries,
    relationships: parseSetDiff(sourceEntries, writtenEntries, relationshipDeclarations),
    contentTypes: parseSetDiff(sourceEntries, writtenEntries, contentTypeDeclarations),
    failures,
  }
}

/**
 * Render a package diff as the Markdown table used in preservation test failures.
 * @param diff - Report returned by {@link comparePackages}.
 * @returns Markdown with one row per entry, declaration-set counts, and failures.
 */
export function formatDiffReport(diff: PackageDiff): string {
  const lines = [
    '| Entry | Status | Source bytes | Written bytes | Source SHA-256 | Written SHA-256 | First difference |',
    '| --- | --- | --- | --- | --- | --- | --- |',
  ]
  for (const entry of diff.entries) {
    lines.push([
      `| \`${entry.name}\``,
      entry.status,
      entry.sourceSize ?? '-',
      entry.writtenSize ?? '-',
      entry.sourceSha256 ?? '-',
      entry.writtenSha256 ?? '-',
      entry.firstDifferenceOffset ?? '-',
    ].join(' | ') + ' |')
  }
  lines.push(
    '',
    `Relationships: ${setSummary(diff.relationships)}`,
    `Content types: ${setSummary(diff.contentTypes)}`,
  )
  if (diff.failures.length > 0) {
    lines.push('', 'Failures:')
    for (const failure of diff.failures) lines.push(`- ${failure}`)
  }
  return lines.join('\n')
}

/**
 * Validate package structure: relationship closure, content-type coverage,
 * relationship-id uniqueness, reference closure, and XML well-formedness.
 * @param entries - Decompressed package entries.
 * @returns Every structural violation found; an empty array means the package is valid.
 */
export function validateOpc(entries: ReadonlyMap<string, Uint8Array>): OpcViolation[] {
  const violations: OpcViolation[] = []
  const partNames = new Map<string, string>()
  for (const name of entries.keys()) {
    if (name.endsWith('/')) continue
    partNames.set(name.toLowerCase(), name)
  }

  const contentTypesBytes = entries.get(CONTENT_TYPES_PART)
  let contentTypes: ParsedContentTypes | undefined
  if (contentTypesBytes === undefined) {
    violations.push({
      code: 'missing-content-types-part',
      part: CONTENT_TYPES_PART,
      detail: 'package has no [Content_Types].xml part',
    })
  } else {
    contentTypes = parseContentTypes(contentTypesBytes)
    if (contentTypes === undefined) {
      violations.push({
        code: 'malformed-xml',
        part: CONTENT_TYPES_PART,
        detail: 'content-type declarations cannot be parsed',
      })
    }
  }

  const relationshipsIds = relationshipIdsByPart(entries, violations)

  for (const name of partNames.values()) {
    if (name === CONTENT_TYPES_PART) continue
    if (contentTypes !== undefined && contentTypeFor(name, contentTypes) === undefined) {
      violations.push({
        code: 'missing-content-type',
        part: name,
        detail: 'part has no matching Default or Override declaration',
      })
    }
  }

  if (contentTypes !== undefined) {
    for (const [partName, declaration] of contentTypes.overrides) {
      if (!partNames.has(partName)) {
        violations.push({
          code: 'content-type-override-missing',
          part: CONTENT_TYPES_PART,
          detail: `Override PartName="${declaration['@_PartName'] ?? partName}" names no package part`,
        })
      }
    }
  }

  validateRelationshipTargets(entries, partNames, violations)

  for (const name of partNames.values()) {
    if (!XML_PART_EXTENSIONS.some(extension => name.toLowerCase().endsWith(extension))) continue
    const bytes = entries.get(name)
    if (bytes === undefined) continue
    validateXml(name, bytes, violations)
    if (name === CONTENT_TYPES_PART || name.toLowerCase().endsWith('.rels')) continue
    const relsName = relationshipsPartName(name)
    const availableIds = relsName === undefined ? undefined : relationshipsIds.get(relsName)
    for (const reference of relationshipReferences(strFromU8(bytes))) {
      if (availableIds?.has(reference.value) !== true) {
        violations.push({
          code: 'dangling-relationship-reference',
          part: name,
          detail: `attribute ${reference.attribute} references relationship id ${reference.value}`,
        })
      }
    }
  }

  return violations
}

function entryDiff(
  name: string,
  status: EntryStatus,
  sourceBytes: Uint8Array | undefined,
  writtenBytes: Uint8Array | undefined,
  firstDifference: number | null = null,
): EntryDiff {
  return {
    name,
    status,
    sourceSize: sourceBytes?.byteLength ?? null,
    writtenSize: writtenBytes?.byteLength ?? null,
    sourceSha256: sourceBytes === undefined ? null : sha256Hex(sourceBytes),
    writtenSha256: writtenBytes === undefined ? null : sha256Hex(writtenBytes),
    firstDifferenceOffset: firstDifference,
  }
}

function firstDifferenceOffset(source: Uint8Array, written: Uint8Array): number | null {
  const limit = Math.min(source.byteLength, written.byteLength)
  for (let index = 0; index < limit; index += 1) {
    if (source[index] !== written[index]) return index
  }
  return source.byteLength === written.byteLength ? null : limit
}

function parseSetDiff(
  source: ReadonlyMap<string, Uint8Array>,
  written: ReadonlyMap<string, Uint8Array>,
  extract: (entries: ReadonlyMap<string, Uint8Array>) => readonly string[],
): PackageSetDiff {
  const sourceSet = extract(source)
  const writtenSet = extract(written)
  const sourceValues = new Set(sourceSet)
  const writtenValues = new Set(writtenSet)
  return {
    source: sourceSet,
    written: writtenSet,
    added: writtenSet.filter(value => !sourceValues.has(value)),
    removed: sourceSet.filter(value => !writtenValues.has(value)),
  }
}

function setSummary(diff: PackageSetDiff): string {
  return `${diff.added.length} added, ${diff.removed.length} removed, ${diff.written.length} total`
}

function relationshipDeclarations(entries: ReadonlyMap<string, Uint8Array>): readonly string[] {
  const declarations: string[] = []
  for (const [name, bytes] of entries) {
    if (!name.toLowerCase().endsWith('.rels')) continue
    const relationships = parseRelationships(strFromU8(bytes))
    for (const relationship of relationships ?? []) {
      const id = relationship['@_Id'] ?? ''
      const type = relationship['@_Type'] ?? ''
      const target = relationship['@_Target'] ?? ''
      const mode = relationship['@_TargetMode'] === 'External' ? ' [external]' : ''
      declarations.push(`${name}#${id} ${type} -> ${target}${mode}`)
    }
  }
  return declarations.sort()
}

function contentTypeDeclarations(entries: ReadonlyMap<string, Uint8Array>): readonly string[] {
  const bytes = entries.get(CONTENT_TYPES_PART)
  if (bytes === undefined) return []
  const parsed = parseContentTypes(bytes)
  if (parsed === undefined) return []
  const declarations: string[] = []
  for (const [extension, contentType] of parsed.defaults) {
    declarations.push(`Default ${extension}=${contentType}`)
  }
  for (const [partName, node] of parsed.overrides) {
    declarations.push(`Override ${partName}=${node['@_ContentType'] ?? ''}`)
  }
  return declarations.sort()
}

function parseContentTypes(bytes: Uint8Array): ParsedContentTypes | undefined {
  let parsed: ContentTypesDocument
  try {
    parsed = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
      isArray: name => name === 'Default' || name === 'Override',
    }).parse(strFromU8(bytes)) as ContentTypesDocument
  } catch {
    return undefined
  }
  const defaults = new Map<string, string>()
  for (const node of parsed.Types?.Default ?? []) {
    const extension = node['@_Extension']
    const contentType = node['@_ContentType']
    if (extension !== undefined && contentType !== undefined) defaults.set(extension.toLowerCase(), contentType)
  }
  const overrides = new Map<string, ContentTypeOverrideNode>()
  for (const node of parsed.Types?.Override ?? []) {
    const partName = normalizePartName(node['@_PartName'])
    if (partName !== undefined) overrides.set(partName.toLowerCase(), node)
  }
  return { defaults, overrides }
}

function contentTypeFor(name: string, contentTypes: ParsedContentTypes): string | undefined {
  const override = contentTypes.overrides.get(name.toLowerCase())
  if (override !== undefined) return override['@_ContentType']
  const dot = name.lastIndexOf('.')
  if (dot < 0) return undefined
  return contentTypes.defaults.get(name.slice(dot + 1).toLowerCase())
}

function normalizePartName(partName: string | undefined): string | undefined {
  if (partName === undefined) return undefined
  return partName.replace(/^\/+/u, '')
}

function parseRelationships(xml: string): readonly RelationshipNode[] | undefined {
  try {
    const parsed = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
      isArray: name => name === 'Relationship',
    }).parse(xml) as RelationshipsDocument
    return parsed.Relationships?.Relationship
  } catch {
    return undefined
  }
}

function relationshipIdsByPart(
  entries: ReadonlyMap<string, Uint8Array>,
  violations: OpcViolation[],
): ReadonlyMap<string, ReadonlySet<string>> {
  const ids = new Map<string, ReadonlySet<string>>()
  for (const [name, bytes] of entries) {
    if (!name.toLowerCase().endsWith('.rels')) continue
    const relationships = parseRelationships(strFromU8(bytes))
    if (relationships === undefined) continue
    const partIds = new Set<string>()
    for (const relationship of relationships) {
      const id = relationship['@_Id']
      if (id === undefined) continue
      if (partIds.has(id)) {
        violations.push({
          code: 'duplicate-relationship-id',
          part: name,
          detail: `relationship id ${id} is declared more than once`,
        })
      }
      partIds.add(id)
    }
    ids.set(name, partIds)
  }
  return ids
}

function validateRelationshipTargets(
  entries: ReadonlyMap<string, Uint8Array>,
  partNames: ReadonlyMap<string, string>,
  violations: OpcViolation[],
): void {
  for (const [name, bytes] of entries) {
    if (!name.toLowerCase().endsWith('.rels')) continue
    const sourcePart = relationshipsSourcePart(name)
    if (sourcePart === undefined) continue
    for (const relationship of parseRelationships(strFromU8(bytes)) ?? []) {
      if (relationship['@_TargetMode'] === 'External') continue
      const target = relationship['@_Target']
      const id = relationship['@_Id'] ?? ''
      if (target === undefined) {
        violations.push({
          code: 'relationship-target-missing',
          part: name,
          detail: `relationship id ${id} has no Target attribute`,
        })
        continue
      }
      const resolved = resolveRelationshipTarget(sourcePart, target)
      if (resolved === undefined || !partNames.has(resolved.toLowerCase())) {
        violations.push({
          code: 'relationship-target-missing',
          part: name,
          detail: `relationship id ${id} target ${target} resolves to ${resolved ?? 'no package part'}`,
        })
      }
    }
  }
}

function validateXml(name: string, bytes: Uint8Array, violations: OpcViolation[]): void {
  try {
    SyntaxValidator.validate(strFromU8(bytes), { xmlDeclaraion: { optional: true } })
  } catch (error) {
    violations.push({
      code: 'malformed-xml',
      part: name,
      detail: error instanceof Error ? error.message : String(error),
    })
  }
}

function relationshipsPartName(partName: string): string | undefined {
  const directory = posix.dirname(partName)
  const baseName = posix.basename(partName)
  if (baseName.length === 0) return undefined
  return directory === '.' ? `_rels/${baseName}.rels` : `${directory}/_rels/${baseName}.rels`
}

function relationshipsSourcePart(relsName: string): string | undefined {
  if (relsName === ROOT_RELATIONSHIPS_PART) return ''
  const match = /^(?:(.*)\/)?_rels\/([^/]+)\.rels$/u.exec(relsName)
  if (match === null || match[2] === undefined) return undefined
  const directory = match[1] ?? ''
  return directory === '' ? match[2] : `${directory}/${match[2]}`
}

function resolveRelationshipTarget(sourcePart: string, target: string): string | undefined {
  const withoutFragment = target.split('#')[0] ?? target
  const withoutQuery = withoutFragment.split('?')[0] ?? withoutFragment
  let decoded = withoutQuery
  try {
    decoded = decodeURIComponent(withoutQuery)
  } catch {
    // Keep the literal target when it contains an invalid percent escape.
  }
  const absolute = decoded.startsWith('/')
  const baseDirectory = sourcePart === '' ? '' : posix.dirname(sourcePart)
  const resolved = absolute
    ? posix.normalize(decoded.slice(1))
    : posix.normalize(posix.join(baseDirectory, decoded))
  if (resolved.length === 0 || resolved.startsWith('../') || resolved === '..') return undefined
  return resolved.replace(/^\/+/u, '')
}

function relationshipReferences(xml: string): readonly { attribute: string; value: string }[] {
  const references: { attribute: string; value: string }[] = []
  scanStartTags(xml, (_tagName, attributes, namespaces) => {
    for (const attribute of attributes) {
      const colon = attribute.name.indexOf(':')
      if (colon <= 0) continue
      const localName = attribute.name.slice(colon + 1)
      if (!RELATIONSHIP_REFERENCE_ATTRIBUTES.has(localName)) continue
      const namespace = namespaces.get(attribute.name.slice(0, colon))
      if (namespace !== undefined && RELATIONSHIP_NAMESPACES.has(namespace)) {
        references.push({ attribute: attribute.name, value: attribute.value })
      }
    }
  })
  return references
}

function scanStartTags(
  xml: string,
  visit: (
    tagName: string,
    attributes: readonly StartTagAttribute[],
    namespaces: ReadonlyMap<string, string>,
  ) => void,
): void {
  const scopes: Map<string, string>[] = [new Map<string, string>()]
  let index = 0
  while (index < xml.length) {
    const open = xml.indexOf('<', index)
    if (open < 0) return
    if (xml.startsWith('<!--', open)) {
      index = skipPast(xml, '-->', open + 4)
      continue
    }
    if (xml.startsWith('<![CDATA[', open)) {
      index = skipPast(xml, ']]>', open + 9)
      continue
    }
    if (xml.startsWith('<?', open)) {
      index = skipPast(xml, '?>', open + 2)
      continue
    }
    if (xml.startsWith('<!', open)) {
      index = skipDeclaration(xml, open + 2)
      continue
    }
    if (xml.startsWith('</', open)) {
      const close = xml.indexOf('>', open + 2)
      if (close < 0) return
      if (scopes.length > 1) scopes.pop()
      index = close + 1
      continue
    }

    let cursor = open + 1
    while (cursor < xml.length && !isWhitespace(xml[cursor]) && xml[cursor] !== '>' && xml[cursor] !== '/') {
      cursor += 1
    }
    const tagName = xml.slice(open + 1, cursor)
    const namespaces = new Map(scopes[scopes.length - 1] ?? [])
    const attributes: StartTagAttribute[] = []
    let selfClosing = false
    while (cursor < xml.length) {
      while (cursor < xml.length && isWhitespace(xml[cursor])) cursor += 1
      if (xml[cursor] === '>') {
        cursor += 1
        break
      }
      if (xml[cursor] === '/' && xml[cursor + 1] === '>') {
        selfClosing = true
        cursor += 2
        break
      }
      const nameStart = cursor
      while (
        cursor < xml.length
        && !isWhitespace(xml[cursor])
        && xml[cursor] !== '='
        && xml[cursor] !== '>'
        && xml[cursor] !== '/'
      ) {
        cursor += 1
      }
      const attributeName = xml.slice(nameStart, cursor)
      while (cursor < xml.length && isWhitespace(xml[cursor])) cursor += 1
      if (xml[cursor] !== '=') break
      cursor += 1
      while (cursor < xml.length && isWhitespace(xml[cursor])) cursor += 1
      const quote = xml[cursor]
      if (quote !== '"' && quote !== "'") break
      cursor += 1
      const valueStart = cursor
      while (cursor < xml.length && xml[cursor] !== quote) cursor += 1
      const value = xml.slice(valueStart, cursor)
      if (cursor < xml.length) cursor += 1
      if (attributeName === 'xmlns') namespaces.set('', value)
      else if (attributeName.startsWith('xmlns:')) namespaces.set(attributeName.slice('xmlns:'.length), value)
      else attributes.push({ name: attributeName, value })
    }
    scopes.push(namespaces)
    visit(tagName, attributes, namespaces)
    if (selfClosing) scopes.pop()
    index = cursor
  }
}

function skipPast(xml: string, terminator: string, from: number): number {
  const end = xml.indexOf(terminator, from)
  return end < 0 ? xml.length : end + terminator.length
}

function skipDeclaration(xml: string, from: number): number {
  let depth = 0
  let cursor = from
  while (cursor < xml.length) {
    const character = xml[cursor]
    if (character === '[') depth += 1
    else if (character === ']') depth -= 1
    else if (character === '>' && depth === 0) return cursor + 1
    cursor += 1
  }
  return xml.length
}

function isWhitespace(character: string | undefined): boolean {
  return character === ' ' || character === '\t' || character === '\n' || character === '\r'
}
