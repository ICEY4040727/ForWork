/**
 * Namespace-aware byte scanner for the OOXML parts an in-place save splices.
 *
 * The scanner works directly on a part's UTF-8 bytes, so every recorded offset
 * is a source byte offset: a BOM, CRLF pairs, and multibyte characters advance
 * the offset by their encoded length rather than by a normalized character
 * count. It is single pass, quote aware, and skips comments, CDATA sections,
 * processing instructions, and DOCTYPE declarations. Elements are matched by
 * namespace URI, not by a literal prefix, so any prefix bound to the expected
 * OOXML namespace is recognized.
 *
 * The scanner fails closed. An unbound prefix, a malformed tag, an unbalanced
 * element, or an `mc:AlternateContent` container produces an `ambiguous-span`
 * failure and an empty span list; callers must never splice from a partial
 * scan.
 * @module @forwork/office-exchange/source/xml-spans
 */

import {
  CDATA_CLOSE,
  CDATA_OPEN,
  COMMENT_CLOSE,
  COMMENT_OPEN,
  PROCESSING_CLOSE,
  declarationEnd,
  findSequence,
  findTagEnd,
  indexOfByte,
  parseStartTag,
  readEndTagName,
  splitQName,
  startsWith,
} from './xml-scan.ts'
import type { ParsedStartTag } from './xml-scan.ts'

/** Editable unit kinds and opaque preserved containers the scanner resolves. */
export type SpanKind =
  | 'paragraph'
  | 'table'
  | 'tableRow'
  | 'tableCell'
  | 'sheetCell'
  | 'contentControl'
  | 'preserved'

/**
 * An element identity that a source part states explicitly. A paragraph may
 * carry `w14:paraId`; a worksheet cell may carry its `r` reference. Other unit
 * kinds have no source-stated identity and must be paired by structure.
 */
export type SpanIdentity =
  | { readonly kind: 'paragraph'; readonly paraId: string }
  | { readonly kind: 'sheetCell'; readonly reference: string }

/** One scanned element span in the part's source bytes. */
export interface ElementSpan {
  /** Unit kind of the matched element. */
  readonly kind: SpanKind
  /** Decoded part name the bytes belong to. */
  readonly part: string
  /** Inclusive byte offset of the opening `<`. */
  readonly start: number
  /** Exclusive byte offset immediately after the element's last `>`. */
  readonly end: number
  /** Identity stated by the source element, when it states one. */
  readonly identity?: SpanIdentity
}

/** A scanner refusal. Spans are never returned together with a failure. */
export interface SpanFailure {
  /** Stable classification: the scan cannot locate an unambiguous span. */
  readonly code: 'ambiguous-span'
  /** Decoded part name being scanned. */
  readonly part: string
  /** Byte offset where the ambiguity was detected. */
  readonly offset: number
  /** Actionable description naming the construct at {@link offset}. */
  readonly detail: string
}

/** The outcome of scanning one part. */
export interface SpanScanResult {
  /** Spans sorted by ascending start offset; empty whenever {@link failure} is present. */
  readonly spans: readonly ElementSpan[]
  /** Present when the scanner refused to locate spans. */
  readonly failure?: SpanFailure
}

interface OpenElement {
  readonly qname: string
  readonly start: number
  readonly kind: SpanKind | undefined
  readonly identity: SpanIdentity | undefined
  readonly bindings: ReadonlyMap<string, string> | undefined
}

interface ResolvedStartTag {
  readonly kind: SpanKind | undefined
  readonly identity: SpanIdentity | undefined
  readonly bindings: ReadonlyMap<string, string> | undefined
}

type ResolveStartTagResult =
  | { readonly ok: true; readonly value: ResolvedStartTag }
  | { readonly ok: false; readonly detail: string }

interface KindTarget {
  readonly localNames: readonly string[]
  readonly namespaces: ReadonlySet<string>
}

const WORD_NAMESPACES: ReadonlySet<string> = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://purl.oclc.org/ooxml/wordprocessingml/main',
])

const SHEET_NAMESPACES: ReadonlySet<string> = new Set([
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://purl.oclc.org/ooxml/spreadsheetml/main',
])

const MARKUP_COMPATIBILITY_NAMESPACES: ReadonlySet<string> = new Set([
  'http://schemas.openxmlformats.org/markup-compatibility/2006',
])

const WORD_2010_NAMESPACES: ReadonlySet<string> = new Set([
  'http://schemas.microsoft.com/office/word/2010/wordml',
  'http://schemas.microsoft.com/office/word/2012/wordml',
])

const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace'

const KIND_TARGETS: Readonly<Record<SpanKind, KindTarget>> = {
  paragraph: { localNames: ['p'], namespaces: WORD_NAMESPACES },
  table: { localNames: ['tbl'], namespaces: WORD_NAMESPACES },
  tableRow: { localNames: ['tr'], namespaces: WORD_NAMESPACES },
  tableCell: { localNames: ['tc'], namespaces: WORD_NAMESPACES },
  sheetCell: { localNames: ['c'], namespaces: SHEET_NAMESPACES },
  contentControl: { localNames: ['sdt'], namespaces: WORD_NAMESPACES },
  preserved: { localNames: ['drawing', 'pict', 'object', 'fldSimple'], namespaces: WORD_NAMESPACES },
}

const QUESTION = 0x3F
const BANG = 0x21

/**
 * Scan one part for the requested element kinds.
 * @param bytes - Uncompressed part bytes, assumed to be UTF-8 XML.
 * @param part - Decoded part name copied onto every span and failure.
 * @param kinds - Unit kinds to record; other elements are still parsed for balance and namespace scope.
 * @returns Spans sorted by start offset, or an `ambiguous-span` failure with no spans.
 */
export function scanSpans(bytes: Uint8Array, part: string, kinds: readonly SpanKind[]): SpanScanResult {
  if (isUtf16Bom(bytes)) {
    return failed(part, 0, 'part starts with a UTF-16 byte-order mark; only UTF-8 source bytes are supported')
  }
  const requested: ReadonlySet<SpanKind> = new Set(kinds)
  const spans: ElementSpan[] = []
  const frames: OpenElement[] = []
  let index = 0
  while (index < bytes.length) {
    const opening = indexOfByte(bytes, 0x3C, index)
    if (opening === -1) break
    if (startsWith(bytes, opening, COMMENT_OPEN)) {
      const close = findSequence(bytes, COMMENT_CLOSE, opening + COMMENT_OPEN.length)
      if (close === -1) return failed(part, opening, 'unterminated XML comment')
      index = close + COMMENT_CLOSE.length
      continue
    }
    if (startsWith(bytes, opening, CDATA_OPEN)) {
      const close = findSequence(bytes, CDATA_CLOSE, opening + CDATA_OPEN.length)
      if (close === -1) return failed(part, opening, 'unterminated CDATA section')
      index = close + CDATA_CLOSE.length
      continue
    }
    if (bytes[opening + 1] === QUESTION) {
      const close = findSequence(bytes, PROCESSING_CLOSE, opening + 2)
      if (close === -1) return failed(part, opening, 'unterminated processing instruction')
      index = close + PROCESSING_CLOSE.length
      continue
    }
    if (bytes[opening + 1] === BANG) {
      const close = declarationEnd(bytes, opening + 2)
      if (close === -1) return failed(part, opening, 'unterminated markup declaration')
      index = close
      continue
    }
    const tagEnd = findTagEnd(bytes, opening + 1)
    if (tagEnd === -1) return failed(part, opening, 'unterminated element tag')
    if (bytes[opening + 1] === 0x2F) {
      const endName = readEndTagName(bytes, opening, tagEnd)
      if (endName === undefined) return failed(part, opening, 'malformed end tag')
      const frame = frames.pop()
      if (frame === undefined) return failed(part, opening, `end tag </${endName}> has no matching start tag`)
      if (frame.qname !== endName) {
        return failed(part, opening, `end tag </${endName}> does not match <${frame.qname}>`)
      }
      if (frame.kind !== undefined) {
        spans.push(spanOf(frame.kind, part, frame.start, tagEnd + 1, frame.identity))
      }
      index = tagEnd + 1
      continue
    }
    const parsed = parseStartTag(bytes, opening, tagEnd)
    if (parsed === undefined) return failed(part, opening, 'malformed start tag')
    const resolved = resolveStartTag(parsed, requested, frames)
    if (!resolved.ok) return failed(part, opening, resolved.detail)
    if (parsed.selfClosing) {
      if (resolved.value.kind !== undefined) {
        spans.push(spanOf(resolved.value.kind, part, opening, tagEnd + 1, resolved.value.identity))
      }
    } else {
      frames.push({
        qname: parsed.qname,
        start: opening,
        kind: resolved.value.kind,
        identity: resolved.value.identity,
        bindings: resolved.value.bindings,
      })
    }
    index = tagEnd + 1
  }
  if (frames.length > 0) {
    for (const frame of frames) {
      return failed(part, frame.start, `element <${frame.qname}> is not closed`)
    }
  }
  spans.sort((left, right) => left.start - right.start)
  return { spans }
}

/**
 * Locate the only span that states the given identity.
 * @param spans - Spans from one successful {@link scanSpans} call.
 * @param identity - Identity to match exactly.
 * @returns The unique match, or `undefined` when no span or more than one span states it.
 */
export function locateSpan(spans: readonly ElementSpan[], identity: SpanIdentity): ElementSpan | undefined {
  let match: ElementSpan | undefined
  for (const span of spans) {
    if (span.identity === undefined) continue
    if (!sameIdentity(span.identity, identity)) continue
    if (match !== undefined) return undefined
    match = span
  }
  return match
}

function sameIdentity(left: SpanIdentity, right: SpanIdentity): boolean {
  if (left.kind === 'paragraph' && right.kind === 'paragraph') return left.paraId === right.paraId
  if (left.kind === 'sheetCell' && right.kind === 'sheetCell') return left.reference === right.reference
  return false
}

function spanOf(
  kind: SpanKind,
  part: string,
  start: number,
  end: number,
  identity: SpanIdentity | undefined,
): ElementSpan {
  return identity === undefined ? { kind, part, start, end } : { kind, part, start, end, identity }
}

function failed(part: string, offset: number, detail: string): SpanScanResult {
  return { spans: [], failure: { code: 'ambiguous-span', part, offset, detail } }
}

function isUtf16Bom(bytes: Uint8Array): boolean {
  const first = bytes[0]
  const second = bytes[1]
  return (first === 0xFF && second === 0xFE) || (first === 0xFE && second === 0xFF)
}

function resolveNamespace(
  prefix: string,
  localBindings: ReadonlyMap<string, string> | undefined,
  frames: readonly OpenElement[],
): string | undefined {
  if (prefix === 'xml') return XML_NAMESPACE
  const local = localBindings?.get(prefix)
  if (local !== undefined) return local
  for (let index = frames.length - 1; index >= 0; index -= 1) {
    const bound = frames[index]?.bindings?.get(prefix)
    if (bound !== undefined) return bound
  }
  return prefix === '' ? '' : undefined
}

function resolveStartTag(
  parsed: ParsedStartTag,
  requested: ReadonlySet<SpanKind>,
  frames: readonly OpenElement[],
): ResolveStartTagResult {
  let bindings: Map<string, string> | undefined
  for (const attribute of parsed.attributes) {
    if (attribute.qname === 'xmlns') {
      bindings ??= new Map()
      bindings.set('', attribute.value)
      continue
    }
    if (!attribute.qname.startsWith('xmlns:')) continue
    const prefix = attribute.qname.slice('xmlns:'.length)
    if (prefix.length === 0) return { ok: false, detail: 'namespace declaration has an empty prefix' }
    bindings ??= new Map()
    bindings.set(prefix, attribute.value)
  }
  const name = splitQName(parsed.qname)
  if (name === undefined) return { ok: false, detail: `malformed element name ${parsed.qname}` }
  const namespace = resolveNamespace(name.prefix, bindings, frames)
  if (namespace === undefined) return { ok: false, detail: `unbound namespace prefix ${name.prefix}` }
  let kind: SpanKind | undefined
  for (const candidate of requested) {
    const target = KIND_TARGETS[candidate]
    if (target.localNames.includes(name.localName) && target.namespaces.has(namespace)) {
      kind = candidate
      break
    }
  }
  if (name.localName === 'AlternateContent' && MARKUP_COMPATIBILITY_NAMESPACES.has(namespace)) {
    return { ok: false, detail: 'mc:AlternateContent duplicates editable markup across alternative branches' }
  }
  let paragraphId: string | undefined
  let sheetReference: string | undefined
  for (const attribute of parsed.attributes) {
    if (attribute.qname === 'xmlns' || attribute.qname.startsWith('xmlns:')) continue
    const attributeName = splitQName(attribute.qname)
    if (attributeName === undefined) {
      return { ok: false, detail: `malformed attribute name ${attribute.qname}` }
    }
    const attributeNamespace = attributeName.prefix === ''
      ? ''
      : resolveNamespace(attributeName.prefix, bindings, frames)
    if (attributeNamespace === undefined) {
      return { ok: false, detail: `unbound namespace prefix ${attributeName.prefix}` }
    }
    if (kind === 'paragraph' && attributeName.prefix !== '' && attributeName.localName === 'paraId'
      && WORD_2010_NAMESPACES.has(attributeNamespace)) {
      if (paragraphId !== undefined) return { ok: false, detail: 'element repeats w14:paraId' }
      paragraphId = attribute.value
    }
    if (kind === 'sheetCell' && attribute.qname === 'r') {
      if (sheetReference !== undefined) return { ok: false, detail: 'element repeats its r reference' }
      sheetReference = attribute.value
    }
  }
  let identity: SpanIdentity | undefined
  if (kind === 'paragraph' && paragraphId !== undefined) {
    identity = { kind: 'paragraph', paraId: paragraphId }
  } else if (kind === 'sheetCell' && sheetReference !== undefined) {
    identity = { kind: 'sheetCell', reference: sheetReference }
  }
  return { ok: true, value: { kind, identity, bindings } }
}
