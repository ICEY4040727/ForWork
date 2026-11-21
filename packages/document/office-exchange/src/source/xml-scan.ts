/**
 * Byte-level XML scanning primitives shared by the span scanner, the splice
 * planner, and the providers' in-place writers.
 *
 * Every offset is a source byte offset, so a BOM, CRLF pair, or multibyte
 * character advances by its encoded length. The scanners are quote aware and
 * skip comments, CDATA sections, processing instructions, and declarations;
 * they do not resolve namespaces or interpret element meaning.
 * @module @forwork/office-exchange/source/xml-scan
 */

/** Opening sequence of an XML comment. */
export const COMMENT_OPEN = [0x3C, 0x21, 0x2D, 0x2D] as const

/** Closing sequence of an XML comment. */
export const COMMENT_CLOSE = [0x2D, 0x2D, 0x3E] as const

/** Opening sequence of a CDATA section. */
export const CDATA_OPEN = [0x3C, 0x21, 0x5B, 0x43, 0x44, 0x41, 0x54, 0x41, 0x5B] as const

/** Closing sequence of a CDATA section. */
export const CDATA_CLOSE = [0x5D, 0x5D, 0x3E] as const

/** Closing sequence of a processing instruction. */
export const PROCESSING_CLOSE = [0x3F, 0x3E] as const

const UTF8_DECODER = new TextDecoder()

const TAB = 0x09
const LINE_FEED = 0x0A
const CARRIAGE_RETURN = 0x0D
const SPACE = 0x20
const QUOTE = 0x22
const APOSTROPHE = 0x27
const SLASH = 0x2F
const EQUALS = 0x3D
const GREATER_THAN = 0x3E
const OPEN_BRACKET = 0x5B
const CLOSE_BRACKET = 0x5D

/** One attribute parsed from a start tag, with its byte range in the part. */
export interface ParsedAttribute {
  /** Qualified name as written in the source. */
  readonly qname: string
  /** Decoded attribute value without its quotes. */
  readonly value: string
  /** Inclusive byte offset of the attribute name. */
  readonly start: number
  /** Exclusive byte offset immediately after the closing quote. */
  readonly end: number
}

/** One start tag parsed from a part's bytes. */
export interface ParsedStartTag {
  /** Qualified element name as written in the source. */
  readonly qname: string
  /** Attributes in source order. */
  readonly attributes: readonly ParsedAttribute[]
  /** Whether the tag ends in `/>`. */
  readonly selfClosing: boolean
}

/**
 * Decode a UTF-8 byte range.
 * @param bytes - Source bytes.
 * @param start - Inclusive byte offset; defaults to the first byte.
 * @param end - Exclusive byte offset; defaults to the buffer length.
 * @returns The decoded text.
 */
export function decodeBytes(bytes: Uint8Array, start = 0, end = bytes.byteLength): string {
  return UTF8_DECODER.decode(bytes.subarray(start, end))
}

/**
 * Report whether a byte is XML whitespace.
 * @param byte - Byte to test, possibly past the end of the buffer.
 * @returns `true` for space, tab, line feed, and carriage return.
 */
export function isWhitespace(byte: number | undefined): boolean {
  return byte === SPACE || byte === TAB || byte === LINE_FEED || byte === CARRIAGE_RETURN
}

/**
 * Match a byte sequence at one offset.
 * @param bytes - Source bytes.
 * @param start - Offset of the first candidate byte.
 * @param sequence - Bytes to match.
 * @returns `true` when every sequence byte matches at {@link start}.
 */
export function startsWith(bytes: Uint8Array, start: number, sequence: readonly number[]): boolean {
  if (start + sequence.length > bytes.byteLength) return false
  for (let index = 0; index < sequence.length; index += 1) {
    if (bytes[start + index] !== sequence[index]) return false
  }
  return true
}

/**
 * Locate a byte at or after an offset.
 * @param bytes - Source bytes.
 * @param target - Byte value to find.
 * @param from - Inclusive offset to start the search.
 * @returns The matching offset, or `-1` when it is absent.
 */
export function indexOfByte(bytes: Uint8Array, target: number, from: number): number {
  return bytes.indexOf(target, from)
}

/**
 * Locate a byte sequence at or after an offset.
 * @param bytes - Source bytes.
 * @param sequence - Bytes to find.
 * @param from - Inclusive offset to start the search.
 * @returns The offset where the sequence starts, or `-1` when it is absent.
 */
export function findSequence(bytes: Uint8Array, sequence: readonly number[], from: number): number {
  const last = bytes.byteLength - sequence.length
  for (let index = from; index <= last; index += 1) {
    if (startsWith(bytes, index, sequence)) return index
  }
  return -1
}

/**
 * Locate the `>` that ends a start or end tag, ignoring quoted attribute values.
 * @param bytes - Source bytes.
 * @param from - Offset after the opening `<`.
 * @returns The offset of the closing `>`, or `-1` when the tag is unterminated.
 */
export function findTagEnd(bytes: Uint8Array, from: number): number {
  let index = from
  while (index < bytes.byteLength) {
    const next = quotedAdvance(bytes, index)
    if (next === -1) return -1
    if (next !== undefined) {
      index = next
      continue
    }
    if (bytes[index] === GREATER_THAN) return index
    index += 1
  }
  return -1
}

/**
 * Locate the end of a markup declaration, ignoring quoted text and bracketed
 * internal subsets.
 * @param bytes - Source bytes.
 * @param from - Offset after `<!`.
 * @returns The offset immediately after the closing `>`, or `-1` when unterminated.
 */
export function declarationEnd(bytes: Uint8Array, from: number): number {
  let bracketDepth = 0
  let index = from
  while (index < bytes.byteLength) {
    const next = quotedAdvance(bytes, index)
    if (next === -1) return -1
    if (next !== undefined) {
      index = next
      continue
    }
    const byte = bytes[index]
    if (byte === OPEN_BRACKET) {
      bracketDepth += 1
    } else if (byte === CLOSE_BRACKET && bracketDepth > 0) {
      bracketDepth -= 1
    } else if (byte === GREATER_THAN && bracketDepth === 0) {
      return index + 1
    }
    index += 1
  }
  return -1
}

/**
 * Read an XML name that ends at whitespace, `/`, or `=`.
 * @param bytes - Source bytes.
 * @param from - Offset of the first name byte.
 * @param limit - Exclusive offset the name may not cross.
 * @returns The decoded name and its end offset, or `undefined` when no name is present.
 */
export function readName(
  bytes: Uint8Array,
  from: number,
  limit: number,
): { readonly name: string; readonly end: number } | undefined {
  let index = from
  while (index < limit) {
    const byte = bytes[index]
    if (isWhitespace(byte) || byte === SLASH || byte === EQUALS) break
    index += 1
  }
  if (index === from) return undefined
  return { name: decodeBytes(bytes, from, index), end: index }
}

/**
 * Skip XML whitespace within a limit.
 * @param bytes - Source bytes.
 * @param from - Offset to start skipping.
 * @param limit - Exclusive offset the cursor may not cross.
 * @returns The first offset at or after {@link from} that is not whitespace.
 */
export function skipWhitespace(bytes: Uint8Array, from: number, limit: number): number {
  let index = from
  while (index < limit) {
    if (!isWhitespace(bytes[index])) break
    index += 1
  }
  return index
}

/**
 * Parse one start tag whose end offset is already known.
 * @param bytes - Source bytes.
 * @param opening - Offset of the opening `<`.
 * @param tagEnd - Offset of the closing `>`.
 * @returns The element name and attributes, or `undefined` when the tag is malformed.
 */
export function parseStartTag(bytes: Uint8Array, opening: number, tagEnd: number): ParsedStartTag | undefined {
  const selfClosing = bytes[tagEnd - 1] === SLASH
  const limit = selfClosing ? tagEnd - 1 : tagEnd
  const name = readName(bytes, opening + 1, limit)
  if (name === undefined) return undefined
  const attributes: ParsedAttribute[] = []
  let index = name.end
  while (index < limit) {
    index = skipWhitespace(bytes, index, limit)
    if (index >= limit) break
    const attributeName = readName(bytes, index, limit)
    if (attributeName === undefined) return undefined
    const attributeStart = index
    index = skipWhitespace(bytes, attributeName.end, limit)
    if (bytes[index] !== EQUALS) return undefined
    index = skipWhitespace(bytes, index + 1, limit)
    const quote = bytes[index]
    if (quote !== QUOTE && quote !== APOSTROPHE) return undefined
    const close = indexOfByte(bytes, quote, index + 1)
    /* v8 ignore next -- findTagEnd already rejected a tag whose attribute quote never closes */
    if (close === -1) return undefined
    attributes.push({
      qname: attributeName.name,
      value: decodeBytes(bytes, index + 1, close),
      start: attributeStart,
      end: close + 1,
    })
    index = close + 1
  }
  return { qname: name.name, attributes, selfClosing }
}

/**
 * Read the element name from an end tag.
 * @param bytes - Source bytes.
 * @param opening - Offset of the opening `<`.
 * @param tagEnd - Offset of the closing `>`.
 * @returns The decoded name, or `undefined` when the end tag is malformed or carries content.
 */
export function readEndTagName(bytes: Uint8Array, opening: number, tagEnd: number): string | undefined {
  const name = readName(bytes, opening + 2, tagEnd)
  if (name === undefined) return undefined
  return skipWhitespace(bytes, name.end, tagEnd) === tagEnd ? name.name : undefined
}

/**
 * Split a qualified name into its prefix and local name.
 * @param qname - Qualified name as written in the source.
 * @returns Both parts, or `undefined` when the name has an empty or repeated separator.
 */
export function splitQName(qname: string): { readonly prefix: string; readonly localName: string } | undefined {
  const separator = qname.indexOf(':')
  if (separator === -1) return { prefix: '', localName: qname }
  if (separator === 0 || separator === qname.length - 1) return undefined
  const prefix = qname.slice(0, separator)
  const localName = qname.slice(separator + 1)
  return localName.includes(':') ? undefined : { prefix, localName }
}

/**
 * Advance past a quoted run that starts at the current offset.
 * @param bytes - Source bytes.
 * @param index - Offset to inspect.
 * @returns The offset after the closing quote; `-1` when the quote is unterminated;
 * `undefined` when no quote starts at {@link index}.
 */
function quotedAdvance(bytes: Uint8Array, index: number): number | undefined {
  const quote = bytes[index]
  if (quote !== QUOTE && quote !== APOSTROPHE) return undefined
  const close = indexOfByte(bytes, quote, index + 1)
  return close === -1 ? -1 : close + 1
}
