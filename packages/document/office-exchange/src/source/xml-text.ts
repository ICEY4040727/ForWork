/**
 * Decoded-XML string scanning shared by the in-place writers.
 *
 * The providers scan already-decoded fragments when they plan a splice; these
 * helpers keep tag and declaration boundaries quote aware without coupling the
 * scan to one provider's fragment parser.
 * @module @forwork/office-exchange/source/xml-text
 */

/**
 * Report whether a character is XML whitespace.
 * @param character - Character to test, possibly past the end of the text.
 * @returns `true` for space, tab, line feed, and carriage return.
 */
export function isXmlTextWhitespace(character: string | undefined): boolean {
  return character === ' ' || character === '\t' || character === '\n' || character === '\r'
}

/**
 * Locate the `>` that ends a start or end tag, ignoring quoted attribute values.
 * @param text - Decoded fragment.
 * @param from - Offset after the opening `<`.
 * @returns The offset of the closing `>`, or `-1` when the tag is unterminated.
 */
export function tagEndInText(text: string, from: number): number {
  let index = from
  while (index < text.length) {
    const next = quotedAdvance(text, index)
    if (next === -1) return -1
    if (next !== undefined) {
      index = next
      continue
    }
    if (text[index] === '>') return index
    index += 1
  }
  return -1
}

/**
 * Locate the end of a markup declaration, ignoring quoted text and bracketed
 * internal subsets.
 * @param text - Decoded fragment.
 * @param from - Offset after `<!`.
 * @returns The offset immediately after the closing `>`, or `-1` when unterminated.
 */
export function declarationEndInText(text: string, from: number): number {
  let bracketDepth = 0
  let index = from
  while (index < text.length) {
    const next = quotedAdvance(text, index)
    if (next === -1) return -1
    if (next !== undefined) {
      index = next
      continue
    }
    const character = text[index]
    if (character === '[') {
      bracketDepth += 1
    } else if (character === ']') {
      bracketDepth -= 1
    } else if (character === '>' && bracketDepth <= 0) {
      return index + 1
    }
    index += 1
  }
  return -1
}

/**
 * Advance past a quoted run that starts at the current offset.
 * @param text - Decoded fragment.
 * @param index - Offset to inspect.
 * @returns The offset after the closing quote; `-1` when the quote is unterminated;
 * `undefined` when no quote starts at {@link index}.
 */
function quotedAdvance(text: string, index: number): number | undefined {
  const quote = text[index]
  if (quote !== '"' && quote !== "'") return undefined
  const close = text.indexOf(quote, index + 1)
  return close === -1 ? -1 : close + 1
}
