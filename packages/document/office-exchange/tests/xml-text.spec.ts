/** Quote-aware decoded-XML scanning shared by the in-place writers. */
import { describe, expect, it } from 'vitest'
import { declarationEndInText, isXmlTextWhitespace, tagEndInText } from '../src/source/xml-text.ts'

describe('isXmlTextWhitespace', () => {
  it('accepts the four XML whitespace characters', () => {
    for (const character of [' ', '\t', '\n', '\r']) {
      expect(isXmlTextWhitespace(character)).toBe(true)
    }
  })

  it('rejects other characters and the end of the text', () => {
    expect(isXmlTextWhitespace('a')).toBe(false)
    expect(isXmlTextWhitespace(undefined)).toBe(false)
  })
})

describe('tagEndInText', () => {
  it('returns the offset of the closing angle bracket', () => {
    expect(tagEndInText('<row r="1">', 1)).toBe(10)
  })

  it('ignores angle brackets inside quoted attribute values', () => {
    expect(tagEndInText('<row a="x>y" b=\'<\'>', 1)).toBe(18)
  })

  it('reports an unterminated attribute quote', () => {
    expect(tagEndInText('<row a="x>', 1)).toBe(-1)
  })

  it('reports a tag that never closes', () => {
    expect(tagEndInText('<row', 1)).toBe(-1)
  })
})

describe('declarationEndInText', () => {
  it('returns the offset immediately after a plain declaration', () => {
    const text = '<!DOCTYPE worksheet>'
    expect(declarationEndInText(text, 2)).toBe(text.length)
  })

  it('skips quoted text inside the declaration', () => {
    const text = '<!DOCTYPE x SYSTEM "a>b">'
    expect(declarationEndInText(text, 2)).toBe(text.length)
  })

  it('skips angle brackets inside a bracketed internal subset', () => {
    const text = '<!DOCTYPE x [ <!ENTITY e "a>b"> ]>'
    expect(declarationEndInText(text, 2)).toBe(text.length)
  })

  it('tracks a closing bracket without an internal subset', () => {
    const text = '<!x ] >'
    expect(declarationEndInText(text, 2)).toBe(text.length)
  })

  it('reports an unterminated quoted value', () => {
    expect(declarationEndInText('<!DOCTYPE x SYSTEM "open', 2)).toBe(-1)
  })

  it('reports a declaration that never closes', () => {
    expect(declarationEndInText('<!DOCTYPE x [ <!ENTITY e "x"> ]', 2)).toBe(-1)
  })
})
