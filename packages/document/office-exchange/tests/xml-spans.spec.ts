/** Byte-span scanner behavior over namespace-aware OOXML fragments. */
import { describe, expect, it } from 'vitest'
import type { ElementSpan, SpanKind, SpanScanResult } from '../src/source/xml-spans.ts'
import { locateSpan, scanSpans } from '../src/source/xml-spans.ts'

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const STRICT_WORD_NS = 'http://purl.oclc.org/ooxml/wordprocessingml/main'
const SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const STRICT_SHEET_NS = 'http://purl.oclc.org/ooxml/spreadsheetml/main'
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml'
const PART = 'word/document.xml'
const ALL_KINDS: readonly SpanKind[] = ['paragraph', 'table', 'tableRow', 'tableCell', 'sheetCell']

const WORD_ROOT = `xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}" xmlns:mc="${MC_NS}"`

function bytesOf(text: string): Uint8Array {
  return new TextEncoder().encode(text)
}

function scan(xml: string, kinds: readonly SpanKind[] = ALL_KINDS): SpanScanResult {
  return scanSpans(bytesOf(xml), PART, kinds)
}

function byteIndexOf(xml: string, index: number): number {
  return bytesOf(xml.slice(0, index)).length
}

function nthIndexOf(xml: string, needle: string, occurrence: number): number {
  let index = -1
  for (let count = 0; count <= occurrence; count += 1) {
    index = xml.indexOf(needle, index + 1)
    if (index === -1) throw new Error(`missing occurrence ${String(occurrence)} of ${needle}`)
  }
  return index
}

function elementOffsets(
  xml: string,
  open: string,
  close: string,
  openOccurrence = 0,
  closeOccurrence = openOccurrence,
): { start: number; end: number } {
  const start = nthIndexOf(xml, open, openOccurrence)
  const end = nthIndexOf(xml, close, closeOccurrence)
  return { start: byteIndexOf(xml, start), end: byteIndexOf(xml, end + close.length) }
}

function spansOfKind(result: SpanScanResult, kind: SpanKind): readonly ElementSpan[] {
  return result.spans.filter(span => span.kind === kind)
}

function envelope(body: string): string {
  return `<w:document ${WORD_ROOT}><w:body>${body}</w:body></w:document>`
}

describe('scanSpans', () => {
  it('records paragraphs, tables, rows, and cells with source byte offsets', () => {
    const xml = envelope('<w:p><w:r><w:t>x</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>')
    const result = scan(xml)
    expect(result.failure).toBeUndefined()
    expect(result.spans.map(span => span.kind)).toEqual([
      'paragraph',
      'table',
      'tableRow',
      'tableCell',
      'paragraph',
    ])
    const paragraph = elementOffsets(xml, '<w:p>', '</w:p>')
    expect(result.spans[0]).toEqual({ kind: 'paragraph', part: PART, ...paragraph })
    const table = elementOffsets(xml, '<w:tbl>', '</w:tbl>')
    const tableSpan = spansOfKind(result, 'table')[0]
    const rowSpan = spansOfKind(result, 'tableRow')[0]
    const cellSpan = spansOfKind(result, 'tableCell')[0]
    expect(tableSpan).toEqual({ kind: 'table', part: PART, ...table })
    expect(rowSpan?.start).toBeGreaterThan(tableSpan?.start ?? -1)
    expect(rowSpan?.end).toBeLessThan(tableSpan?.end ?? -1)
    expect(cellSpan?.start).toBeGreaterThan(rowSpan?.start ?? -1)
    const selfClosing = xml.indexOf('<w:p/>')
    expect(result.spans[4]).toEqual({
      kind: 'paragraph',
      part: PART,
      start: byteIndexOf(xml, selfClosing),
      end: byteIndexOf(xml, selfClosing + '<w:p/>'.length),
    })
  })

  it('recognizes any prefix bound to the WordprocessingML namespace, including the strict URI', () => {
    const xml = `<root xmlns:a="${WORD_NS}" xmlns:b="${STRICT_WORD_NS}"><a:p>one</a:p><b:p>two</b:p></root>`
    const result = scan(xml)
    expect(result.failure).toBeUndefined()
    expect(result.spans).toEqual([
      { kind: 'paragraph', part: PART, ...elementOffsets(xml, '<a:p>', '</a:p>') },
      { kind: 'paragraph', part: PART, ...elementOffsets(xml, '<b:p>', '</b:p>') },
    ])
  })

  it('resolves an unprefixed element through the default namespace', () => {
    const xml = `<p xmlns="${WORD_NS}">text</p>`
    expect(scan(xml).spans).toEqual([
      { kind: 'paragraph', part: PART, start: 0, end: bytesOf(xml).length },
    ])
  })

  it('leaves an unprefixed element in no namespace unmatched', () => {
    const xml = '<root><p>text</p></root>'
    expect(scan(xml)).toEqual({ spans: [] })
  })

  it('honors prefix rebinding for nested content and restores the outer binding', () => {
    const xml = `<root ${WORD_ROOT}><w:p><w:p xmlns:w="urn:other">ignored</w:p></w:p><w:p/></root>`
    const result = scan(xml)
    expect(result.failure).toBeUndefined()
    expect(result.spans.map(span => span.kind)).toEqual(['paragraph', 'paragraph'])
    const outer = elementOffsets(xml, '<w:p>', '</w:p>', 0, 1)
    expect(result.spans[0]).toEqual({ kind: 'paragraph', part: PART, ...outer })
    const rebound = xml.indexOf('<w:p/>')
    expect(result.spans[1]).toEqual({
      kind: 'paragraph',
      part: PART,
      start: byteIndexOf(xml, rebound),
      end: byteIndexOf(xml, rebound + '<w:p/>'.length),
    })
  })

  it('skips comments, CDATA, processing instructions, and DOCTYPE declarations', () => {
    const xml = '<!DOCTYPE w:document [ <!ENTITY marker "a>b"> <!ELEMENT w:p ANY> ]>'
      + '<?target <w:p>ignored</w:p> ?>'
      + `<root ${WORD_ROOT}>`
      + '<!-- <w:p>ignored</w:p> -->'
      + '<w:p><w:r><w:t><![CDATA[<w:p>ignored</w:p>]]></w:t></w:r></w:p>'
      + '</root>'
    const result = scan(xml)
    expect(result.failure).toBeUndefined()
    expect(result.spans).toHaveLength(1)
    const paragraph = elementOffsets(xml, '<w:p>', '</w:p>', 2, 3)
    expect(result.spans[0]).toEqual({ kind: 'paragraph', part: PART, ...paragraph })
  })

  it('handles a declaration with a closing bracket outside the internal subset', () => {
    const xml = `<!DOCTYPE root ]><root ${WORD_ROOT}><w:p/></root>`
    expect(scan(xml).spans).toHaveLength(1)
  })

  it('keeps attribute values containing > and single-quoted values', () => {
    const xml = envelope('<w:p w:rsidR="a>b" w:rsidRDefault=\'c>d\'><w:r><w:t>x</w:t></w:r></w:p>')
    const result = scan(xml)
    expect(result.failure).toBeUndefined()
    expect(result.spans[0]).toEqual({ kind: 'paragraph', part: PART, ...elementOffsets(xml, '<w:p ', '</w:p>') })
  })

  it('records self-closing elements and tolerates whitespace before the closing bracket', () => {
    const xml = envelope('<w:p/><w:p ></w:p><w:p w:rsidR="1" />')
    const result = scan(xml)
    expect(result.failure).toBeUndefined()
    expect(result.spans).toHaveLength(3)
    const spaced = xml.indexOf('<w:p >')
    expect(result.spans[1]).toEqual({
      kind: 'paragraph',
      part: PART,
      start: byteIndexOf(xml, spaced),
      end: byteIndexOf(xml, xml.indexOf('</w:p>') + '</w:p>'.length),
    })
    const attributed = xml.indexOf('<w:p w:rsidR="1" />')
    expect(result.spans[2]).toEqual({
      kind: 'paragraph',
      part: PART,
      start: byteIndexOf(xml, attributed),
      end: byteIndexOf(xml, attributed + '<w:p w:rsidR="1" />'.length),
    })
  })

  it('preserves byte offsets across a BOM, CRLF pairs, and multibyte text', () => {
    const xml = `\uFEFF<w:document ${WORD_ROOT}>\r\n<w:body>\r\n<w:p><w:r><w:t>文本</w:t></w:r></w:p>\r\n<w:p/>\r\n</w:body>\r\n</w:document>`
    const result = scan(xml)
    expect(result.failure).toBeUndefined()
    const first = elementOffsets(xml, '<w:p>', '</w:p>')
    expect(result.spans[0]).toEqual({ kind: 'paragraph', part: PART, ...first })
    expect(result.spans[0]?.start).toBeGreaterThan(3)
    const second = xml.indexOf('<w:p/>')
    expect(result.spans[1]?.start).toBe(byteIndexOf(xml, second))
    expect(result.spans[1]?.end).toBe(byteIndexOf(xml, second + '<w:p/>'.length))
  })

  it('records nested same-name tables, rows, and cells without truncating the outer span', () => {
    const xml = envelope('<w:tbl><w:tr><w:tc><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>')
    const result = scan(xml)
    expect(result.failure).toBeUndefined()
    const tables = spansOfKind(result, 'table')
    const rows = spansOfKind(result, 'tableRow')
    const cells = spansOfKind(result, 'tableCell')
    expect(tables).toHaveLength(2)
    expect(rows).toHaveLength(2)
    expect(cells).toHaveLength(2)
    expect(tables[0]).toEqual({ kind: 'table', part: PART, ...elementOffsets(xml, '<w:tbl>', '</w:tbl>', 0, 1) })
    expect(tables[1]).toEqual({ kind: 'table', part: PART, ...elementOffsets(xml, '<w:tbl>', '</w:tbl>', 1, 0) })
    expect(tables[0]?.start).toBeLessThan(tables[1]?.start ?? -1)
    expect(tables[0]?.end).toBeGreaterThan(tables[1]?.end ?? -1)
    expect(cells[0]?.end).toBeGreaterThan(tables[1]?.end ?? -1)
  })

  it('records spreadsheet cells in the SpreadsheetML namespace', () => {
    const xml = `<worksheet xmlns="${SHEET_NS}"><sheetData><row r="1"><c r="A1" t="n"><v>1</v></c><c r="B1"/></row></sheetData></worksheet>`
    const result = scan(xml)
    expect(result.failure).toBeUndefined()
    const cells = spansOfKind(result, 'sheetCell')
    expect(cells).toHaveLength(2)
    expect(cells[0]).toEqual({
      kind: 'sheetCell',
      part: PART,
      start: byteIndexOf(xml, xml.indexOf('<c r="A1"')),
      end: byteIndexOf(xml, xml.indexOf('</c>') + '</c>'.length),
      identity: { kind: 'sheetCell', reference: 'A1' },
    })
    expect(cells[1]?.identity).toEqual({ kind: 'sheetCell', reference: 'B1' })
  })

  it('recognizes the strict SpreadsheetML namespace', () => {
    const xml = `<worksheet xmlns="${STRICT_SHEET_NS}"><c r="A1"/></worksheet>`
    expect(scan(xml).spans).toEqual([{
      kind: 'sheetCell',
      part: PART,
      start: byteIndexOf(xml, xml.indexOf('<c')),
      end: bytesOf(xml).length - '</worksheet>'.length,
      identity: { kind: 'sheetCell', reference: 'A1' },
    }])
  })

  it('refuses mc:AlternateContent instead of choosing a branch', () => {
    const xml = `<root ${WORD_ROOT}><mc:AlternateContent><mc:Choice Requires="w"><w:p/></mc:Choice><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent></root>`
    const result = scan(xml)
    expect(result.spans).toEqual([])
    expect(result.failure?.code).toBe('ambiguous-span')
    expect(result.failure?.offset).toBe(byteIndexOf(xml, xml.indexOf('<mc:AlternateContent')))
  })

  it('does not treat an AlternateContent element in another namespace as ambiguous', () => {
    const xml = `<root ${WORD_ROOT} xmlns:x="urn:other"><x:AlternateContent><w:p/></x:AlternateContent></root>`
    const result = scan(xml)
    expect(result.failure).toBeUndefined()
    expect(result.spans).toHaveLength(1)
  })

  it('returns no spans and no failure for an empty kind request', () => {
    const result = scan(envelope('<w:p/>'), [])
    expect(result).toEqual({ spans: [] })
  })

  it('returns no spans for a part with no markup', () => {
    expect(scan('plain text only')).toEqual({ spans: [] })
  })
})

describe('scanSpans refusals', () => {
  const malformed: readonly { readonly name: string; readonly xml: string }[] = [
    { name: 'an unbound element prefix', xml: `<root ${WORD_ROOT.replace(`xmlns:w="${WORD_NS}"`, '')}><w:p/></root>` },
    { name: 'an unbound attribute prefix', xml: `<root ${WORD_ROOT}><w:p bad:id="1"/></root>` },
    { name: 'a mismatched end tag', xml: `<root ${WORD_ROOT}><w:p></w:tbl></root>` },
    { name: 'an end tag without a start tag', xml: `</w:p><root ${WORD_ROOT}/>` },
    { name: 'an unclosed element', xml: `<root ${WORD_ROOT}><w:p>` },
    { name: 'an empty start-tag name', xml: `<root ${WORD_ROOT}><></root>` },
    { name: 'an attribute without a value', xml: `<root ${WORD_ROOT}><w:p w:rsidR></root>` },
    { name: 'an unquoted attribute value', xml: `<root ${WORD_ROOT}><w:p w:rsidR=value></root>` },
    { name: 'an attribute name without a name', xml: `<root ${WORD_ROOT}><w:p =value></root>` },
    { name: 'an end tag with trailing content', xml: `<root ${WORD_ROOT}><w:p></w:p extra></root>` },
    { name: 'an empty end-tag name', xml: `<root ${WORD_ROOT}></></root>` },
    { name: 'an unterminated element tag', xml: `<root ${WORD_ROOT}><w:p` },
    { name: 'an unterminated quoted attribute', xml: `<root ${WORD_ROOT}><w:p w:rsidR="x>` },
    { name: 'an unterminated comment', xml: `<root ${WORD_ROOT}><!-- x` },
    { name: 'an unterminated CDATA section', xml: `<root ${WORD_ROOT}><![CDATA[x` },
    { name: 'an unterminated processing instruction', xml: `<root ${WORD_ROOT}><?pi x` },
    { name: 'an unterminated DOCTYPE declaration', xml: `<root ${WORD_ROOT}><!DOCTYPE x` },
    { name: 'an unterminated quoted DOCTYPE value', xml: `<root ${WORD_ROOT}><!DOCTYPE x "y>` },
    { name: 'an empty namespace declaration prefix', xml: `<root ${WORD_ROOT}><w:p xmlns:="urn:x"/></root>` },
    { name: 'a leading colon in an element name', xml: `<root ${WORD_ROOT}><:p/></root>` },
    { name: 'a trailing colon in an element name', xml: `<root ${WORD_ROOT}><w:/></root>` },
    { name: 'a second colon in an element name', xml: `<root ${WORD_ROOT}><w:a:b/></root>` },
    { name: 'a leading colon in an attribute name', xml: `<root ${WORD_ROOT}><w:p :id="1"/></root>` },
    { name: 'a repeated paragraph identity attribute', xml: `<root ${WORD_ROOT}><w:p w14:paraId="1" w14:paraId="2"/></root>` },
    { name: 'a repeated sheet-cell reference', xml: `<root xmlns:s="${SHEET_NS}"><s:c r="A1" r="A2"/></root>` },
  ]

  for (const fixture of malformed) {
    it(`fails closed on ${fixture.name}`, () => {
      const result = scan(fixture.xml)
      expect(result.spans).toEqual([])
      expect(result.failure?.code).toBe('ambiguous-span')
      expect(result.failure?.part).toBe(PART)
      expect(result.failure?.detail.length).toBeGreaterThan(0)
    })
  }

  it('rejects a UTF-16 little-endian byte-order mark', () => {
    const result = scanSpans(new Uint8Array([0xFF, 0xFE, 0x3C, 0x00]), PART, ALL_KINDS)
    expect(result.failure?.offset).toBe(0)
  })

  it('rejects a UTF-16 big-endian byte-order mark', () => {
    const result = scanSpans(new Uint8Array([0xFE, 0xFF, 0x00, 0x3C]), PART, ALL_KINDS)
    expect(result.failure?.code).toBe('ambiguous-span')
  })
})

describe('locateSpan', () => {
  it('finds the only paragraph with a given w14:paraId', () => {
    const xml = `<root ${WORD_ROOT}><w:p w14:paraId="1A2B"><w:r><w:t>x</w:t></w:r></w:p><w:p w14:paraId="3C4D"/></root>`
    const result = scan(xml)
    const paragraph = elementOffsets(xml, '<w:p w14:paraId="1A2B"', '</w:p>')
    expect(locateSpan(result.spans, { kind: 'paragraph', paraId: '1A2B' })).toEqual({
      kind: 'paragraph',
      part: PART,
      ...paragraph,
      identity: { kind: 'paragraph', paraId: '1A2B' },
    })
  })

  it('finds a paragraph identity declared on the element and tolerates xml:space', () => {
    const xml = `<w:p xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}" w14:paraId="AA" xml:space="preserve"/>`
    const result = scan(xml)
    expect(locateSpan(result.spans, { kind: 'paragraph', paraId: 'AA' })?.identity).toEqual({
      kind: 'paragraph',
      paraId: 'AA',
    })
  })

  it('does not assert an identity for an unprefixed or foreign paraId', () => {
    const xml = `<root ${WORD_ROOT} xmlns:x="urn:other"><w:p paraId="1"/><w:p x:paraId="2"/></root>`
    const result = scan(xml)
    expect(result.spans.every(span => span.identity === undefined)).toBe(true)
    expect(locateSpan(result.spans, { kind: 'paragraph', paraId: '1' })).toBeUndefined()
  })

  it('returns undefined when a paragraph identity repeats', () => {
    const xml = `<root ${WORD_ROOT}><w:p w14:paraId="AB"/><w:p w14:paraId="AB"/></root>`
    const result = scan(xml)
    expect(result.spans).toHaveLength(2)
    expect(locateSpan(result.spans, { kind: 'paragraph', paraId: 'AB' })).toBeUndefined()
  })

  it('finds a sheet cell by reference and refuses an absent or cross-kind identity', () => {
    const xml = `<worksheet xmlns="${SHEET_NS}"><c r="B2"/><c/></worksheet>`
    const result = scan(xml)
    expect(locateSpan(result.spans, { kind: 'sheetCell', reference: 'B2' })?.identity).toEqual({
      kind: 'sheetCell',
      reference: 'B2',
    })
    expect(locateSpan(result.spans, { kind: 'sheetCell', reference: 'C3' })).toBeUndefined()
    expect(locateSpan(result.spans, { kind: 'paragraph', paraId: 'B2' })).toBeUndefined()
  })

  it('ignores spans without identity and spans of another kind', () => {
    const xml = envelope('<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>')
    const result = scan(xml)
    expect(locateSpan(result.spans, { kind: 'paragraph', paraId: 'missing' })).toBeUndefined()
    expect(locateSpan([], { kind: 'paragraph', paraId: 'missing' })).toBeUndefined()
  })
})
