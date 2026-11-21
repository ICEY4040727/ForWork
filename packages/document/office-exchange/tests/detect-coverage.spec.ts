/** Detection branch coverage: coupled marker ordering and malformed formula/shared-string scans. */
import { strToU8, zipSync } from 'fflate'
import { expect, it } from 'vitest'
import { detectOfficePackage, detectPreserved } from '../src/detect.ts'

const DOCUMENT_XML = 'word/document.xml'
const WORKSHEET_XML = 'xl/worksheets/sheet1.xml'
const PLAIN_WORKSHEET = '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'

function packageOf(entries: Readonly<Record<string, string | Uint8Array>>): Uint8Array {
  const files: Record<string, Uint8Array> = {}
  for (const [name, content] of Object.entries(entries)) {
    files[name] = typeof content === 'string' ? strToU8(content) : content
  }
  return zipSync(files)
}

it('sorts coupled markers with the ascending part comparison first', () => {
  const header = '<w:hdr><w:p><w:fldSimple w:instr=" PAGE "/></w:p></w:hdr>'
  const documentXml = '<w:document><w:body><w:p><w:ins w:id="1"/></w:p></w:body></w:document>'
  const result = detectPreserved(packageOf({
    'word/header1.xml': header,
    [DOCUMENT_XML]: documentXml,
  }))
  expect(result.coupled.map(marker => [marker.part, marker.kind])).toEqual([
    [DOCUMENT_XML, 'tracked-change'],
    ['word/header1.xml', 'field-owned-text'],
  ])
})

it('scans shared strings with malformed item boundaries without throwing', () => {
  const richSharedString = '<sst><si><r><t>a</t></r></si><si><t>b</t></si></sst>'
  const cases = [
    '<sst><si/></sst>',
    '<sst><si',
    '<sst><si ',
    '<sst><si>',
    '<sst><si><t>b</t>',
    richSharedString,
  ] as const
  for (const shared of cases) {
    const result = detectPreserved(packageOf({ [WORKSHEET_XML]: PLAIN_WORKSHEET, 'xl/sharedStrings.xml': shared }))
    expect(result.preserved).toEqual(new Set())
  }
  expect(detectPreserved(packageOf({
    [WORKSHEET_XML]: PLAIN_WORKSHEET,
    'xl/sharedStrings.xml': richSharedString,
  })).coupled.map(marker => marker.kind)).toEqual(['rich-text-shared-string'])
})

it('scans shared strings whose item starts exactly at the first run', () => {
  const shared = '<sst><si><r><t>a</t></r></si><si><r><t>b</t></r></si></sst>'
  const result = detectPreserved(packageOf({ [WORKSHEET_XML]: PLAIN_WORKSHEET, 'xl/sharedStrings.xml': shared }))
  expect(result.coupled).toHaveLength(2)
})

it('scans worksheets whose formula markup has no usable boundary', () => {
  const cases = [
    '<worksheet><sheetData><f',
    '<worksheet><sheetData><f ',
    '<worksheet><sheetData><fx/></sheetData></worksheet>',
    '<worksheet><sheetData><f t="shared"</sheetData></worksheet>',
  ] as const
  for (const worksheet of cases) {
    const result = detectPreserved(packageOf({ [WORKSHEET_XML]: worksheet }))
    expect(result.preserved).toEqual(new Set())
  }
})

it('leaves a package without worksheets or documents empty', () => {
  expect(detectPreserved(packageOf({ 'docProps/core.xml': '<props/>' }))).toEqual({
    preserved: new Set(['document-properties']),
    coupled: [],
  })
})

it('finds rich-text items across many shared strings with one forward cursor', () => {
  const items: string[] = []
  for (let index = 0; index < 2_000; index += 1) {
    items.push(index % 2 === 0 ? `<si><r><t>${String(index)}</t></r></si>` : `<si><t>${String(index)}</t></si>`)
  }
  const result = detectPreserved(packageOf({
    [WORKSHEET_XML]: PLAIN_WORKSHEET,
    'xl/sharedStrings.xml': `<sst>${items.join('')}</sst>`,
  }))
  const offsets = result.coupled
    .filter(marker => marker.kind === 'rich-text-shared-string')
    .map(marker => marker.offset)

  expect(offsets).toHaveLength(1_000)
  expect(new Set(offsets).size).toBe(offsets.length)
  expect([...offsets].sort((left, right) => left - right)).toEqual(offsets)
})

it('classifies both reports from one package scan', () => {
  const result = detectOfficePackage(packageOf({
    'word/vbaProject.bin': Uint8Array.of(1),
    [DOCUMENT_XML]: '<w:document><w:body><w:p><w:ins w:id="1"/></w:p></w:body></w:document>',
    [WORKSHEET_XML]: PLAIN_WORKSHEET,
    'xl/sharedStrings.xml': '<sst><si><r><t>a</t></r></si></sst>',
    'xl/workbook.xml': '<workbook><definedNames><definedName name="x">S!A1</definedName></definedNames></workbook>',
    'xl/drawings/drawing1.xml': '<xdr:wsDr><xdr:sp/></xdr:wsDr>',
  }))

  expect(result.unsupported.map(feature => feature.kind)).toEqual(['macro', 'tracked-changes', 'drawing'])
  expect([...result.preserved]).toEqual(['macro', 'defined-names', 'drawing'])
  expect(result.coupled.map(marker => marker.kind)).toEqual(['tracked-change', 'rich-text-shared-string'])
})
