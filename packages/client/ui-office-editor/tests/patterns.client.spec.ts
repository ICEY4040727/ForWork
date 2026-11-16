/**
 * Which addresses the office type claims, and which stay preview-only.
 *
 * `.docx` and `.xlsx` are editable, and the legacy `.doc`, `.xls`, and `.ppt`
 * sources are claimed so opening one converts it; `.pptx` never reaches this
 * type and keeps the read-only document preview the Sidebar ships as the
 * fallback viewer.
 */
import { describe, expect, it } from 'vitest'
import { Context } from '@forwork/cordis'
import { sessionFileAddress } from '@forwork/util-workspace-path'
import { SidebarRightTabRegistry } from '@forwork/client-ui-sidebar-right/src/client/tab-registry.ts'
import { TEXTPREVIEW_KIND, textDefinition } from '@forwork/client-ui-sidebar-documentpreview/src/client/definition.ts'
import { OFFICE_KIND, basenameOf, officeDefinition } from '../src/client/definition.ts'

function registry() {
  const tabs = new SidebarRightTabRegistry(new Context())
  tabs.register(textDefinition())
  tabs.register(officeDefinition())
  return tabs
}

describe('basenameOf', () => {
  it('decodes the last segment, keeps an empty tail as the address, and leaves a malformed escape raw', () => {
    expect(basenameOf('dsh-resource://file/session/s-1/work/a%20b%23c.xlsx')).toBe('a b#c.xlsx')
    expect(basenameOf('dsh-resource://file/session/s-1/')).toBe('dsh-resource://file/session/s-1/')
    expect(basenameOf('dsh-resource://file/session/s-1/work/%E0%A4%A.docx')).toBe('%E0%A4%A.docx')
  })
})

describe('office tab patterns', () => {
  it('declares the editable suffixes and the legacy sources it converts', () => {
    const definition = officeDefinition()
    expect(definition.patterns).toEqual(['*.docx', '*.xlsx', '*.doc', '*.xls', '*.ppt'])
    expect(definition.priority).toBe('builtin')
    expect(definition.id).toBe('@forwork/client-ui-office-editor')
    expect(definition.kind).toBe(OFFICE_KIND)
  })

  it.each([
    sessionFileAddress('s-1', 'budget.xlsx'),
    sessionFileAddress('s-1', 'work/report.docx'),
    sessionFileAddress('s-1', '/home/me/REPORT.DOCX'),
    sessionFileAddress('s-1', 'legacy.doc'),
    sessionFileAddress('s-1', 'legacy.xls'),
    sessionFileAddress('s-1', 'deck.ppt'),
  ])('claims %s and titles it by basename', (address) => {
    expect(registry().claim(address).kind).toBe(OFFICE_KIND)
  })

  it.each([
    sessionFileAddress('s-1', 'deck.pptx'),
    sessionFileAddress('s-1', 'notes.txt'),
  ])('leaves %s to the read-only preview', (address) => {
    const tabs = registry()
    expect(tabs.candidates(address).map(type => type.kind)).toEqual([TEXTPREVIEW_KIND])
    expect(tabs.claim(address).kind).toBe(TEXTPREVIEW_KIND)
  })

  it('does not claim an address whose path ends in an editable suffix outside a Session', () => {
    const tabs = registry()
    expect(tabs.candidates('https://example.test/report.docx').map(type => type.kind)).toEqual([])
    expect(tabs.candidates('dsh-resource://file/absolute/home/me/report.xlsx')).toEqual([])
  })
})
