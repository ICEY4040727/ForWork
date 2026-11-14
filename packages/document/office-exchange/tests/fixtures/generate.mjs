/**
 * Regenerates the Office Exchange round-trip fixtures.
 *
 * Run from anywhere:
 *   node packages/document/office-exchange/tests/fixtures/generate.mjs
 *
 * Output is deterministic: every zip entry uses the fixed `FIXED_DATE`
 * timestamp, workbook metadata is pinned to that date, and no random ids or
 * current-time values reach the archives. Do not hand-edit the generated
 * binaries; edit this script instead.
 *
 * `exceljs` and `xlsx` are installed for `packages/client/ui-sidebar-documentpreview`;
 * `fflate` is a direct dependency of `packages/document/office-exchange`. Each
 * module is resolved from its owning workspace package, falling back to the
 * repository root.
 */
import { createRequire } from 'node:module'
import { mkdirSync, unlinkSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const FIXTURES_DIR = dirname(fileURLToPath(import.meta.url))
const REPO_ROOT = join(FIXTURES_DIR, '..', '..', '..', '..', '..')

/** Single determinism source for archive timestamps and document properties. */
const FIXED_DATE = new Date('2026-01-01T00:00:00.000Z')
const ZIP_ENTRY_OPTIONS = { mtime: FIXED_DATE }

const requireFromRoot = createRequire(join(REPO_ROOT, 'package.json'))

/**
 * Resolve a module from its owning workspace package, then from the root.
 * @param {string} name - Bare module specifier to resolve.
 * @param {string[]} searchPaths - Candidate directories, most specific first.
 * @returns {string} Absolute path to the resolved module entry.
 */
function resolveModule(name, searchPaths) {
  for (const searchPath of [...searchPaths, REPO_ROOT]) {
    try {
      return requireFromRoot.resolve(name, { paths: [searchPath] })
    } catch {
      // Try the next candidate directory.
    }
  }
  throw new Error(`Cannot resolve ${name}; expected it in a workspace package or the repository root`)
}

const ExcelJS = requireFromRoot(
  resolveModule('exceljs', [join(REPO_ROOT, 'packages/client/ui-sidebar-documentpreview')]),
)
const XLSX = requireFromRoot(
  resolveModule('xlsx', [join(REPO_ROOT, 'packages/client/ui-sidebar-documentpreview')]),
)
const { strToU8, unzipSync, zipSync } = requireFromRoot(
  resolveModule('fflate', [join(REPO_ROOT, 'packages/document/office-exchange')]),
)

/**
 * Re-encode an archive with fixed timestamps and sorted entries.
 * @param {Uint8Array} archive - Archive bytes produced by a third-party writer.
 * @returns {Uint8Array} Deterministic archive bytes with identical entries.
 */
function normalizeArchive(archive) {
  const entries = unzipSync(archive)
  const sorted = {}
  for (const name of Object.keys(entries).sort()) {
    sorted[name] = [entries[name], ZIP_ENTRY_OPTIONS]
  }
  return zipSync(sorted, { level: 6 })
}

/**
 * Build a zip from named text and binary entries with fixed timestamps.
 * @param {Record<string, Uint8Array>} entries - Archive member contents keyed by path.
 * @returns {Uint8Array} Deterministic archive bytes.
 */
function buildArchive(entries) {
  const sorted = {}
  for (const name of Object.keys(entries).sort()) {
    sorted[name] = [entries[name], ZIP_ENTRY_OPTIONS]
  }
  return zipSync(sorted, { level: 6 })
}

async function buildRoundtripXlsx() {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Office Exchange fixtures'
  workbook.lastModifiedBy = 'Office Exchange fixtures'
  workbook.created = FIXED_DATE
  workbook.modified = FIXED_DATE

  const sales = workbook.addWorksheet('Sales')
  sales.addRow(['Item', 'Units', 'Revenue'])
  sales.addRow(['Widget', 4, 1200.5])
  sales.addRow(['Gadget', 7, 900])
  sales.addRow(['Cog', 2, 45.25])
  sales.getRow(1).eachCell(cell => {
    cell.font = { bold: true }
  })
  for (let row = 2; row <= 4; row += 1) {
    sales.getCell(`C${row}`).numFmt = '$#,##0.00'
  }
  sales.getCell('A5').value = 'Total units'
  sales.getCell('B5').value = { formula: 'SUM(B2:B4)' }
  sales.mergeCells('A7:C7')
  sales.getCell('A7').value = 'Merged summary row'

  const notes = workbook.addWorksheet('Notes')
  notes.addRow(['Note', 'Detail'])
  notes.addRow(['Reviewed', 'Fixture generated deterministically'])
  notes.addRow(['Owner', 'Office Exchange tests'])

  const raw = new Uint8Array(await workbook.xlsx.writeBuffer())
  return normalizeArchive(raw)
}

/**
 * Build `word/document.xml` with a heading, a bold run, and a 2x2 table.
 * @param {{trackedChanges?: boolean, header?: boolean}} options - Variant switches.
 * @returns {string} Document part XML.
 */
function documentXml({ trackedChanges = false, header = false } = {}) {
  const rootNamespaces = header
    ? ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    : ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  const headerReference = header
    ? '<w:headerReference w:type="default" r:id="rId2"/>'
    : ''
  const trackedParagraph = trackedChanges
    ? `
    <w:p>
      <w:ins w:id="1" w:author="Fixture Author" w:date="2026-01-01T00:00:00Z">
        <w:r><w:t>Inserted revision</w:t></w:r>
      </w:ins>
      <w:del w:id="2" w:author="Fixture Author" w:date="2026-01-01T00:00:00Z">
        <w:r><w:delText>Deleted revision</w:delText></w:r>
      </w:del>
    </w:p>`
    : ''

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document${rootNamespaces}>
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Quarterly Report</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t xml:space="preserve">This paragraph has </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>bold text</w:t></w:r>
      <w:r><w:t xml:space="preserve"> for round-trip checks.</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:color="auto"/>
          <w:left w:val="single" w:sz="4" w:color="auto"/>
          <w:bottom w:val="single" w:sz="4" w:color="auto"/>
          <w:right w:val="single" w:sz="4" w:color="auto"/>
          <w:insideH w:val="single" w:sz="4" w:color="auto"/>
          <w:insideV w:val="single" w:sz="4" w:color="auto"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>1200</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>${trackedParagraph}
    <w:sectPr>${headerReference}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>
`
}

/**
 * Build `word/styles.xml` with the normal and heading styles used by the fixture.
 * @returns {string} Styles part XML.
 */
function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
</w:styles>
`
}

/**
 * Build the package content-type manifest.
 * @param {{header?: boolean}} options - Variant switches.
 * @returns {string} Content types XML.
 */
function contentTypesXml({ header = false } = {}) {
  const headerOverride = header
    ? '\n  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    : ''
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${headerOverride}
</Types>
`
}

/**
 * Build the package-level relationships part.
 * @returns {string} Root relationships XML.
 */
function rootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
`
}

/**
 * Build the document relationships part.
 * @param {{header?: boolean}} options - Variant switches.
 * @returns {string} Document relationships XML.
 */
function documentRelsXml({ header = false } = {}) {
  const headerRelationship = header
    ? '\n  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
    : ''
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${headerRelationship}
</Relationships>
`
}

/**
 * Build a header part.
 * @returns {string} Header XML.
 */
function headerXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Confidential header</w:t></w:r></w:p>
</w:hdr>
`
}

/**
 * Build a minimal `.docx` package.
 * @param {{trackedChanges?: boolean, macro?: boolean, header?: boolean}} options - Variant switches.
 * @returns {Uint8Array} Deterministic archive bytes.
 */
function buildDocx({ trackedChanges = false, macro = false, header = false } = {}) {
  const entries = {
    '[Content_Types].xml': strToU8(contentTypesXml({ header })),
    '_rels/.rels': strToU8(rootRelsXml()),
    'word/_rels/document.xml.rels': strToU8(documentRelsXml({ header })),
    'word/document.xml': strToU8(documentXml({ trackedChanges, header })),
    'word/styles.xml': strToU8(stylesXml()),
  }
  if (macro) {
    entries['word/vbaProject.bin'] = new Uint8Array([
      0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    ])
  }
  if (header) entries['word/header1.xml'] = strToU8(headerXml())
  return buildArchive(entries)
}

function buildLegacyXls() {
  const workbook = XLSX.utils.book_new()
  workbook.Props = {
    Title: 'Office Exchange legacy fixture',
    CreatedDate: FIXED_DATE,
    ModifiedDate: FIXED_DATE,
  }
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Item', 'Units'],
    ['Widget', 4],
    ['Gadget', 7],
  ])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Legacy')
  return new Uint8Array(XLSX.write(workbook, { type: 'buffer', bookType: 'xls' }))
}

function writeFixture(name, bytes) {
  const target = join(FIXTURES_DIR, name)
  writeFileSync(target, bytes)
  console.log(`${name}: ${bytes.byteLength} bytes`)
}

async function main() {
  mkdirSync(FIXTURES_DIR, { recursive: true })

  writeFixture('roundtrip.xlsx', await buildRoundtripXlsx())
  writeFixture('roundtrip.docx', buildDocx())
  writeFixture('unsupported-macro.docx', buildDocx({ macro: true }))
  writeFixture('unsupported-tracked.docx', buildDocx({ trackedChanges: true }))
  writeFixture('unsupported-header.docx', buildDocx({ header: true }))

  try {
    writeFixture('legacy.xls', buildLegacyXls())
  } catch (error) {
    const stale = join(FIXTURES_DIR, 'legacy.xls')
    try {
      unlinkSync(stale)
    } catch {
      // No stale file to remove.
    }
    console.warn(`Skipping legacy.xls: ${error instanceof Error ? error.message : String(error)}`)
  }
}

await main()
