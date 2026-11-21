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
import { mkdirSync, readFileSync, unlinkSync, writeFileSync } from 'node:fs'
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
const { deflateSync, strFromU8, strToU8, unzipSync, zipSync } = requireFromRoot(
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
 * Build a drawing run that references an image relationship.
 * @param {string} attribute - `r:embed="rId5"` or `r:link="rId6"`.
 * @returns {string} Run XML.
 */
function drawingRun(attribute) {
  return `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">`
    + '<wp:extent cx="952500" cy="952500"/>'
    + '<wp:docPr id="1" name="Picture 1"/>'
    + '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    + '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + '<pic:nvPicPr><pic:cNvPr id="0" name="image1.jpeg"/><pic:cNvPicPr/></pic:nvPicPr>'
    + '<pic:blipFill><a:blip'
    + ` ${attribute}/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>'
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
    + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
}

/**
 * Build the document part for an embedded or linked image fixture.
 * @param {string} attribute - Image relationship attribute.
 * @returns {string} Document part XML.
 */
function imageDocumentXml(attribute) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Before image </w:t></w:r>
      ${drawingRun(attribute)}
      <w:r><w:t xml:space="preserve">after image.</w:t></w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>
`
}

/**
 * Build the document part for the footnote/endnote fixture.
 * @returns {string} Document part XML.
 */
function notesDocumentXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Body text</w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>
      <w:r><w:t xml:space="preserve"> with an endnote</w:t></w:r>
      <w:r><w:endnoteReference w:id="2"/></w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>
`
}

/** A deterministic JPEG-signature payload; fixtures assert byte preservation, not decoding. */
const JPEG_BYTES = Uint8Array.from([
  0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46,
  0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00,
  0xff, 0xdb, 0x00, 0x43, 0x00, 0x08, 0x06, 0x06, 0x07, 0x06,
  0x05, 0x08, 0x07, 0x07, 0x07, 0x09, 0x09, 0x08, 0x0a, 0x0c,
  0xff, 0xd9,
])

/**
 * Build the content types for the image fixtures.
 * @param {{jpeg?: boolean, notes?: boolean}} options - Declarations to include.
 * @returns {string} Content types XML.
 */
function preservationContentTypesXml({ jpeg = false, notes = false } = {}) {
  const jpegDefault = jpeg
    ? '\n  <Default Extension="jpeg" ContentType="image/jpeg"/>'
    : ''
  const noteOverrides = notes
    ? '\n  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
      + '\n  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
    : ''
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>${jpegDefault}
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${noteOverrides}
</Types>
`
}

/**
 * Build a package that carries an embedded or linked image.
 * @param {{external?: boolean}} options - Relationship target mode.
 * @returns {Uint8Array} Deterministic archive bytes.
 */
function buildImageDocx({ external = false } = {}) {
  const attribute = external ? 'r:link="rId6"' : 'r:embed="rId5"'
  const relationship = external
    ? '\n  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/linked.jpeg" TargetMode="External"/>'
    : '\n  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.jpeg"/>'
  const entries = {
    '[Content_Types].xml': strToU8(preservationContentTypesXml({ jpeg: !external })),
    '_rels/.rels': strToU8(rootRelsXml()),
    'word/_rels/document.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${relationship}
</Relationships>
`),
    'word/document.xml': strToU8(imageDocumentXml(attribute)),
    'word/styles.xml': strToU8(stylesXml()),
  }
  if (!external) entries['word/media/image1.jpeg'] = JPEG_BYTES
  return buildArchive(entries)
}

/**
 * Build a package with a footnote and an endnote.
 * @returns {Uint8Array} Deterministic archive bytes.
 */
function buildNotesDocx() {
  return buildArchive({
    '[Content_Types].xml': strToU8(preservationContentTypesXml({ notes: true })),
    '_rels/.rels': strToU8(rootRelsXml()),
    'word/_rels/document.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>
`),
    'word/document.xml': strToU8(notesDocumentXml()),
    'word/styles.xml': strToU8(stylesXml()),
    'word/footnotes.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:t>Footnote text</w:t></w:r></w:p></w:footnote>
</w:footnotes>
`),
    'word/endnotes.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:id="2"><w:p><w:r><w:t>Endnote text</w:t></w:r></w:p></w:endnote>
</w:endnotes>
`),
  })
}

/**
 * Build a minimal `.docx` package.
 * @param {{trackedChanges?: boolean, macro?: boolean, header?: boolean}} options - Variant switches.
 * @returns {Uint8Array} Deterministic archive bytes.
 */
function buildDocx({ trackedChanges = false, macro = false, header = false } = {}) {
  const macroContentType = '<Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>'
  const macroRelationship = '<Relationship Id="rId4"'
    + ' Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"'
    + ' Target="vbaProject.bin"/>'
  const contentTypes = macro
    ? contentTypesXml({ header }).replace('</Types>', `  ${macroContentType}\n</Types>`)
    : contentTypesXml({ header })
  const documentRelationships = macro
    ? documentRelsXml({ header }).replace('</Relationships>', `  ${macroRelationship}\n</Relationships>`)
    : documentRelsXml({ header })
  const entries = {
    '[Content_Types].xml': strToU8(contentTypes),
    '_rels/.rels': strToU8(rootRelsXml()),
    'word/_rels/document.xml.rels': strToU8(documentRelationships),
    'word/document.xml': strToU8(documentXml({ trackedChanges, header })),
    'word/styles.xml': strToU8(stylesXml()),
  }
  if (macro) entries['word/vbaProject.bin'] = VBA_PROJECT_BYTES
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

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const DEFAULT_SECTION = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
const DOCUMENT_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
const MACRO_DOCUMENT_CONTENT_TYPE = 'application/vnd.ms-word.document.macroEnabled.main+xml'
const XLSX_WORKBOOK_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
const MACRO_XLSX_CONTENT_TYPE = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
const XLSX_WORKSHEET_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
const XLSX_STYLES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
const VBA_CONTENT_TYPE = 'application/vnd.ms-office.vbaProject'
const VBA_RELATIONSHIP = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'

const PNG_BYTES = Uint8Array.from(Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
))
const EMF_BYTES = Uint8Array.from([
  0x01, 0x00, 0x00, 0x00, 0x6c, 0x00, 0x00, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  0x20, 0x45, 0x4d, 0x46, 0x00, 0x00, 0x01, 0x00,
  0x6c, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  0x14, 0x00, 0x00, 0x00, 0x0e, 0x00, 0x00, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  0x14, 0x00, 0x00, 0x00,
])
const VBA_PROJECT_BYTES = Uint8Array.from([
  0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
])

/**
 * Build a content-type manifest with the package defaults plus caller declarations.
 * @param {{defaults?: readonly (readonly [string, string])[], overrides?: readonly (readonly [string, string])[]}} options - Extra declarations.
 * @returns {string} Content-types XML.
 */
function contentTypesXmlWith({ defaults = [], overrides = [] } = {}) {
  const lines = [
    ['rels', 'application/vnd.openxmlformats-package.relationships+xml'],
    ['xml', 'application/xml'],
    ...defaults,
  ].map(([extension, contentType]) => `  <Default Extension="${extension}" ContentType="${contentType}"/>`)
  for (const [partName, contentType] of overrides) {
    lines.push(`  <Override PartName="${partName}" ContentType="${contentType}"/>`)
  }
  return `${XML_HEADER}\n<Types xmlns="${CONTENT_TYPES_NS}">\n${lines.join('\n')}\n</Types>\n`
}

/**
 * Build one relationship element, expanding the standard relationship base.
 * @param {string} id - Relationship id.
 * @param {string} type - Full type URI or a suffix under the officeDocument relationship base.
 * @param {string} target - Target part or external URI.
 * @param {string} [mode] - `External` for external targets.
 * @returns {string} Relationship XML element.
 */
function relationship(id, type, target, mode = '') {
  const typeUri = type.includes('://') ? type : `${R_NS}/${type}`
  const targetMode = mode === '' ? '' : ` TargetMode="${mode}"`
  return `  <Relationship Id="${id}" Type="${typeUri}" Target="${target}"${targetMode}/>`
}

/**
 * Build a relationships part from relationship elements.
 * @param {readonly string[]} relationships - Relationship XML elements.
 * @returns {string} Relationships XML.
 */
function relationshipsPart(relationships) {
  return `${XML_HEADER}\n<Relationships xmlns="${PACKAGE_REL_NS}">\n${relationships.join('\n')}\n</Relationships>\n`
}

/**
 * Build a `w:document` part.
 * @param {string} body - Body children.
 * @param {{namespaces?: string, section?: string}} [options] - Extra namespaces and section properties.
 * @returns {string} Document XML.
 */
function documentPart(body, { namespaces = '', section = DEFAULT_SECTION } = {}) {
  return `${XML_HEADER}\n<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"${namespaces}>\n<w:body>\n${body}\n${section}\n</w:body>\n</w:document>\n`
}

/**
 * Build a drawing run with a stable drawing id.
 * @param {string} attribute - `r:embed="rId5"` or `r:link="rId6"`.
 * @param {number} id - `wp:docPr` id.
 * @param {string} name - `wp:docPr` name.
 * @returns {string} Run XML.
 */
function pictureRun(attribute, id, name) {
  return drawingRun(attribute)
    .replace('id="1"', `id="${id}"`)
    .replace('name="Picture 1"', `name="${name}"`)
}

/**
 * Build a DOCX package with the shared styles part and document relationships.
 * @param {{
 *   body?: string,
 *   documentXml?: string,
 *   namespaces?: string,
 *   section?: string,
 *   documentRelationships: readonly string[],
 *   contentDefaults?: readonly (readonly [string, string])[],
 *   contentOverrides?: readonly (readonly [string, string])[],
 *   mainContentType?: string,
 *   rootRelationships?: readonly string[],
 *   extraEntries?: Record<string, Uint8Array>,
 * }} options - Package contents.
 * @returns {Uint8Array} Deterministic archive bytes.
 */
function docxPackage({
  body = '',
  documentXml,
  namespaces = '',
  section = DEFAULT_SECTION,
  documentRelationships,
  contentDefaults = [],
  contentOverrides = [],
  mainContentType = DOCUMENT_CONTENT_TYPE,
  rootRelationships = [relationship('rId1', 'officeDocument', 'word/document.xml')],
  extraEntries = {},
}) {
  const entries = {
    '[Content_Types].xml': strToU8(contentTypesXmlWith({
      defaults: contentDefaults,
      overrides: [['/word/document.xml', mainContentType], ...contentOverrides],
    })),
    '_rels/.rels': strToU8(relationshipsPart(rootRelationships)),
    'word/_rels/document.xml.rels': strToU8(relationshipsPart(documentRelationships)),
    'word/document.xml': strToU8(documentXml ?? documentPart(body, { namespaces, section })),
    'word/styles.xml': strToU8(stylesXml()),
    ...extraEntries,
  }
  return buildArchive(entries)
}

function buildHeaderImageDocx() {
  return docxPackage({
    body: '<w:p><w:r><w:t>Header image fixture.</w:t></w:r></w:p>',
    section: '<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId2', 'header', 'header1.xml'),
    ],
    contentDefaults: [['png', 'image/png']],
    contentOverrides: [['/word/header1.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml']],
    extraEntries: {
      'word/header1.xml': strToU8(`${XML_HEADER}\n<w:hdr xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:p>${pictureRun('r:embed="rId1"', 1, 'Header image')}</w:p></w:hdr>\n`),
      'word/_rels/header1.xml.rels': strToU8(relationshipsPart([
        relationship('rId1', 'image', 'media/header1.png'),
      ])),
      'word/media/header1.png': PNG_BYTES,
    },
  })
}

function buildFooterImageDocx() {
  return docxPackage({
    body: '<w:p><w:r><w:t>Footer image fixture.</w:t></w:r></w:p>',
    section: '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId2', 'footer', 'footer1.xml'),
    ],
    contentDefaults: [['png', 'image/png']],
    contentOverrides: [['/word/footer1.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml']],
    extraEntries: {
      'word/footer1.xml': strToU8(`${XML_HEADER}\n<w:ftr xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:p>${pictureRun('r:embed="rId1"', 1, 'Footer image')}</w:p></w:ftr>\n`),
      'word/_rels/footer1.xml.rels': strToU8(relationshipsPart([
        relationship('rId1', 'image', 'media/footer1.png'),
      ])),
      'word/media/footer1.png': PNG_BYTES,
    },
  })
}

function buildNoteImageDocx() {
  return docxPackage({
    body: '<w:p><w:r><w:t>Body text</w:t></w:r>'
      + '<w:r><w:footnoteReference w:id="1"/></w:r>'
      + '<w:r><w:endnoteReference w:id="2"/></w:r></w:p>',
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId2', 'footnotes', 'footnotes.xml'),
      relationship('rId3', 'endnotes', 'endnotes.xml'),
    ],
    contentDefaults: [['png', 'image/png']],
    contentOverrides: [
      ['/word/footnotes.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'],
      ['/word/endnotes.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml'],
    ],
    extraEntries: {
      'word/footnotes.xml': strToU8(`${XML_HEADER}\n<w:footnotes xmlns:w="${W_NS}" xmlns:r="${R_NS}">`
        + '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
        + `<w:footnote w:id="1"><w:p>${pictureRun('r:embed="rId1"', 1, 'Footnote image')}</w:p></w:footnote>`
        + '</w:footnotes>\n'),
      'word/_rels/footnotes.xml.rels': strToU8(relationshipsPart([
        relationship('rId1', 'image', 'media/note1.png'),
      ])),
      'word/endnotes.xml': strToU8(`${XML_HEADER}\n<w:endnotes xmlns:w="${W_NS}">`
        + '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>'
        + '<w:endnote w:id="2"><w:p><w:r><w:t>Endnote text</w:t></w:r></w:p></w:endnote>'
        + '</w:endnotes>\n'),
      'word/media/note1.png': PNG_BYTES,
    },
  })
}

function buildMultiMediaDocx() {
  return docxPackage({
    body: '<w:p><w:r><w:t xml:space="preserve">Before media </w:t></w:r>'
      + pictureRun('r:embed="rId5"', 1, 'PNG image')
      + pictureRun('r:embed="rId6"', 2, 'JPEG image')
      + pictureRun('r:embed="rId7"', 3, 'EMF image')
      + '<w:r><w:t xml:space="preserve"> after media.</w:t></w:r></w:p>',
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId5', 'image', 'media/image1.png'),
      relationship('rId6', 'image', 'media/image2.jpeg'),
      relationship('rId7', 'image', 'media/image3.emf'),
    ],
    contentDefaults: [['png', 'image/png'], ['jpeg', 'image/jpeg'], ['emf', 'image/x-emf']],
    extraEntries: {
      'word/media/image1.png': PNG_BYTES,
      'word/media/image2.jpeg': JPEG_BYTES,
      'word/media/image3.emf': EMF_BYTES,
    },
  })
}

function buildTableCellImageDocx() {
  return docxPackage({
    body: '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>'
      + '<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
      + `<w:p>${pictureRun('r:embed="rId5"', 1, 'Cell image')}</w:p></w:tc></w:tr></w:tbl>`,
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId5', 'image', 'media/cell1.png'),
    ],
    contentDefaults: [['png', 'image/png']],
    extraEntries: {
      'word/media/cell1.png': PNG_BYTES,
    },
  })
}

function buildOffsetShiftDocx() {
  const body = '<w:p><w:r><w:t>First paragraph before the preserved block.</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t xml:space="preserve">Second paragraph with enough text so the later image starts at a nonzero byte offset.</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t xml:space="preserve">Before image </w:t></w:r>'
    + pictureRun('r:embed="rId5"', 1, 'Offset image')
    + '<w:r><w:t xml:space="preserve"> after image.</w:t></w:r></w:p>'
  return docxPackage({
    documentXml: documentPart(body).replaceAll('\n', '\r\n'),
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId5', 'image', 'media/offset1.png'),
    ],
    contentDefaults: [['png', 'image/png']],
    extraEntries: {
      'word/media/offset1.png': PNG_BYTES,
    },
  })
}

function buildPreservedIdCollisionDocx() {
  return docxPackage({
    body: '<w:p><w:r><w:t xml:space="preserve">Collision </w:t></w:r>'
      + pictureRun('r:embed="rId2"', 1, 'Collision image')
      + '<w:hyperlink r:id="rId9"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>',
    section: '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId2', 'image', 'media/collision1.png'),
      relationship('rId3', 'header', 'header1.xml'),
      relationship('rId9', 'hyperlink', 'https://example.test/collision', 'External'),
    ],
    contentDefaults: [['png', 'image/png']],
    contentOverrides: [['/word/header1.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml']],
    extraEntries: {
      'word/header1.xml': strToU8(`${XML_HEADER}\n<w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>Collision header</w:t></w:r></w:p></w:hdr>\n`),
      'word/media/collision1.png': PNG_BYTES,
    },
  })
}

function buildCommentDocx() {
  return docxPackage({
    body: '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented text</w:t></w:r>'
      + '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>',
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId5', 'comments', 'comments.xml'),
      relationship('rId6', 'commentsExtended', 'commentsExtended.xml'),
      relationship('rId7', 'commentsIds', 'commentsIds.xml'),
      relationship('rId8', 'people', 'people.xml'),
    ],
    contentOverrides: [
      ['/word/comments.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'],
      ['/word/commentsExtended.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml'],
      ['/word/commentsIds.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml'],
      ['/word/people.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml'],
    ],
    extraEntries: {
      'word/comments.xml': strToU8(`${XML_HEADER}\n<w:comments xmlns:w="${W_NS}"`
        + ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        + '<w:comment w:id="0" w:author="Fixture Author" w:date="2026-01-01T00:00:00Z" w:initials="FA" w14:paraId="00000001">'
        + '<w:p><w:r><w:t>Comment text</w:t></w:r></w:p></w:comment>'
        + '</w:comments>\n'),
      'word/commentsExtended.xml': strToU8(`${XML_HEADER}\n<w15:commentsEx`
        + ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
        + '<w15:commentEx w15:paraId="00000001" w15:done="0"/>'
        + '</w15:commentsEx>\n'),
      'word/commentsIds.xml': strToU8(`${XML_HEADER}\n<w16cid:commentsIds`
        + ' xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">'
        + '<w16cid:commentId w16cid:paraId="00000001" w16cid:durableId="1001"/>'
        + '</w16cid:commentsIds>\n'),
      'word/people.xml': strToU8(`${XML_HEADER}\n<w15:people`
        + ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
        + '<w15:person w15:author="Fixture Author">'
        + '<w15:presenceInfo w15:providerId="None" w15:userId="Fixture Author"/>'
        + '</w15:person></w15:people>\n'),
    },
  })
}

function buildChartDocx() {
  const chartDrawing = '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
    + '<wp:extent cx="3429000" cy="2286000"/><wp:docPr id="7" name="Chart 1"/>'
    + '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    + '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    + `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId5"/>`
    + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  return docxPackage({
    body: `<w:p><w:r><w:t xml:space="preserve">Chart </w:t></w:r>${chartDrawing}</w:p>`,
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId5', 'chart', 'charts/chart1.xml'),
    ],
    contentOverrides: [['/word/charts/chart1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml']],
    extraEntries: {
      'word/charts/chart1.xml': strToU8(`${XML_HEADER}\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${R_NS}">`
        + '<c:chart><c:plotArea><c:layout/></c:plotArea></c:chart></c:chartSpace>\n'),
    },
  })
}

function buildEmbeddedObjectDocx() {
  return docxPackage({
    namespaces: ' xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"',
    body: '<w:p><w:r><w:object>'
      + '<v:shape id="Object1" style="width:72pt;height:32pt"><v:imagedata r:id="rId6" o:title=""/></v:shape>'
      + '<o:OLEObject Type="Embed" ProgID="Package" ShapeID="Object1" DrawAspect="Content" ObjectID="_1" r:id="rId5"/>'
      + '</w:object></w:r><w:r><w:t xml:space="preserve"> embedded object</w:t></w:r></w:p>',
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId5', 'oleObject', 'embeddings/oleObject1.bin'),
      relationship('rId6', 'image', 'media/ole-preview.png'),
    ],
    contentDefaults: [['png', 'image/png']],
    contentOverrides: [['/word/embeddings/oleObject1.bin', 'application/vnd.openxmlformats-officedocument.oleObject']],
    extraEntries: {
      'word/embeddings/oleObject1.bin': Uint8Array.from([
        0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
      ]),
      'word/media/ole-preview.png': PNG_BYTES,
    },
  })
}

function buildShapeDocx() {
  const shape = '<w:pict><v:rect id="Rectangle1" style="width:120pt;height:40pt" fillcolor="#4472C4">'
    + '<v:textbox><w:txbxContent><w:p><w:r><w:t>Shape text</w:t></w:r></w:p></w:txbxContent></v:textbox>'
    + '</v:rect></w:pict>'
  return docxPackage({
    namespaces: ' xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"',
    body: `<w:p>${shape}</w:p>`,
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId5', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing', 'drawings/vmlDrawing1.vml'),
    ],
    contentDefaults: [['vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing']],
    extraEntries: {
      'word/drawings/vmlDrawing1.vml': strToU8(`${XML_HEADER}\n<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="${W_NS}">`
        + '<v:shape id="Rectangle1" style="width:120pt;height:40pt" fillcolor="#4472C4">'
        + '<v:textbox><w:txbxContent><w:p><w:r><w:t>Shape text</w:t></w:r></w:p></w:txbxContent></v:textbox>'
        + '</v:shape></xml>\n'),
    },
  })
}

function buildExternalLinkDocx() {
  return docxPackage({
    body: '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>'
      + '<w:hyperlink r:id="rId9"><w:r><w:t>the external reference</w:t></w:r></w:hyperlink></w:p>',
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId9', 'hyperlink', 'https://example.test/external-target', 'External'),
    ],
  })
}

function buildFieldDocx() {
  return docxPackage({
    body: '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>'
      + '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>'
      + '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
      + '<w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r>'
      + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
      + '<w:r><w:t>2026-01-01</w:t></w:r>'
      + '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
    documentRelationships: [relationship('rId1', 'styles', 'styles.xml')],
  })
}

function buildBookmarkDocx() {
  return docxPackage({
    body: '<w:p><w:bookmarkStart w:id="0" w:name="target_bookmark"/>'
      + '<w:r><w:t>Bookmarked text</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>'
      + '<w:p><w:hyperlink w:anchor="target_bookmark"><w:r><w:t>Jump to bookmark</w:t></w:r></w:hyperlink></w:p>',
    documentRelationships: [relationship('rId1', 'styles', 'styles.xml')],
  })
}

function buildContentControlDocx() {
  return docxPackage({
    body: '<w:sdt><w:sdtPr><w:alias w:val="Summary"/><w:tag w:val="summary"/><w:id w:val="123456"/></w:sdtPr>'
      + '<w:sdtContent><w:p><w:r><w:t>Content control text</w:t></w:r></w:p></w:sdtContent></w:sdt>',
    documentRelationships: [relationship('rId1', 'styles', 'styles.xml')],
  })
}

function buildUnmodeledPartsDocx() {
  const theme = `${XML_HEADER}\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Fixture Theme">`
    + '<a:themeElements>'
    + '<a:clrScheme name="Fixture">'
    + '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    + '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
    + '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
    + '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
    + '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
    + '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
    + '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
    + '</a:clrScheme>'
    + '<a:fontScheme name="Fixture">'
    + '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    + '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
    + '</a:fontScheme>'
    + '<a:fmtScheme name="Fixture">'
    + '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
    + '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
    + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    + '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
    + '</a:fmtScheme>'
    + '</a:themeElements></a:theme>\n'
  return docxPackage({
    body: '<w:p><w:r><w:t>Unmodeled parts fixture.</w:t></w:r></w:p>',
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId2', 'theme', 'theme/theme1.xml'),
      relationship('rId3', 'fontTable', 'fontTable.xml'),
      relationship('rId4', 'webSettings', 'webSettings.xml'),
      relationship('rId5', 'customXml', '../customXml/item1.xml'),
    ],
    rootRelationships: [
      relationship('rId1', 'officeDocument', 'word/document.xml'),
      relationship('rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'),
      relationship('rId3', 'extended-properties', 'docProps/app.xml'),
    ],
    contentOverrides: [
      ['/word/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
      ['/word/fontTable.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'],
      ['/word/webSettings.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml'],
      ['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
      ['/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'],
      ['/customXml/itemProps1.xml', 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'],
    ],
    extraEntries: {
      'word/theme/theme1.xml': strToU8(theme),
      'word/fontTable.xml': strToU8(`${XML_HEADER}\n<w:fonts xmlns:w="${W_NS}">`
        + '<w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/>'
        + '<w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>'
        + '</w:fonts>\n'),
      'word/webSettings.xml': strToU8(`${XML_HEADER}\n<w:webSettings xmlns:w="${W_NS}"><w:optimizeForBrowser/></w:webSettings>\n`),
      'docProps/core.xml': strToU8(`${XML_HEADER}\n<cp:coreProperties`
        + ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
        + ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
        + ' xmlns:dcterms="http://purl.org/dc/terms/"'
        + ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        + '<dc:title>Unmodeled parts fixture</dc:title>'
        + '<dc:creator>Office Exchange fixtures</dc:creator>'
        + '<cp:lastModifiedBy>Office Exchange fixtures</cp:lastModifiedBy>'
        + '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>'
        + '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>'
        + '</cp:coreProperties>\n'),
      'docProps/app.xml': strToU8(`${XML_HEADER}\n<Properties`
        + ' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        + '<Application>Office Exchange fixtures</Application><DocSecurity>0</DocSecurity>'
        + '<ScaleCrop>false</ScaleCrop></Properties>\n'),
      'customXml/item1.xml': strToU8(`${XML_HEADER}\n<fixture xmlns="urn:example:custom-data"><value>custom xml payload</value></fixture>\n`),
      'customXml/itemProps1.xml': strToU8(`${XML_HEADER}\n<ds:datastoreItem`
        + ' ds:itemID="{11111111-1111-1111-1111-111111111111}"'
        + ' xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
        + '<ds:schemaRefs><ds:schemaRef ds:uri="urn:example:custom-data"/></ds:schemaRefs>'
        + '</ds:datastoreItem>\n'),
      'customXml/_rels/item1.xml.rels': strToU8(relationshipsPart([
        relationship('rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps', 'itemProps1.xml'),
      ])),
    },
  })
}

function buildMacroDocm() {
  return docxPackage({
    documentXml: documentXml(),
    documentRelationships: [
      relationship('rId1', 'styles', 'styles.xml'),
      relationship('rId4', VBA_RELATIONSHIP, 'vbaProject.bin'),
    ],
    mainContentType: MACRO_DOCUMENT_CONTENT_TYPE,
    contentOverrides: [['/word/vbaProject.bin', VBA_CONTENT_TYPE]],
    extraEntries: {
      'word/vbaProject.bin': VBA_PROJECT_BYTES,
    },
  })
}

const XLSX_STYLES_XML = `${XML_HEADER}\n<styleSheet xmlns="${SPREADSHEET_NS}">`
  + '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
  + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
  + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
  + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
  + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
  + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
  + '</styleSheet>\n'

const XLSX_DATA_ROWS = '<row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c>'
  + '<c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row>'
  + '<row r="2"><c r="A2" t="inlineStr"><is><t>Alpha</t></is></c><c r="B2"><v>10</v></c></row>'
  + '<row r="3"><c r="A3" t="inlineStr"><is><t>Beta</t></is></c><c r="B3"><v>20</v></c></row>'
  + '<row r="4"><c r="A4" t="inlineStr"><is><t>Alpha</t></is></c><c r="B4"><v>30</v></c></row>'

/**
 * Build a minimal workbook part.
 * @param {{definedNames?: string, extra?: string}} [options] - Defined names and trailing elements.
 * @returns {string} Workbook XML.
 */
function workbookPart({ definedNames = '', extra = '' } = {}) {
  return `${XML_HEADER}\n<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${R_NS}">`
    + '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>'
    + `${definedNames}${extra}</workbook>\n`
}

/**
 * Build one worksheet part.
 * @param {{sheetData?: string, extra?: string}} [options] - Sheet cells and trailing elements.
 * @returns {string} Worksheet XML.
 */
function worksheetPart({ sheetData = XLSX_DATA_ROWS, extra = '' } = {}) {
  return `${XML_HEADER}\n<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${R_NS}">`
    + '<dimension ref="A1:B4"/><sheetData>'
    + `${sheetData}</sheetData>${extra}</worksheet>\n`
}

/**
 * Build an XLSX package with the shared workbook, sheet, and styles parts.
 * @param {{
 *   workbook: string,
 *   workbookRelationships: readonly string[],
 *   sheet: string,
 *   sheetRelationships?: readonly string[],
 *   defaults?: readonly (readonly [string, string])[],
 *   overrides?: readonly (readonly [string, string])[],
 *   extraEntries?: Record<string, Uint8Array>,
 * }} options - Package contents.
 * @returns {Uint8Array} Deterministic archive bytes.
 */
function xlsxPackage({
  workbook,
  workbookRelationships,
  sheet,
  sheetRelationships = [],
  defaults = [],
  overrides = [],
  extraEntries = {},
}) {
  const entries = {
    '[Content_Types].xml': strToU8(contentTypesXmlWith({
      defaults,
      overrides: [
        ['/xl/workbook.xml', XLSX_WORKBOOK_CONTENT_TYPE],
        ['/xl/worksheets/sheet1.xml', XLSX_WORKSHEET_CONTENT_TYPE],
        ['/xl/styles.xml', XLSX_STYLES_CONTENT_TYPE],
        ...overrides,
      ],
    })),
    '_rels/.rels': strToU8(relationshipsPart([
      relationship('rId1', 'officeDocument', 'xl/workbook.xml'),
    ])),
    'xl/workbook.xml': strToU8(workbook),
    'xl/_rels/workbook.xml.rels': strToU8(relationshipsPart(workbookRelationships)),
    'xl/worksheets/sheet1.xml': strToU8(sheet),
    'xl/styles.xml': strToU8(XLSX_STYLES_XML),
    ...extraEntries,
  }
  if (sheetRelationships.length > 0) {
    entries['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(relationshipsPart(sheetRelationships))
  }
  return buildArchive(entries)
}

function buildPivotTableXlsx() {
  return xlsxPackage({
    workbook: workbookPart({
      extra: '<pivotCaches><pivotCache cacheId="1" r:id="rId2"/></pivotCaches>',
    }),
    workbookRelationships: [
      relationship('rId1', 'worksheet', 'worksheets/sheet1.xml'),
      relationship('rId2', 'pivotCacheDefinition', 'pivotCache/pivotCacheDefinition1.xml'),
      relationship('rId3', 'styles', 'styles.xml'),
    ],
    sheet: worksheetPart({
      extra: '<pivotTableParts count="1"><pivotTablePart r:id="rId1"/></pivotTableParts>',
    }),
    sheetRelationships: [relationship('rId1', 'pivotTable', '../pivotTables/pivotTable1.xml')],
    overrides: [
      ['/xl/pivotTables/pivotTable1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml'],
      ['/xl/pivotCache/pivotCacheDefinition1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'],
      ['/xl/pivotCache/pivotCacheRecords1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml'],
    ],
    extraEntries: {
      'xl/pivotTables/pivotTable1.xml': strToU8(`${XML_HEADER}\n<pivotTableDefinition xmlns="${SPREADSHEET_NS}" name="PivotTable1" cacheId="1" dataCaption="Values">`
        + '<location ref="D3:E5" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>'
        + '<pivotFields count="2"><pivotField showAll="0"/><pivotField showAll="0"/></pivotFields>'
        + '<rowFields count="1"><field x="0"/></rowFields><colFields count="1"><field x="1"/></colFields>'
        + '<dataFields count="1"><dataField name="Sum of Amount" fld="1" baseField="0" baseItem="0"/></dataFields>'
        + '</pivotTableDefinition>\n'),
      'xl/pivotCache/pivotCacheDefinition1.xml': strToU8(`${XML_HEADER}\n<pivotCacheDefinition xmlns="${SPREADSHEET_NS}" xmlns:r="${R_NS}" r:id="rId1" refreshOnLoad="1" recordCount="3">`
        + '<cacheSource type="worksheet"><worksheetSource ref="A1:B4" sheet="Data"/></cacheSource>'
        + '<cacheFields count="2">'
        + '<cacheField name="Category" numFmtId="0"><sharedItems count="2"><s v="Alpha"/><s v="Beta"/></sharedItems></cacheField>'
        + '<cacheField name="Amount" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="10" maxValue="30"/></cacheField>'
        + '</cacheFields></pivotCacheDefinition>\n'),
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': strToU8(relationshipsPart([
        relationship('rId1', 'pivotCacheRecords', 'pivotCacheRecords1.xml'),
      ])),
      'xl/pivotCache/pivotCacheRecords1.xml': strToU8(`${XML_HEADER}\n<pivotCacheRecords xmlns="${SPREADSHEET_NS}" xmlns:r="${R_NS}" count="3">`
        + '<r><s v="Alpha"/><n v="10"/></r><r><s v="Beta"/><n v="20"/></r><r><s v="Alpha"/><n v="30"/></r>'
        + '</pivotCacheRecords>\n'),
    },
  })
}

function buildExcelTableXlsx() {
  return xlsxPackage({
    workbook: workbookPart(),
    workbookRelationships: [
      relationship('rId1', 'worksheet', 'worksheets/sheet1.xml'),
      relationship('rId2', 'styles', 'styles.xml'),
    ],
    sheet: worksheetPart({
      extra: '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
    }),
    sheetRelationships: [relationship('rId1', 'table', '../tables/table1.xml')],
    overrides: [
      ['/xl/tables/table1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'],
    ],
    extraEntries: {
      'xl/tables/table1.xml': strToU8(`${XML_HEADER}\n<table xmlns="${SPREADSHEET_NS}" id="1" name="DataTable" displayName="DataTable" ref="A1:B3" totalsRowShown="0">`
        + '<autoFilter ref="A1:B3"/>'
        + '<tableColumns count="2"><tableColumn id="1" name="Category"/><tableColumn id="2" name="Amount"/></tableColumns>'
        + '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
        + '</table>\n'),
    },
  })
}

function buildDefinedNamesXlsx() {
  return xlsxPackage({
    workbook: workbookPart({
      definedNames: '<definedNames>'
        + '<definedName name="TaxRate">0.2</definedName>'
        + '<definedName name="SalesRange" localSheetId="0">Data!$A$1:$B$4</definedName>'
        + '<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$4</definedName>'
        + '</definedNames>',
    }),
    workbookRelationships: [
      relationship('rId1', 'worksheet', 'worksheets/sheet1.xml'),
      relationship('rId2', 'styles', 'styles.xml'),
    ],
    sheet: worksheetPart(),
  })
}

function buildCalcChainXlsx() {
  return xlsxPackage({
    workbook: workbookPart(),
    workbookRelationships: [
      relationship('rId1', 'worksheet', 'worksheets/sheet1.xml'),
      relationship('rId2', 'calcChain', 'calcChain.xml'),
      relationship('rId3', 'styles', 'styles.xml'),
    ],
    sheet: worksheetPart({
      sheetData: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>SUM(A1:A2)</f><v>3</v></c></row>'
        + '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f>B1*2</f><v>6</v></c></row>',
    }),
    overrides: [
      ['/xl/calcChain.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'],
    ],
    extraEntries: {
      'xl/calcChain.xml': strToU8(`${XML_HEADER}\n<calcChain xmlns="${SPREADSHEET_NS}"><c r="B1" i="1"/><c r="B2" i="1"/></calcChain>\n`),
    },
  })
}

async function buildMacroXlsm() {
  const entries = unzipSync(await buildRoundtripXlsx())
  const contentTypes = strFromU8(entries['[Content_Types].xml'])
    .replace(XLSX_WORKBOOK_CONTENT_TYPE, MACRO_XLSX_CONTENT_TYPE)
    .replace('</Types>', `  <Override PartName="/xl/vbaProject.bin" ContentType="${VBA_CONTENT_TYPE}"/>\n</Types>`)
  entries['[Content_Types].xml'] = strToU8(contentTypes)
  const workbookRelationships = strFromU8(entries['xl/_rels/workbook.xml.rels'])
    .replace('</Relationships>', `  <Relationship Id="rId99" Type="${VBA_RELATIONSHIP}" Target="vbaProject.bin"/>\n</Relationships>`)
  entries['xl/_rels/workbook.xml.rels'] = strToU8(workbookRelationships)
  entries['xl/vbaProject.bin'] = VBA_PROJECT_BYTES
  return buildArchive(entries)
}

function buildStoredContainerDocx() {
  const entries = unzipSync(buildDocx())
  const sorted = {}
  for (const name of Object.keys(entries).sort()) {
    sorted[name] = [entries[name], name === 'word/styles.xml'
      ? { ...ZIP_ENTRY_OPTIONS, level: 0 }
      : ZIP_ENTRY_OPTIONS]
  }
  return zipSync(sorted, { level: 6 })
}

function buildCommentedContainerDocx() {
  return appendArchiveComment(buildDocx(), 'Office Exchange fixture archive comment')
}

function buildDescriptorContainerDocx() {
  return buildDescriptorArchive(unzipSync(buildDocx()))
}

function appendArchiveComment(archive, comment) {
  const commentBytes = strToU8(comment)
  let eocd = -1
  const lowerBound = Math.max(0, archive.length - 22 - 0xffff)
  for (let index = archive.length - 22; index >= lowerBound; index -= 1) {
    if (archive[index] === 0x50 && archive[index + 1] === 0x4b
      && archive[index + 2] === 0x05 && archive[index + 3] === 0x06) {
      eocd = index
      break
    }
  }
  if (eocd < 0) throw new Error('Cannot append an archive comment: EOCD not found')
  const output = new Uint8Array(archive.length + commentBytes.length)
  output.set(archive)
  output.set(commentBytes, archive.length)
  output[eocd + 20] = commentBytes.length & 0xff
  output[eocd + 21] = (commentBytes.length >>> 8) & 0xff
  return output
}

function buildDescriptorArchive(entries) {
  const CRC_TABLE = new Uint32Array(256)
  for (let index = 0; index < 256; index += 1) {
    let value = index
    for (let bit = 0; bit < 8; bit += 1) value = (value & 1) !== 0 ? 0xedb88320 ^ (value >>> 1) : value >>> 1
    CRC_TABLE[index] = value >>> 0
  }
  const crc32 = (bytes) => {
    let crc = 0xffffffff
    for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8)
    return (crc ^ 0xffffffff) >>> 0
  }
  const u16 = (value) => Uint8Array.from([value & 0xff, (value >>> 8) & 0xff])
  const u32 = (value) => Uint8Array.from([
    value & 0xff,
    (value >>> 8) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 24) & 0xff,
  ])
  const concat = (parts) => {
    const total = parts.reduce((size, part) => size + part.length, 0)
    const output = new Uint8Array(total)
    let offset = 0
    for (const part of parts) {
      output.set(part, offset)
      offset += part.length
    }
    return output
  }

  const DOS_TIME = 0
  const DOS_DATE = (46 << 9) | (1 << 5) | 1
  const localParts = []
  const centralParts = []
  let localOffset = 0
  for (const name of Object.keys(entries).sort()) {
    const payload = entries[name]
    const nameBytes = strToU8(name)
    const compressed = deflateSync(payload, { level: 6 })
    const crc = crc32(payload)
    localParts.push(
      u32(0x04034b50), u16(20), u16(0x0808), u16(8), u16(DOS_TIME), u16(DOS_DATE),
      u32(0), u32(0), u32(0), u16(nameBytes.length), u16(0), nameBytes,
      compressed,
      u32(0x08074b50), u32(crc), u32(compressed.length), u32(payload.length),
    )
    centralParts.push(
      u32(0x02014b50), u16(20), u16(20), u16(0x0808), u16(8), u16(DOS_TIME), u16(DOS_DATE),
      u32(crc), u32(compressed.length), u32(payload.length),
      u16(nameBytes.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(localOffset), nameBytes,
    )
    localOffset += 30 + nameBytes.length + compressed.length + 16
  }
  const localBytes = concat(localParts)
  const centralBytes = concat(centralParts)
  return concat([
    localBytes,
    centralBytes,
    u32(0x06054b50), u16(0), u16(0), u16(Object.keys(entries).length), u16(Object.keys(entries).length),
    u32(centralBytes.length), u32(localBytes.length), u16(0),
  ])
}

/**
 * Write or verify one fixture.
 * @param {string} name - Fixture file name.
 * @param {Uint8Array} bytes - Expected fixture bytes.
 * @param {boolean} checkOnly - Compare with the existing file instead of writing.
 * @returns {boolean} True when the fixture matches (check mode) or is written.
 */
function emitFixture(name, bytes, checkOnly) {
  const target = join(FIXTURES_DIR, name)
  if (checkOnly) {
    let existing
    try {
      existing = readFileSync(target)
    } catch {
      console.error(`missing: ${name}`)
      return false
    }
    if (!existing.equals(Buffer.from(bytes))) {
      console.error(`drift: ${name}`)
      return false
    }
    return true
  }
  writeFileSync(target, bytes)
  console.log(`${name}: ${bytes.byteLength} bytes`)
  return true
}

async function main() {
  mkdirSync(FIXTURES_DIR, { recursive: true })
  const checkOnly = process.argv.includes('--check')

  const fixtures = [
    ['roundtrip.xlsx', await buildRoundtripXlsx()],
    ['roundtrip.docx', buildDocx()],
    ['preserve-image.docx', buildImageDocx()],
    ['preserve-external-image.docx', buildImageDocx({ external: true })],
    ['preserve-notes.docx', buildNotesDocx()],
    ['preserve-header-image.docx', buildHeaderImageDocx()],
    ['preserve-footer-image.docx', buildFooterImageDocx()],
    ['preserve-note-image.docx', buildNoteImageDocx()],
    ['preserve-multi-media.docx', buildMultiMediaDocx()],
    ['preserve-table-cell-image.docx', buildTableCellImageDocx()],
    ['offset-shift.docx', buildOffsetShiftDocx()],
    ['preserve-id-collision.docx', buildPreservedIdCollisionDocx()],
    ['preserve-comment.docx', buildCommentDocx()],
    ['preserve-chart.docx', buildChartDocx()],
    ['preserve-embedded-object.docx', buildEmbeddedObjectDocx()],
    ['preserve-shape.docx', buildShapeDocx()],
    ['preserve-external-link.docx', buildExternalLinkDocx()],
    ['preserve-field.docx', buildFieldDocx()],
    ['preserve-bookmark.docx', buildBookmarkDocx()],
    ['preserve-content-control.docx', buildContentControlDocx()],
    ['preserve-unmodeled-parts.docx', buildUnmodeledPartsDocx()],
    ['preserve-pivot-table.xlsx', buildPivotTableXlsx()],
    ['preserve-excel-table.xlsx', buildExcelTableXlsx()],
    ['preserve-defined-names.xlsx', buildDefinedNamesXlsx()],
    ['preserve-calc-chain.xlsx', buildCalcChainXlsx()],
    ['macro.docm', buildMacroDocm()],
    ['macro.xlsm', await buildMacroXlsm()],
    ['container-stored.docx', buildStoredContainerDocx()],
    ['container-comment.docx', buildCommentedContainerDocx()],
    ['container-descriptor.docx', buildDescriptorContainerDocx()],
    ['unsupported-macro.docx', buildDocx({ macro: true })],
    ['unsupported-tracked.docx', buildDocx({ trackedChanges: true })],
    ['unsupported-header.docx', buildDocx({ header: true })],
  ]

  let matches = true
  let checkedCount = 0
  for (const [name, bytes] of fixtures) {
    checkedCount += 1
    matches = emitFixture(name, bytes, checkOnly) && matches
  }

  try {
    const legacy = buildLegacyXls()
    checkedCount += 1
    matches = emitFixture('legacy.xls', legacy, checkOnly) && matches
  } catch (error) {
    if (!checkOnly) {
      const stale = join(FIXTURES_DIR, 'legacy.xls')
      try {
        unlinkSync(stale)
      } catch {
        // No stale file to remove.
      }
    }
    console.warn(`Skipping legacy.xls: ${error instanceof Error ? error.message : String(error)}`)
  }

  if (checkOnly) {
    if (!matches) process.exitCode = 1
    else console.log(`fixtures are deterministic (${checkedCount} checked)`)
  }
}

await main()
