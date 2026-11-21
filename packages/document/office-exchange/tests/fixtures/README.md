# Office Exchange fixtures

English | [中文](README.zh.md)

Deterministic binary fixtures for `@forwork/office-exchange` round-trip, mapping, preservation,
and unsupported-feature detection tests. The corpus contains 34 fixtures and covers every
construct class in FR-004 of `specs/002-in-place-ooxml-save`.

| File | Purpose |
| --- | --- |
| `roundtrip.xlsx` | Editable workbook: `Sales` sheet with a bold header, three data rows, `=SUM(B2:B4)`, a merged cell, and `$#,##0.00` currency formatting; `Notes` sheet with text cells. Detects clean. |
| `roundtrip.docx` | Minimal OOXML document: `Heading1` paragraph, normal paragraph with a bold run, and a 2x2 table. No macros, comments, tracked changes, headers, drawings, or embedded objects. Detects clean. |
| `preserve-image.docx` | Inline picture with an embedded JPEG (`r:embed`), a `word/media/` part, its image relationship, and a JPEG content type. Detection must stay clean and the picture, media bytes, and relationship must survive a save. |
| `preserve-external-image.docx` | Inline picture linked to an external target (`r:link`, `TargetMode="External"`) with no media part. Detection must stay clean and the picture and external relationship must survive a save. |
| `preserve-notes.docx` | Body with `w:footnoteReference` and `w:endnoteReference` markers plus `word/footnotes.xml` and `word/endnotes.xml`, their relationships, and content types. Detection must stay clean and the note parts and reference runs must survive a save. |
| `preserve-header-image.docx` | Header part with an inline PNG, a header relationship, and a header-to-media relationship. Detection must stay clean and the header picture must survive a save. |
| `preserve-footer-image.docx` | Footer part with an inline PNG, a footer relationship, and a footer-to-media relationship. Detection must stay clean and the footer picture must survive a save. |
| `preserve-note-image.docx` | Footnote containing an inline PNG with a footnote relationship and media part, plus an endnote with text. Detection must stay clean and the note picture and parts must survive a save. |
| `preserve-multi-media.docx` | One paragraph with inline PNG, JPEG, and EMF pictures, three image relationships, and three media parts. Detection must stay clean and every media entry must survive a save. |
| `preserve-table-cell-image.docx` | Inline PNG inside a table cell. Detection must stay clean and the cell picture, its relationship, and its media part must survive a save. |
| `offset-shift.docx` | CRLF document with two paragraphs before a paragraph containing an inline PNG, so the preserved run starts at a nonzero byte offset. Detection must stay clean and the picture must survive an edit elsewhere. |
| `preserve-id-collision.docx` | Document whose preserved image uses `rId2` and whose header uses `rId3`, colliding with ids a regeneration writer typically allocates. Detection must stay clean and each relationship must keep its original id. |
| `preserve-comment.docx` | Body with `w:commentRangeStart`, `w:commentRangeEnd`, and `w:commentReference`, plus `word/comments.xml`, `word/commentsExtended.xml`, `word/commentsIds.xml`, and `word/people.xml` with their relationships and content types. Detection must report `comment`, and the threaded-comment parts must survive a save. |
| `preserve-chart.docx` | Body with a chart drawing and `word/charts/chart1.xml`, its relationship, and content type. The chart part must survive a save. |
| `preserve-embedded-object.docx` | Body with a VML `w:object` preview and `word/embeddings/oleObject1.bin`, its OLE relationship, image relationship, and content types. Detection must report `embedded-object`. |
| `preserve-shape.docx` | Body with a VML `w:pict` shape and `word/drawings/vmlDrawing1.vml`, its relationship, and content type. Detection must report `shape`. |
| `preserve-external-link.docx` | Body hyperlink to an external target with `TargetMode="External"`. Detection must stay clean and the link target and relationship must survive a save. |
| `preserve-field.docx` | Body with a `w:fldSimple` field and a complex field built from `w:fldChar` and `w:instrText`. The field runs must survive a save. |
| `preserve-bookmark.docx` | Body with `w:bookmarkStart`, `w:bookmarkEnd`, and a `w:hyperlink` using `w:anchor`. The bookmark markers must survive a save. |
| `preserve-content-control.docx` | Body with a `w:sdt` content control containing a paragraph. The content control must survive a save. |
| `preserve-unmodeled-parts.docx` | Document plus `word/theme/theme1.xml`, `word/fontTable.xml`, `word/webSettings.xml`, `docProps/core.xml`, `docProps/app.xml`, `customXml/item1.xml`, and `customXml/itemProps1.xml` with their relationships and content types. Every unmodeled part must survive a save. |
| `preserve-pivot-table.xlsx` | Workbook with a pivot table, pivot cache definition and records, their relationships and content types. All pivot entries must survive a save. |
| `preserve-excel-table.xlsx` | Worksheet with an Excel table part, its relationship, and content type. The table part must survive a save. |
| `preserve-defined-names.xlsx` | Workbook with a constant defined name, a range defined name, and a print-area defined name. The defined names must survive a save. |
| `preserve-calc-chain.xlsx` | Workbook with formula cells and `xl/calcChain.xml` plus its relationship and content type. The calc chain must survive a save. |
| `macro.docm` | Macro-enabled document with the macro-enabled main content type, `word/vbaProject.bin`, its content-type override, and its relationship. Detection must report `macro` and a save must not change the format. |
| `macro.xlsm` | Macro-enabled workbook with the macro-enabled main content type, `xl/vbaProject.bin`, its content-type override, and its relationship. Detection must report `macro` and a save must not change the format. |
| `container-stored.docx` | `roundtrip.docx` package with `word/styles.xml` written as a stored entry. No-op saves must preserve the stored method and the local record. |
| `container-comment.docx` | `roundtrip.docx` package with a non-empty archive comment. No-op saves must preserve the comment bytes. |
| `container-descriptor.docx` | `roundtrip.docx` package whose entries use non-ZIP64 data descriptors (general-purpose flag bit 3). No-op saves must preserve every descriptor. |
| `unsupported-macro.docx` | `roundtrip.docx` plus a `word/vbaProject.bin` entry, its content-type override, and its relationship. Detection must report `macro`. |
| `unsupported-tracked.docx` | `roundtrip.docx` with `<w:ins>` and `<w:del>` revisions in `word/document.xml`. Detection must report `tracked-changes`. |
| `unsupported-header.docx` | `roundtrip.docx` plus `word/header1.xml` with `Confidential header` text and its content-type override and relationship. Detection must report nothing, and the header text must survive a save. |
| `legacy.xls` | BIFF8 workbook written by SheetJS with one `Legacy` sheet. Exercises legacy conversion. |
| `generate.mjs` | Regenerates every fixture above. |

## Regenerate

```sh
node packages/document/office-exchange/tests/fixtures/generate.mjs
```

`generate.mjs` resolves `exceljs` and `xlsx` from
`packages/client/ui-sidebar-documentpreview` (their owning workspace package) and `fflate` from
`packages/document/office-exchange`, falling back to the repository root.

## Determinism

```sh
node packages/document/office-exchange/tests/fixtures/generate.mjs --check
```

`--check` rebuilds every fixture in memory and compares it with the checked-in binary without
writing. It exits non-zero and prints `missing:` or `drift:` for any fixture that differs; a clean
tree prints `fixtures are deterministic (34 checked)`. Run it after editing `generate.mjs` or any
fixture assertion.

Output is byte-for-byte reproducible: archive entries are sorted and written with the fixed
`FIXED_DATE` timestamp, workbook properties are pinned to the same date, and no random ids or
current-time values are used. Re-running the script on a clean tree produces identical hashes.

Do not hand-edit the generated binaries; edit `generate.mjs` and regenerate. The `FIXED_DATE`
global in `generate.mjs` is the single source of determinism and must only change when every
fixture is intentionally re-baselined.

If the installed SheetJS build cannot write `bookType: 'xls'`, the script skips `legacy.xls`,
removes any stale copy, and prints the reason.
