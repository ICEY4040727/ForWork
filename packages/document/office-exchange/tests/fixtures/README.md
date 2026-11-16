# Office Exchange fixtures

Deterministic binary fixtures for `@forwork/office-exchange` round-trip, mapping, and
unsupported-feature detection tests.

| File | Purpose |
| --- | --- |
| `roundtrip.xlsx` | Editable workbook: `Sales` sheet with a bold header, three data rows, `=SUM(B2:B4)`, a merged cell, and `$#,##0.00` currency formatting; `Notes` sheet with text cells. Detects clean. |
| `roundtrip.docx` | Minimal OOXML document: `Heading1` paragraph, normal paragraph with a bold run, and a 2x2 table. No macros, comments, tracked changes, headers, drawings, or embedded objects. Detects clean. |
| `unsupported-macro.docx` | `roundtrip.docx` plus a `word/vbaProject.bin` entry. Detection must report `macro`. |
| `unsupported-tracked.docx` | `roundtrip.docx` with `<w:ins>` and `<w:del>` revisions in `word/document.xml`. Detection must report `tracked-changes`. |
| `unsupported-header.docx` | `roundtrip.docx` plus `word/header1.xml` and its content-type override and relationship. Detection must report `header-footer`. |
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

Output is byte-for-byte reproducible: archive entries are sorted and written with the fixed
`FIXED_DATE` timestamp, workbook properties are pinned to the same date, and no random ids or
current-time values are used. Re-running the script on a clean tree produces identical hashes.

Do not hand-edit the generated binaries; edit `generate.mjs` and regenerate. The `FIXED_DATE`
global in `generate.mjs` is the single source of determinism and must only change when every
fixture is intentionally re-baselined.

If the installed SheetJS build cannot write `bookType: 'xls'`, the script skips `legacy.xls`,
removes any stale copy, and prints the reason.
