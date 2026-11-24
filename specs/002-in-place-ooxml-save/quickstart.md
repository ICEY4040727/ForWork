# Quickstart: In-Place OOXML Save

**Feature**: 002-in-place-ooxml-save

Validation guide only. Implementation lives in `tasks.md`; contracts and entity fields live in [data-model.md](./data-model.md) and [contracts/](./contracts/).

## Prerequisites

- Node >= 24.9 and `pnpm@11.28.5` (`corepack enable` once).
- `pnpm install` at the repository root.
- Windows for the system-LibreOffice release matrix (S9). The automated LibreOffice smoke (S7) runs on every host where the engine package is available; on Linux it uses the WASM backend. macOS is out of scope for this feature's validation, and desktop Word/Excel is not installed on this host.
- A working `D:\kilo\temp` scratch directory is not required; the test harness owns its temp directories.

## S1. Corpus generation

```sh
node packages/document/office-exchange/tests/fixtures/generate.mjs
git status --short packages/document/office-exchange/tests/fixtures
```

Expected: the second command prints nothing. Any fixture diff means the generator is not deterministic or a fixture was hand-edited.

The corpus must contain at least one fixture per construct class in FR-004 and at least 30 fixtures in total. `tests/fixtures/README.md` is the inventory.

## S2. No-op identity

```sh
pnpm exec vitest run packages/document/office-exchange/tests/ooxml-preservation.spec.ts -t "no-op identity"
```

Expected: for every corpus fixture, `write(read(source), source)` returns bytes equal to `source`. The failure output is the byte-diff report: entry name, status, sizes, digests, and first differing offset.

## S3. Single-span splice

```sh
pnpm exec vitest run packages/document/office-exchange/tests/ooxml-preservation.spec.ts -t "splice"
```

Expected: each case changes exactly the expected part and span. Every other entry's local record is byte-identical, and every byte outside the reported spans in the edited part is identical. The report lists `verbatim`, `spliced`, `added`, and `removed` entries.

## S4. Unmodeled constructs

```sh
pnpm exec vitest run packages/document/office-exchange/tests/ooxml-preservation.spec.ts -t "preserved construct"
```

Expected, per construct: macro, chart, pivot table, Excel table, defined name, calc chain, comment, embedded object, drawing, theme, font table, custom XML, external link, content control, field, bookmark, and document properties survive an edit elsewhere. No save is refused for a construct outside the edited range.

## S5. Refusal postcondition

```sh
pnpm exec vitest run packages/document/office-exchange/tests/write-postcondition.spec.ts
```

Expected: a tracked change, comment anchor, content control, or field that overlaps the edited span refuses with the construct named; the source SHA-256 is unchanged; no `document/save` fact exists; the projection version does not advance.

## S6. Compatibility paths

```sh
pnpm exec vitest run packages/document/office-exchange/tests/remote.spec.ts \
  packages/document/office-exchange/tests/remote-edges.spec.ts
```

Expected: a snapshot without `officeExchange.source` saves through the legacy writer; a draft written by the previous binary restores and saves; a second save in the same window with stale anchors is refused as a conflict rather than spliced against the wrong source.

## S7. LibreOffice load smoke

```sh
pnpm exec vitest run packages/document/office-exchange/tests/ooxml-engine.spec.ts
```

Expected: saved DOCX and XLSX fixtures import and render through `@deepseek-ai/libreoffice-kit`, and the output PDF starts with `%PDF-`. Hosts without the engine skip with the kit's `unavailable` message. This is an import/render smoke, not a "no repair prompt" proof.

## S8. Web save snapshot

```sh
DSH_SNAPSHOT=refresh pnpm run test:web
```

Then inspect `snapshots/web/document-preview/document.expected.md`.

Expected: the Office section records a completed save and an unchanged preserved part, and no `data-office-save-error` banner is rendered. Re-running without `DSH_SNAPSHOT=refresh` passes.

## S9. Windows system-LibreOffice matrix

This replaces the former macOS and desktop-Word/Excel matrix. Produce the saved outputs either through the desktop editor or through the exchange writer with the recorded edit, and record which producer was used for each fixture; never validate an unsaved fixture. Use `soffice.com` with a private profile so the run does not share a profile with an interactive session.

| fixture | edit |
| --- | --- |
| roundtrip.docx | Quarterly→Financial |
| macro.docm | Quarterly→Financial |
| preserve-chart.docx | Chart →Graph |
| preserve-comment.docx | insert paragraph "Inserted paragraph" |
| preserve-unmodeled-parts.docx | Unmodeled parts fixture→Carried parts fixture |
| unsupported-tracked.docx | Quarterly→Financial |
| container-descriptor.docx | Quarterly→Financial |
| roundtrip.xlsx | first sheet A2=Sprocket |
| macro.xlsm | first sheet A2=Sprocket |
| preserve-pivot-table.xlsx | first sheet A2=Gamma |
| preserve-calc-chain.xlsx | first sheet A1=5 |
| preserve-excel-table.xlsx | first sheet A2=Gamma |

PDF render and read-back:

```sh
& "D:\Program Files\LibreOffice\program\soffice.com" --headless --norestore -env:UserInstallation=file:///D:/kilo/temp/lo-s9-profile --convert-to pdf --outdir D:\kilo\temp\s9\pdf <saved>
& "D:\Program Files\LibreOffice\program\soffice.com" --headless --norestore -env:UserInstallation=file:///D:/kilo/temp/lo-s9-profile --convert-to "txt:Text (encoded):UTF8" --outdir D:\kilo\temp\s9\txt <saved>
& "D:\Program Files\LibreOffice\program\soffice.com" --headless --norestore -env:UserInstallation=file:///D:/kilo/temp/lo-s9-profile --convert-to csv --outdir D:\kilo\temp\s9\csv <saved>
```

Use the text command for `.docx`/`.docm` and the CSV command for `.xlsx`/`.xlsm`. Expected: exit 0, the PDF starts with `%PDF-`, the edited text or value is present in the read-back, and for `.docm`/`.xlsm` the VBA project entry and the macro-enabled content type are still present. Record per fixture: fixture, producer, saved-output SHA-256, LibreOffice build, Windows build, PDF result, read-back result, and preserved-construct result.

A result is import/render/read-back evidence, not a no-repair proof. Desktop Word and Excel are not installed on this host, so their no-repair behavior is a disclosed residual risk, never a matrix claim.

Observed limitation: LibreOffice 26.8.3.1 hangs on the direct DOCX→txt filter for table-bearing documents (`roundtrip.docx`, `macro.docm`, `container-descriptor.docx`). Use the DOCX→ODT→txt fallback for those fixtures; the untouched control also hangs, so this is an engine filter defect, not a writer failure.

## S10. Focused coverage and build checks

```sh
pnpm exec vitest run packages/document/office-exchange/tests/ooxml-preservation.spec.ts \
  packages/document/office-exchange/tests/ooxml-engine.spec.ts \
  packages/document/office-exchange/tests/write-postcondition.spec.ts \
  packages/document/office-exchange/tests/remote.spec.ts \
  --coverage \
  --coverage.include='packages/document/office-exchange/src/**/*.ts'
```

```sh
pnpm run typecheck
pnpm run lint
pnpm run build:artifacts
pnpm run build:check
```

Expected: focused coverage is 100% per file for `packages/document/office-exchange/src`; typecheck, lint, and the built-source check pass. Run `pnpm run test:coverage` once before landing and let CI own the full matrix.

## S11. Documentation and persistence gates

```sh
pnpm run doc-sync
pnpm run hygiene
```

Expected: the new persistence records for `event:document/save` and the widened `DocumentFormat` are present and regenerated; README, subsystem docs, generated catalog, and both language sides agree; no package/exports/dependency check fails.

## Failure triage

- A no-op identity failure means an entry was re-encoded or reordered; inspect the local-record comparison first.
- A splice failure with an accepted save usually means the span scanner selected the wrong range; check for comments, CDATA, CRLF, namespace prefixes, and `mc:AlternateContent`.
- A refusal on a construct that is clearly outside the edited span means anchor or digest mapping is wrong; do not relax the refusal, fix the mapping.
- A LibreOffice failure with `invalid-document` indicates a structurally invalid package; run the OPC validator before changing the engine smoke.
