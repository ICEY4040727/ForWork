# Quickstart: In-Place OOXML Save

**Feature**: 002-in-place-ooxml-save

Validation guide only. Implementation lives in `tasks.md`; contracts and entity fields live in [data-model.md](./data-model.md) and [contracts/](./contracts/).

## Prerequisites

- Node >= 24.9 and `pnpm@11.28.5` (`corepack enable` once).
- `pnpm install` at the repository root.
- Windows or macOS for the manual Office evidence. The LibreOffice smoke runs on all three hosts when the engine package is available; on Linux it uses the WASM backend.
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

## S9. Manual Word and Excel matrix

On Windows and macOS, open each fixture in the fixed release subset in desktop Word and Excel.

Expected: no repair prompt, no "unreadable content" dialog, the edit is present, and the preserved construct (for example the macro, chart, or tracked change) still exists. Record the build, host, and result in the PR.

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
