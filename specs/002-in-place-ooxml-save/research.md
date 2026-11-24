# Research: In-Place OOXML Save

**Input**: Feature specification from `/specs/002-in-place-ooxml-save/spec.md`

**Date**: 2026-11-18

## R1. Container layer: a minimal ZIP reader/writer over `node:zlib`

**Decision**: Add a Host-only ZIP layer under `packages/document/office-exchange/src/package/`. The reader indexes every entry and records raw views of its local record (header + compressed data + optional descriptor) and central-directory record. The writer returns the original source bytes when nothing changed; otherwise it copies untouched local records verbatim, writes new records for changed/added entries, patches only the offset field in carried central records, and rebuilds the end-of-central-directory with the original archive comment. Changed entries use `zlib.crc32` and `zlib.deflateRawSync`; stored entries stay stored.

**Rationale**: No installed library can pass through a compressed entry. `fflate.unzipSync` returns decompressed payloads only, and `zipSync` recomputes CRC, recompresses, sorts entries, writes a fixed mtime, and cannot write an archive comment. `yauzl` can expose raw compressed bytes but has no writer and ignores data descriptors. `jszip`, `archiver`, `zip-stream`, and `unzipper` all recompress or extract. `node:zlib` is built in, deterministic for a fixed input and explicit options, and already used in the repo (`packages/api/session-controller/tests/image-read.perf.ts` imports `crc32`).

**Alternatives considered**:
- `fflate` with `level: 0` — stores the deflate stream as content, which corrupts the entry; and it still rebuilds headers.
- `yazl`/`archiver` — not installed (yazl absent) and recompress from uncompressed input.
- Whole-file byte patching in place — violates the atomic replace model and the spec's assumption.

**Evidence**: `fflate@0.8.3` `esm/index.mjs` (unzipSync 2692-2704, zipSync 2321-2329, `wzh` 1893-1905, `wzf` 1933-1939); `@types/node/zlib.d.ts` (`crc32` 289-297, `deflateRawSync` 409-413); `pnpm-lock.yaml` for installed ZIP packages.

## R2. No-op identity is a fast path, not a special case

**Decision**: `OfficePackage.save()` returns the original source bytes when no part is dirty, added, or removed. A part is dirty only when its replacement bytes differ from the source bytes; `replacePart` compares bytes instead of trusting a caller flag.

**Rationale**: Re-serialization determinism cannot reproduce an arbitrary source archive byte-for-byte, and the spec requires a byte-identical no-op save (FR-002, SC-001). Returning the source is both the strongest guarantee and the cheapest path.

**Alternatives considered**:
- Rebuild and hope for identity — fails on comments, extra fields, entry order, and DOS timestamps.
- Skip the write entirely — the spec's Assumption says a no-op save still records the save and advances the version.

## R3. Span scanner: a custom namespace-aware byte scanner

**Decision**: Add `src/source/xml-spans.ts`, a single-pass scanner over the part's UTF-8 bytes that records start/end offsets for the unit kinds: `p`, `tbl`, `tr`, `tc` (WordprocessingML) and `c` (SpreadsheetML). It is quote-aware, skips comments, CDATA, processing instructions, and DOCTYPE, resolves namespace URIs, handles self-closing tags and whitespace before `>`, preserves CRLF and BOM, and fails closed on any ambiguity.

**Rationale**: `fast-xml-parser` 5.11.1 has `captureMetaData` with `startIndex`/`endIndex`, but it unconditionally normalizes `\r\n` and `\r` to `\n` before computing offsets (`OrderedObjParser.js:285-286`), so every offset after the first CRLF is wrong for the source bytes. `fast-xml-validator` exposes only line/column and strips the BOM/declaration before scanning. The existing regex catalog (`docx.ts:1383-1461`) is not quote/comment/CDATA aware and assumes a literal `w:` prefix.

**Alternatives considered**:
- `captureMetaData` as the production source — the CRLF normalization disqualifies it; keep it only as a test oracle on LF-only fixtures.
- Extending the existing regex scanner — cannot satisfy comments, CDATA, attributes containing `>`.

**Evidence**: `fast-xml-parser` `src/xmlparser/OrderedObjParser.js:285-286`, `:342-344`, `:554-556`; `src/fxp.d.ts:745-754`; `fast-xml-validator` `src/fxv.d.ts:98-111`.

## R4. DOCX splice granularity: paragraphs first, table cells only when fully reproducible

**Decision**: Prefer the `<w:p>` span for text edits, including paragraphs inside a table cell, so `w:tcPr`, `w:trPr`, and unmapped cell children stay untouched. Splice a `<w:tc>` span only for a cell-property edit that the writer can reproduce completely; otherwise refuse. Whole-table regeneration is not used when a single unit changed.

**Rationale**: The paragraph span is the smallest unit the model can serialize, and it keeps everything outside it byte-identical. A `<w:tc>` span includes properties the writer may not reproduce. Structural table resize is a coupled change (R7).

**Alternatives considered**:
- Cell-level replacement for text edits — drops unmapped cell children.
- Whole-part regeneration with raw-block carry-over — the current behavior; violates span identity.

## R5. Anchor transport: `data.officeExchange.source` only

**Decision**: Store source identity and per-unit anchors in a provider-owned, plain-JSON namespace inside the editable data object: `data.officeExchange.source` for both `IDocumentData` and `IWorkbookData`.

**Rationale**: Univer's `saveUnit` deep-clones the whole snapshot and replaces only `resources`; `createDocumentSnapshot`/`Workbook` copy unknown top-level keys; the wire type is `JsonValue` with a pass-through `jsonValueOf`; drafts serialize the whole snapshot. By contrast, the `OfficeDocumentSnapshot` wrapper is a declared object whose unknown fields Typert strips; `resources` is replaced on save; `body.payloads` is documented as not saved to disk; paragraphs/table cells have no custom slot except `ICellData.custom`, which auto-fill deletes and copy/paste duplicates.

**Alternatives considered**:
- Wrapper-level field — stripped by the normal codec and never built by the Client.
- Per-paragraph/per-cell fields — no durable slot in documents; cell `custom` is "at your own risk".
- Read-time digest only — lost on Host restart and draft recovery.

**Evidence**: `@univerjs/core` `index.js` 28539-28546 (`saveUnit`), 17107-17121, 24707-24714; `i-document-data.d.ts:26-54`, `:176-207`, `:779-792`; `typedef.d.ts` `IWorkbookData.custom`, `ICellData.custom`; `remote.ts` 917-949, 973-1040; `packages/typert/protocol/README.md:51`.

## R6. Changed-unit detection: canonical digest per unit, not `startIndex`

**Decision**: Each anchor stores the unit key plus a digest of a canonical projection of the unit that excludes positional fields. At save, the writer computes the same digest for each current unit and compares it with the read-time digest. Paragraphs are keyed by `(part, paragraphId)`; table cells by `(part, tableId, row/cell path)`; worksheet cells by `(sheetId, row, column)`.

**Rationale**: `startIndex` is positional and shifts with edits; `paragraphId` is the only persisted identity. Today the provider generates `paragraph-N`, which fails Univer's `isValidParagraphId` (`para_` prefix), so a clone path can replace it. The plan must generate valid `para_*` ids and include the part in the key because header/footer contexts restart their serial counters.

**Alternatives considered**:
- Diff source XML against a re-serialized part — requires re-serializing, which is what the feature forbids.
- Index-based matching — breaks on any insertion before the unit.

**Evidence**: `docx.ts:1087-1104`, `:1763-1815`, `:2089-2098`; `@univerjs/core` `index.js:13211-13222`; `i-document-data.d.ts:1301-1330`.

## R7. Structural edits: a splice plan, with table structure refused

**Decision**: Compute an edit script over the anchored unit sequence: unchanged (carry), replaced (splice), inserted (insert at the neighbouring anchor), deleted (remove the span). If the script cannot be expressed as span operations — table row/column resize, reordering that changes table geometry, a unit whose anchor is missing — refuse and name the construct.

**Rationale**: A single-span replacement cannot express insertion, deletion, or reordering. The common agent and user edits are text/value changes plus row/column insertion at the sheet level; the table geometry case is uncommon and demonstrably unsafe.

**Alternatives considered**:
- Fall back to whole-part regeneration for structural edits — silently violates span identity; use only for source-less snapshots (R11).
- Reconstruct the whole table from the model — drops unmapped properties.

## R8. Auxiliary parts: append-only splices, otherwise untouched

**Decision**: `word/styles.xml`, `word/numbering.xml`, `word/_rels/document.xml.rels`, story rels, and `[Content_Types].xml` are not regenerated. When an edit needs a new style, list, relationship, or content-type declaration, the writer appends the minimal element before the closing tag; when an edit removes the last reference, it removes only that element. `word/settings.xml` is never rewritten.

**Rationale**: Regenerating these parts is the largest source of untouched-byte loss (`styles.xml` is reduced to Normal plus five headings; `settings.xml` to one element; rels ids reallocated). Append-only edits keep every existing byte.

**Alternatives considered**:
- Keep regeneration and accept the loss — contradicts SC-002 and FR-003.
- Full styles rewrite — same problem.

**Evidence**: `docx.ts:2575-2580`, `:3597-3634`, `:3683-3760`, `:3735-3760`.

## R9. XLSX: cell-span splice and inline strings

**Decision**: Resolve worksheet parts through `xl/workbook.xml` + `xl/_rels/workbook.xml.rels`; splice the edited `<c>` element only. When an edited cell's string value changes, write it as `t="inlineStr"` with `<is><t xml:space="preserve">…</t></is>` so `xl/sharedStrings.xml` is untouched. Preserve the cell's `s` style index. Refuse a cell whose source representation is a rich-text shared string, a shared/array formula, or otherwise coupled to the edited value.

**Rationale**: A cell with `t="s"` stores only an index; changing its text requires touching another part or changing the representation. Inline strings are self-contained and valid. `sst@count` becomes advisory-stale by one, which Excel and LibreOffice tolerate; the alternative is a second part splice, and the shared-string index may be referenced by other cells. ExcelJS always emits `r` on rows and cells, but arbitrary files may omit it, so the scanner tracks sequential position.

**Alternatives considered**:
- Append a new `<si>` and repoint — two part splices plus root-attribute updates; still leaves an unreferenced entry.
- Edit the existing `<si>` — changes every cell that references the index.
- Leave the cell as `t="s"` — cannot change the visible text without touching another part.

**Evidence**: `xlsx.ts:497-502`, `:568-619`, `:1997-2011`; ExcelJS `lib/xlsx/xform/sheet/cell-xform.js` (shared strings), `row-xform.js` (`r` attribute); `xlsx-archive.ts` OPC helpers.

## R10. Source transport and save ordering

**Decision**: Extend `OfficeExchange.write(snapshot, format, source, signal)` and return `{ bytes, disposition }`. In `OfficeDocumentRemote.save`: confine; `stat` and require `version === expectedVersion`; `readBytes` under `maxBytes`; `stat` again and re-require the version; run construct detection on the same bytes; call `write`; check the output size; publish with `replaceIfVersion`; record the save fact with the disposition.

**Rationale**: The writer needs the source bytes (FR-013), and the bytes used for detection must be the bytes that are spliced. A single `stat` followed by a path-based detector that re-reads the file can analyze different bytes (existing `refuseUnsupportedSource` behavior). The final version guard remains authoritative, so a change during serialization is still rejected as a conflict.

**Alternatives considered**:
- Carry the source package in the wire snapshot — inflates every read/save/draft payload and duplicates what the Host can read.
- Keep path-based detection — races with the read and violates FR-013.

**Evidence**: `service.ts:39-46`; `provider/index.ts:80-141`, `:188-197`; `remote.ts:124-147`, `:164-193`, `:781-793`; `fs` `index.ts:236-246`, `:292-315`; `fs-local/src/index.ts:252-269`, `:369-388`; `fsio.ts:431-461`, `:75-78`.

## R11. Legacy snapshots and drafts

**Decision**: If a snapshot has no `officeExchange.source`, use the existing regeneration writer and the existing `officeExchange.docx` preserved-package sidecar. New reads always attach `source` and do not attach base64 preserved parts; old drafts continue to load and save through the legacy path with their current behavior.

**Rationale**: Draft recovery happens after a restart, and an in-flight draft created before this change carries no source state. Keeping the old writer for that case is additive and reversible (Principle V). Because the property is additive within the same snapshot kind, no session-format decision is involved.

**Alternatives considered**:
- Refuse old drafts — loses recoverable user work.
- Rebuild the source from the sidecar — the sidecar never carried the whole package.

## R12. Refusal classification: coupled versus independent

**Decision**: During splicing, inspect the source span. Carry independent unmapped properties (cell `s`, `w:tcPr`/`w:trPr` children the serializer does not emit, unmapped `w:pPr`/`w:rPr` children, cell width/fit-to-text) by merging them into the regenerated unit. Refuse when the span contains a construct whose meaning or anchor depends on the edited content: tracked-change and comment ranges overlapping the unit, a content control wrapping the edited text, a field whose result is the edited text, rich-text shared strings, shared/array formulas, and ambiguous `mc:AlternateContent` branches.

**Rationale**: Refusing every unmapped property inside an edited unit would leave many ordinary files unsavable (SC-012), while carrying a coupled construct can produce an invalid or semantically wrong document. The classification is derived from the source span, not from a file-level list.

**Alternatives considered**:
- Always refuse — the residual refusal set stays large.
- Always carry — can invalidate anchors.

## R13. Session save fact: an optional disposition property

**Decision**: Add `readonly disposition?: DocumentSaveDisposition` to `DocumentSaveEventData`, containing exact counts plus capped entry-name lists for spliced/added/removed entries and the verbatim count. Keep `SESSION_FORMAT_VERSION` at 4; the change is the "add an optional event-body property" row. Do not use the envelope `ignorable` flag. Record a persistence acknowledgement and regenerate the schema/catalog. Extend the projection only if the UI needs the last disposition; otherwise document that the fact is audit-only.

**Rationale**: The rule permits same-version changes only when older readers preserve required meaning, ignore optional features, or refuse required ones. An optional property is safely ignorable. The write result must carry the disposition truthfully because only the writer knows which entries changed.

**Alternatives considered**:
- Required property — incompatible with older logs; needs a bump or fallback.
- Tool result only — not durable.

**Evidence**: `packages/core/session/src/document-events.ts:212-232`; `types.ts:87-88`, `:95`; `docs/persistence-changes/README.md:51-61`; `document-facts.ts:37-47`, `:96-103`; `remote.ts:603-632`.

## R14. Format widening: `.docm` and `.xlsm`

**Decision**: Add both extensions everywhere a format is classified: `OfficeFormat`/`EditableOfficeFormat`, `formatOf`, provider dispatch, Remote format handling, the persisted `DocumentFormat`, the projection schema, client tab patterns, icons, preview registration, tool diagnostics, fixtures, and Web tests. Preserve the macro-enabled main content type by never regenerating `[Content_Types].xml` (R8).

**Rationale**: Word and Excel store macros in `.docm`/`.xlsm`; extension-only classification cannot satisfy US2. Widening `DocumentFormat` is a persisted type change and needs its own acknowledgement; the projection cache schema may need a state version bump if an older client cannot validate the new literals.

**Constitution**: The v1 format sentence and the conversion-stack sentence changed; amendment 3.2.0 (2026-11-18) records the widened `.docx/.docm/.xlsx/.xlsm` set and the package-level write layer, with rationale, affected requirements (FR-022, SC-011, plus `001` FR-001/FR-008 supersession), and a migration note (no data migration; old files remain readable).

**Alternatives considered**: read-only macro formats — leaves the refusal case unsolved.

## R15. Validation ladder and evidence limits

**Decision**: Build a corpus of at least 30 fixtures; add a test-local OPC validator (relationship closure, content-type coverage, dangling references, duplicate ids, well-formed XML) and a byte-diff report helper. Use the LibreOffice kit only as a `libreoffice-load-smoke` (import/render), never as proof of "no repair"; the Windows system-LibreOffice matrix opens, renders, and reads back the fixed release subset, and desktop Word/Excel no-repair validation is a disclosed residual risk because it is unavailable on the release host. Extend the Web `document-preview` snapshot to click Save and read back the preserved entry.

**Rationale**: LibreOffice can silently repair a package and returns no repair diagnostics; no OOXML schema validator exists in the repo. Naming the check accurately keeps the evidence claim honest (Principle VI).

**Alternatives considered**: Microsoft Office automation — not portable, unavailable to CI, and not installed on the release host; claiming LibreOffice proves no repair — unsupported by its API.

**Evidence**: `@deepseek-ai/libreoffice-kit` `lib/types/index.d.ts`; `office-to-pdf/src/index.ts:147`; `provider-service.spec.ts` engine probe; `snapshots/web/document-preview/document.expected.md`.

## R16. Performance and bounds

**Decision**: Build one span index per part per save; never rescan per paragraph. Cap the anchor payload; when the package exceeds the anchor budget, re-map the source at save and diff, or refuse loudly. Read the source once per save. Keep `maxFileBytes`, `maxCells`, `maxConversionMs`, and the 10,000-entry scan bound.

**Rationale**: Repeated whole-part scanning is O(n²); a per-cell anchor for the 2M-cell bound would be enormous. The performance target is 10,000 cells, where compact anchors are manageable. Silent anchor dropping would turn a bounded failure into a silent loss.

**Alternatives considered**: No anchors, always diff source — loses identities across drafts and restarts; unbounded anchors — payload and memory risk.
