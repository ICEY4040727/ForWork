---

description: "Task list for the In-Place OOXML Save feature"
---

# Tasks: In-Place OOXML Save

**Input**: Design documents from `/specs/002-in-place-ooxml-save/`

**Prerequisites**: plan.md, spec.md, research.md, data-model.md, contracts/, quickstart.md

**Tests**: Included. The repository mandates behavior tests and a per-file coverage gate; the feature's acceptance is byte-level, so tests are the primary deliverable, not an optional add-on.

**Organization**: Tasks are grouped by user story. US2 and US3 build on the US1 writer (US2 needs the intersection classification, US3 needs the plan's dispositions); each story remains independently testable once its prerequisites land.

## Format: `[ID] [P?] [Story] Description`

- **[P]**: Can run in parallel (different files, no dependency on an incomplete task)
- **[Story]**: US1, US2, US3 for story phases only
- Every task names its exact file path

## Path Conventions

Monorepo: workspace packages under `packages/<group>/<pkg>/`, desktop and Web under `apps/`, spec artifacts under `specs/002-in-place-ooxml-save/`. The new writer is Host-only; nothing under `src/package/` or `src/source/` may be imported by a Client package.

---

## Phase 1: Setup (Shared Infrastructure)

**Purpose**: Establish the baseline, the module skeleton, the test harness, and the fixture corpus everything else uses.

- [x] T001 Confirm the baseline and record it in `specs/002-in-place-ooxml-save/quickstart-results.md`: run `corepack pnpm install` and `pnpm exec vitest run packages/document/office-exchange/tests`; record the commands and results. No source change.
- [x] T002 [P] Create `packages/document/office-exchange/src/package/zip-format.ts`: ZIP signatures (`0x04034b50`, `0x02014b50`, `0x06054b50`, `0x08074b50`, `0x07064b50`, `0x06064b50`), flag bits (`bit 0`, `bit 3`, `bit 6`, `bit 11`, `bit 13`), methods (`0`, `8`), DOS time helpers, and bounds guards.
- [x] T003 [P] Create `packages/document/office-exchange/src/package/types.ts` with `ZipEntryRecord`, `ZipArchive`, `SourcePackage`, `ZipChanges`, and `ZipWriteOptions` exactly as specified in `specs/002-in-place-ooxml-save/data-model.md` (`localRecord`, `centralRecord`, `descriptorLength`, `localOrder`, `centralOrder`, `comment`, `digest`).
- [x] T004 [P] Create the test harness `packages/document/office-exchange/tests/ooxml-package.ts`: `unzipEntries(bytes)`, `sha256Hex(bytes)`, `comparePackages(source, written, expectation)` returning the JSON report shape in `quickstart.md`, `validateOpc(entries)` returning violations, and `formatDiffReport(report)` rendering the Markdown table.
- [x] T005 [P] Extend `packages/document/office-exchange/tests/fixtures/generate.mjs` to build the full corpus enumerated in `packages/document/office-exchange/tests/fixtures/README.md` (at least 30 fixtures, one per FR-004 class): `.docm` and `.xlsm` with macro-enabled content types plus `vbaProject.bin`; header-image, footer-image, note-image, multi-media, table-cell-image, offset-shift, and preserved-id-collision DOCX cases; comment, chart, pivot-table, Excel-table, defined-names, calc-chain, embedded-object, shape, external-link, field, bookmark, and content-control cases; an unmodeled-parts case (`word/theme/theme1.xml`, `word/fontTable.xml`, `word/webSettings.xml`, `docProps/core.xml`, `docProps/app.xml`, `customXml/item1.xml`); and container cases (stored entry, archive comment, non-ZIP64 data descriptor). Regenerate the binaries and update the fixture README and its pair.

---

## Phase 2: Foundational (Blocking Prerequisites)

**Purpose**: Container layer, span layer, source identity, contract changes, and compatibility switch. Nothing in any user story can start until this phase is complete.

**CRITICAL**: No user story work can begin until this phase is complete.

- [x] T006 Implement the ZIP reader in `packages/document/office-exchange/src/package/zip-reader.ts`: locate EOCD with the comment-length consistency check; reject ZIP64, multi-disk, encryption, unsupported methods, and out-of-bounds ranges; parse every local and central record; validate local/central agreement; compute the descriptor length when bit 3 is set; build the ASCII case-folded part map. Tests in `packages/document/office-exchange/tests/zip-reader.spec.ts`.
- [x] T007 [P] Implement the ZIP writer in `packages/document/office-exchange/src/package/zip-writer.ts`: return the source bytes when nothing changed; copy untouched `localRecord` bytes; rewrite changed records from the original header template (clear bit 3, `zlib.crc32`, `zlib.deflateRawSync` or stored); patch only the offset in carried central records; append new entries deterministically; write the EOCD with the original comment. Tests in `packages/document/office-exchange/tests/zip-writer.spec.ts`.
- [x] T008 [P] Implement `packages/document/office-exchange/src/package/ooxml-package.ts`: `open(source, limits)`, `readPart(name)` (inflate and verify length/CRC), `replacePart(name, bytes)` (byte-compare before marking dirty), `addPart`, `removePart`, `save(options)` with the no-op fast path, and rejection of duplicate or ASCII case-equivalent OPC names. Tests in `packages/document/office-exchange/tests/ooxml-package.spec.ts`.
- [x] T009 [P] Implement `packages/document/office-exchange/src/source/xml-spans.ts`: single-pass, UTF-8 byte-offset scanner for `p`, `tbl`, `tr`, `tc`, and `c`; quote/comment/CDATA/PI/DOCTYPE aware; namespace-URI resolved; self-closing and whitespace-before-`>` handled; ambiguous spans report `ambiguous-span` and never guess. Tests in `packages/document/office-exchange/tests/xml-spans.spec.ts` covering attributes containing `>`, comments, CDATA, prefix rebinding, CRLF, BOM, and `mc:AlternateContent`.
- [x] T010 [P] Implement `packages/document/office-exchange/src/source/unit-digest.ts`: canonical projections for paragraph, table cell, and sheet cell from `data-model.md`; SHA-256 hex over deterministic JSON; excludes `startIndex`, array positions, and snapshot ids. Tests in `packages/document/office-exchange/tests/unit-digest.spec.ts`.
- [x] T011 [P] Implement `packages/document/office-exchange/src/source/source-state.ts`: `OfficeSourceState` schema, `UnitAnchor`, `SheetSummary`, validation rules from `data-model.md` (part exists, `0 ≤ start < end ≤ partByteLength`, non-overlapping spans, unique keys, 64-hex digest, anchor count/byte budget), and a typed `PreservationFailure`. Tests in `packages/document/office-exchange/tests/source-state.spec.ts`.
- [x] T012 Add validated Config fields `maxSourceAnchors` (default `100_000`) and `maxSourceAnchorBytes` (default `8 * 1024 * 1024`) to `packages/document/office-exchange/src/config.ts` and the `Config` schema in `packages/document/office-exchange/src/provider/index.ts`; update `packages/document/office-exchange/tests/provider-service.spec.ts` for invalid values.
- [x] T013 Extend the seam in `packages/document/office-exchange/src/service.ts`, `src/snapshots.ts`, `src/types.ts`, and `src/index.ts`: `write(snapshot, format, source, signal)` returning `{ bytes, disposition }`; add `SaveDisposition` and `PreservedConstructs`; add `detectPreserved(bytes, config?)`. Update `packages/document/office-exchange/src/provider/index.ts` dispatch signatures.
- [x] T014 Implement the Remote save ordering in `packages/document/office-exchange/src/remote.ts`: confine, `stat` and require `version === expectedVersion`, `readBytes` under `maxBytes`, `stat` again, detect on the same bytes, call `write` with the source, check output size, publish with `replaceIfVersion`, record the fact. Update `packages/document/office-exchange/tests/remote.spec.ts` and `tests/remote-edges.spec.ts` stubs.
- [x] T015 Add the optional `disposition` property to `DocumentSaveEventData` in `packages/core/session/src/document-events.ts` and forward it in `packages/session/session-document/src/document-facts.ts`; keep `SESSION_FORMAT_VERSION` at 4. Create the persistence acknowledgement files under `docs/persistence-changes/` for `event:document/save` by following `docs/cookbook/reviewing-persistence-type-changes.md` (prose, JSON, and generated catalog), then run `pnpm run doc-sync`.
- [x] T016 Widen the persisted format type: add `docm`/`xlsm` to `DocumentFormat` in `packages/core/session/src/document-events.ts`, `DocumentProjectionEntry.format` and its zod schema in `packages/session/session-document/src/types.ts` / `document-projection.ts`, and create the persistence acknowledgement for `event:document/open`. Bump `documentProjectionDefinition.stateVersion` from 1 to 2 because an older client's union rejects the new literals; update `sameEntry`, the schema, and `packages/session/session-document/tests/`.
- [x] T017 Implement `detectPreserved(bytes, config?)` in `packages/document/office-exchange/src/detect.ts`: report preserved construct classes and coupled markers separately from the editable set; do not return a whole-file refusal list. Update `packages/document/office-exchange/tests/detect.spec.ts`.
- [x] T018 Attach `data.officeExchange.source` on read in `packages/document/office-exchange/src/provider/docx.ts` and `src/provider/xlsx.ts`, and plumb it through `packages/document/office-exchange/src/provider/index.ts`. DOCX: paragraph/table-cell anchors from `xml-spans.ts` plus `unit-digest.ts`, keyed by `(part, paragraphId)` and `(part, tableId, row/cell path)`; generate valid `para_*` paragraph ids. XLSX: per-cell anchors when under `maxSourceAnchors`, otherwise `sheet-summary` mode. Validate before returning. Tests in `packages/document/office-exchange/tests/source-state.spec.ts` and `tests/provider-service.spec.ts`.
- [x] T019 Select the legacy fallback writer when `officeExchange.source` is absent (snapshots and drafts from before this change) in `packages/document/office-exchange/src/provider/index.ts`; keep the existing regeneration path and `officeExchange.docx` sidecar for that case. Tests in `packages/document/office-exchange/tests/offline.spec.ts` and `tests/remote-edges.spec.ts`; extend `offline.spec.ts` so the splice path, not only the legacy writer, runs under the transport guard.
- [x] T020 Switch the `.docm`/`.xlsm` *provider* surfaces: `packages/document/office-exchange/src/types.ts` (`OfficeFormat`, `EditableOfficeFormat`), `src/format.ts` (`EDITABLE`, `formatOf`), `src/provider/index.ts` (read/write dispatch, `requireFormat`), `src/remote.ts` (`editableDocumentFormat`, `nativeSnapshot`). Tests in `packages/document/office-exchange/tests/provider-service.spec.ts` and `tests/detect.spec.ts`.

**Checkpoint**: Container, spans, source identity, contract, session field, format types, and fallback are ready. User stories can start.

---

## Phase 3: User Story 1 - Save one edited region without disturbing the rest of the package (Priority: P1) MVP

**Goal**: An edit to one paragraph or cell rewrites only that span; every other byte is preserved; a no-op save is byte-identical.

**Independent Test**: Run `pnpm exec vitest run packages/document/office-exchange/tests/ooxml-preservation.spec.ts` and confirm no-op identity, untouched-local-record identity, span minimality, OPC validity, determinism, and the LibreOffice load smoke.

### Tests for User Story 1

> Write these first and confirm they fail against the regeneration writer.

- [x] T021 [P] [US1] Add the corpus identity matrix in `packages/document/office-exchange/tests/ooxml-preservation.spec.ts`: for every fixture, no-op save equals the source; every untouched local record is byte-identical; the diff report accounts for every difference; and the corpus contains at least 30 fixtures with at least one per FR-004 construct class.
- [x] T022 [P] [US1] Add DOCX splice tests in `packages/document/office-exchange/tests/docx-splice.spec.ts`: paragraph text edit, paragraph inside a table cell, header and footer paragraph, a paragraph containing a raw drawing block, paragraph insert, paragraph delete, reordering refusal, and an edited paragraph that carries an unmapped `w:rPr`/`w:pPr` child which must be re-emitted (SC-012).
- [x] T023 [P] [US1] Add XLSX splice tests in `packages/document/office-exchange/tests/xlsx-splice.spec.ts`: numeric cell, formula cell, style-only change, string change to `inlineStr`, merged-cell master, date cell, worksheet row insert/delete, column insert/delete, and an edited cell that keeps its `s` style index and an independent unmapped property (SC-012).
- [x] T024 [US1] Add OPC validity and determinism assertions to `packages/document/office-exchange/tests/ooxml-preservation.spec.ts`: relationship closure, content-type coverage, no dangling `r:id`/`r:embed`, well-formed XML, and identical bytes for repeated writes.

### Implementation for User Story 1

- [x] T025 [US1] Implement the DOCX splice writer in `packages/document/office-exchange/src/provider/docx.ts`: dispatch by `paragraphId` and table-cell path, serialize one paragraph from the model, merge independent unmapped properties, re-emit `docxRawBlocks` inside the regenerated paragraph, and refuse a unit whose span is ambiguous.
- [x] T026 [US1] Implement DOCX auxiliary append-only splices in `packages/document/office-exchange/src/provider/docx.ts`: append a `w:style` to `word/styles.xml`, a `w:num`/`w:abstractNum` to `word/numbering.xml`, a `Relationship` before `</Relationships>`, and an `Override`/`Default` before `</Types>`; remove only the elements the edit orphaned. Never regenerate these parts.
- [x] T027 [US1] Implement the XLSX cell-span writer in `packages/document/office-exchange/src/provider/xlsx.ts`: resolve worksheet parts through `xl/workbook.xml` + `xl/_rels/workbook.xml.rels`, splice the edited `<c>`, write changed strings as `t="inlineStr"`, preserve the `s` style index, track sequential row/cell position when `r` is absent, and refuse rich-text/shared/array formula cells.
- [x] T028 [US1] Implement the edit-script builder in `packages/document/office-exchange/src/source/splice.ts`: compare canonical digests, emit `carry`/`replace`/`insert`/`delete`/`append` operations, sort and de-overlap per part, enforce minimality and total coverage, and return a plan with `touchedParts` and per-entry `PartDisposition`.
- [x] T029 [US1] Implement construct intersection and independent-property merge in `packages/document/office-exchange/src/source/splice.ts` and the two providers: classify coupled constructs from the source span (tracked change, comment range, content control wrapping the text, field owning the text, rich-text shared string, shared/array formula, ambiguous `mc:AlternateContent`) and return `ConstructIntersection`; carry independent properties (cell `s`, `w:tcPr`/`w:trPr` children, unmapped `w:pPr`/`w:rPr` children).
- [x] T030 [US1] Wire the plan into the writers in `packages/document/office-exchange/src/provider/index.ts`, `src/provider/docx.ts`, and `src/provider/xlsx.ts`: open the source package, apply the plan through `ooxml-package.ts`, validate the output with `fast-xml-validator` and the OPC checks, and return `{ bytes, disposition }`.
- [x] T031 [US1] Enforce bounds and typed failures in `packages/document/office-exchange/src/package/zip-writer.ts` and `src/provider/index.ts`: output size over `maxFileBytes`, entry count over the scan bound, anchor budget over `maxSourceAnchors`/`maxSourceAnchorBytes`, and unsupported container structures map to `bound-exceeded`/`conversion-failed` with no write.
- [x] T032 [US1] Refresh source anchors across saves in `packages/document/office-exchange/src/remote.ts` and the provider entry. Implemented by re-anchoring from the current source when the sidecar digest is stale (the version guard already proved the file is the one the caller expects), so a second edit/save in the same window succeeds; a malformed sidecar or unanchorable unit refuses. Tests in `packages/document/office-exchange/tests/remote.spec.ts`. Add the FR-014 editor and draft round-trip tests: mount Univer in `packages/client/ui-office-editor/tests/office-roundtrip.spec.tsx`, edit, call `unit.save()`, and assert `data.officeExchange.source` survived; write, read, and recover a draft in `packages/document/office-exchange/tests/source-state.spec.ts`.
- [x] T033 [US1] Add the LibreOffice load smoke in `packages/document/office-exchange/tests/ooxml-engine.spec.ts` using the existing engine probe pattern: render one saved DOCX and one saved XLSX, assert `%PDF-`, and skip with the kit's `unavailable` message. Name the test `libreoffice-load-smoke`.
- [x] T034 [US1] Extend the Web save scenario in `apps/web/tests/document-preview.e2e.ts` to click `[data-office-save]`, read the saved file from `workspaceCwd`, assert a preserved media entry is byte-identical, and add one stable line to `snapshots/web/document-preview/document.expected.md`; refresh with `DSH_SNAPSHOT=refresh pnpm run test:web`.

**Checkpoint**: US1 is independently demonstrable — open, edit one region, save, reopen, only the edited span differs.

---

## Phase 4: User Story 2 - Complex files stop being refused (Priority: P1)

**Goal**: A construct outside the edited range never blocks a save; only a coupled construct inside it refuses. `.docm`/`.xlsm` are editable under the same rules.

**Independent Test**: For every construct class, open, edit elsewhere, save, and confirm the save is accepted and the construct is byte-identical; then edit a coupled construct and confirm a named refusal.

**Dependency**: Requires T029 and T030 (intersection classification and plan) from US1.

### Tests for User Story 2

- [x] T035 [P] [US2] Add preserved-construct cases in `packages/document/office-exchange/tests/ooxml-preservation.spec.ts`: macro, chart, pivot table, Excel table, defined names, calc chain, comment, embedded object, drawing, theme, font table, web settings, custom XML, external link, content control, field, bookmark, document properties.
- [x] T036 [P] [US2] Add coupled-refusal cases in `packages/document/office-exchange/tests/write-postcondition.spec.ts`: tracked change over the edited text, comment anchor inside the range, content control wrapping the text, field owning the text, rich-text shared string, shared formula; each must refuse with the construct named, leave the source SHA-256 unchanged, and append no save fact.
- [x] T037 [P] [US2] Add macro-enabled cases in `packages/document/office-exchange/tests/provider-docm.spec.ts`: read, edit elsewhere, save, assert the VBA project entry and macro-enabled main content type are byte-identical, and assert the output is not converted to a macro-free type.

### Implementation for User Story 2

- [x] T038 [US2] Replace the whole-file save gate in `packages/document/office-exchange/src/remote.ts`: delete `refuseUnsupportedSource` and refuse only on a `ConstructIntersection` returned by the write plan. Update its JSDoc and the Remote error mapping.
- [x] T039 [US2] Remove the whole-file edit gate in `packages/document/tool-office/src/edit.ts` and the read-side refusal in `src/read.ts`; keep the typed `unsupported-feature` result for an intersection refusal. Update `packages/document/tool-office/tests/document-tools.spec.ts` and add an agent-save test proving a non-intersecting save preserves bytes while a coupled edit is refused (FR-021).
- [x] T040 [US2] Update the client gates in `packages/client/ui-office-editor/src/client/useOfficeDocument.ts` (`status: 'unsupported'` and `canSave`) and `command-bridge.ts` (`target.unsupported()`), so a preserved construct does not disable saving or block commands. Update `packages/client/ui-office-editor/tests/use-office-document.spec.tsx` and `editor.spec.tsx`.
- [x] T041 [US2] Update the refusal surface in `packages/client/ui-office-editor/src/client/OfficePanel.tsx` and `locales.ts` (all eight dictionaries): separate "preserved, not editable" from "this edit is refused", following FR-019/FR-020; use the `dsh-client-ui-ux` and `dsh-prose-standard` skills for copy and layout. Add the SC-013 assertion in `packages/client/ui-office-editor/tests/editor.spec.tsx`: every preserved class shows the localized marker and the save control stays enabled.
- [x] T042 [US2] Add `.docm`/`.xlsm` client surfaces: `packages/client/ui-office-editor/src/client/definition.ts` patterns, `target.ts` (ensure they are not legacy-converted), `useOfficeDocument.ts` (carry the opened format instead of deriving it from the snapshot kind), and `packages/client/ui-primitives/src/FileTypeIcon.tsx`. Update `tests/patterns.client.spec.ts`, `tests/editor.spec.tsx`, and `tests/file-type-icon.client.spec.tsx`.
- [x] T043 [US2] Add `.docm`/`.xlsm` preview, tool, and conversion surfaces where the extension list is hardcoded: `packages/client/ui-sidebar-documentpreview/src/client/{office,excel}/index.ts`, `src/client/document/unviewable.ts`, `src/client/excel/format.ts`, `packages/document/tool-office/src/{edit,read}.ts` diagnostics, and `packages/document/office-to-pdf/src/{types,index}.ts` (both formats are convertible). Update the corresponding tests.
- [x] T044 [US2] Record the constitution amendment 3.1.0 → 3.2.0 in `.specify/memory/constitution.md`: widen the v1 editable format sentence to `.docx/.docm/.xlsx/.xlsm`, record the package-level write layer over `node:zlib`, cite affected requirements (FR-022, SC-011, `001` FR-001/FR-008 supersession), and add the migration note (no data migration; old files remain readable). **Done (2026-11-18): Technology Constraints and the Sync Impact Report record the widened format set and the package-level write stack; version is 3.2.0.**

**Checkpoint**: US1 and US2 both pass independently; a macro-enabled or chart-bearing file saves when the edit is elsewhere.

---

## Phase 5: User Story 3 - Every save is attributable and auditable at entry level (Priority: P2)

**Goal**: Each save records a durable disposition whose counts and names exactly match the byte diff and survive restart.

**Independent Test**: Perform saves, compare the recorded disposition with an independent byte diff, and read the fact back after a restart.

**Dependency**: Requires T030 (writer returns a disposition) from US1 and T015 (event field) from Foundational.

### Tests for User Story 3

- [x] T045 [P] [US3] Add disposition tests in `packages/document/office-exchange/tests/remote.spec.ts` and `packages/session/session-document/tests/document-facts.spec.ts`: exact counts, sorted capped names, absent property tolerated, projection unchanged by the property, and a restart test that reopens the session log, folds the projection, and reads the disposition (SC-009).
- [x] T046 [P] [US3] Add a mismatch test in `packages/document/office-exchange/tests/ooxml-preservation.spec.ts` that compares the reported disposition against `comparePackages` and fails on any unaccounted entry.

### Implementation for User Story 3

- [x] T047 [US3] Compute `SaveDisposition` from the plan in `packages/document/office-exchange/src/package/ooxml-package.ts` and `src/provider/index.ts`: exact counts, sorted ASCII names capped at 256, and a bounded payload that never fails the save.
- [x] T048 [US3] Thread the disposition through `packages/document/office-exchange/src/remote.ts` `recordSave` into `packages/session/session-document/src/document-facts.ts` `recordDocumentSave`, and include it in the returned save result.
- [x] T049 [US3] Record the SDK decision in `specs/002-in-place-ooxml-save/quickstart-results.md`: no checked-in TypeScript or Python scenario emits `document/save`, so add no scenario and verify `packages/sdk/client/tests/expected/` and `python/sdk/tests/expected/` are unchanged. If a later scenario adds document events, update both projections in the same change.
- [x] T050 [US3] Document and test that the projection remains audit-only for the new property in `packages/session/session-document/src/document-projection.ts` and `packages/session/session-document/tests/`; if the projection exposes it, update the entry type, zod schema, `sameEntry`, `stateVersion`, and the compatibility record together.

**Checkpoint**: All three stories work together; the save fact is auditable after restart.

---

## Phase 6: Polish & Cross-Cutting Concerns

**Purpose**: Documentation, generated surfaces, persistence records, and the full validation ladder.

- [x] T051 [P] Update `packages/document/office-exchange/README.md` and `README.zh.md`: supported formats (`.docx/.docm/.xlsx/.xlsm`), source-package write behavior, preserved versus coupled constructs, disposition fact, and the revised Known Limitations. Update the pairing record through the bilingual workflow.
- [x] T052 [P] Update `docs/subsystems/office-exchange.md` and `.zh.md`, `docs/subsystems/office-editing.md` and `.zh.md`, and `docs/capability-seams.md` and `.zh.md` for the new write contract and refusal semantics.
- [x] T053 [P] Regenerate the Cordis catalog and API surface with `scripts/gen-cordis-catalog.ts`; update `packages/extensions/tool-cordis/src/api-catalog.ts` and the Cordis surface section of `docs/subsystems/office-exchange.md` and `.zh.md` from the generator, never by hand.
- [x] T054 [P] Update the superseded `001` artifacts: `specs/001-office-agent-framework/contracts/office-exchange.contract.md` (write signature), `specs/001-office-agent-framework/spec.md` (FR-008 note pointing at `002`), and `specs/001-office-agent-framework/plan.md` (format set).
- [x] T055 [P] Update `packages/document/office-exchange/tests/fixtures/README.md` (and its pair) for the new corpus entries and the `--check` regeneration procedure.
- [x] T056 Write the durable design note `.agents/notes/implemented/architecture/2026-11-18-in-place-ooxml-save.md`: why raw ZIP record copy, why the custom scanner over `captureMetaData`, why anchors live in `data.officeExchange.source`, and the coupled/independent classification.
- [x] T057 Run the `quickstart.md` scenarios S1–S11 and record commands, hosts, and outcomes in `specs/002-in-place-ooxml-save/quickstart-results.md`. **Done: S1–S8, S10, S11 ran; S9 (manual Word/Excel) is tracked by T061 and has not been performed.**
- [x] T058 Measure the performance target in `packages/document/office-exchange/tests/ooxml-preservation.spec.ts` (5 MB / 10,000 cells, save under 30 s) and record the measurement in `specs/002-in-place-ooxml-save/quickstart-results.md`; if it misses, profile `src/package/zip-writer.ts` and the span index before changing the target.
- [x] T059 Run focused coverage and close gaps: `pnpm exec vitest run packages/document/office-exchange/tests/{ooxml-preservation,ooxml-engine,write-postcondition,remote,zip-reader,zip-writer,ooxml-package,xml-spans,unit-digest,source-state}.spec.ts --coverage --coverage.include='packages/document/office-exchange/src/**/*.ts'`; then `pnpm run test:coverage` before landing. **Done for every modified/new source file (100% statements/branches/functions/lines). Only the unmodified `src/commands.ts:168` pre-existing gap remains; the repo-wide `test:coverage` gate was not run.**
- [x] T060 Run `pnpm run typecheck`, `pnpm run lint`, `pnpm run build:artifacts`, `pnpm run build:check`, `pnpm run hygiene`, and `pnpm run doc-sync`; record results in `specs/002-in-place-ooxml-save/quickstart-results.md`. **Done: typecheck, build:artifacts, build:check, hygiene, and duplication pass; lint and doc-sync fail only on pre-existing unmodified files plus the environmental CLI-help SIGKILL, all recorded with evidence.**
- [ ] T061 Run the manual Word/Excel release-subset matrix on Windows and macOS and record the builds, hosts, and outcomes in `specs/002-in-place-ooxml-save/quickstart-results.md`.
- [ ] T062 Final consistency pass: run `pnpm run duplication`, confirm every task above is checked or explicitly deferred with a reason, and write the PR description with `dsh-pr-description`; record the final task status and evidence summary in `specs/002-in-place-ooxml-save/quickstart-results.md`.

---

## Dependencies & Execution Order

### Phase Dependencies

- **Setup (Phase 1)**: T001 records the baseline; T002–T005 touch different files and run in parallel with it.
- **Foundational (Phase 2)**: Depends on Setup. Blocks every user story.
  - T006 blocks T007, T008, T018, T030.
  - T009 blocks T018, T025, T027, T028.
  - T010 blocks T018, T028.
  - T011 and T012 block T018.
  - T013 blocks T014, T030.
  - T014 blocks T030, T032.
  - T015 blocks T048.
  - T016 blocks T042, T043.
  - T017 blocks T029, T038.
  - T018 and T019 block T028, T030.
  - T020 blocks T037, T043.
- **User Stories (Phase 3+)**: All start after Foundational.
  - US1 (Phase 3) is the MVP.
  - US2 (Phase 4) depends on T029 and T030 from US1.
  - US3 (Phase 5) depends on T030 and T015.
- **Polish (Phase 6)**: Depends on all three stories.

### Within Each User Story

- Tests first (T021–T024, T035–T037, T045–T046), then implementation.
- `src/source/splice.ts` before the provider writers.
- Provider writers before the Remote wiring and the Web scenario.
- Refusal classification before the gate removal.
- Disposition computation before the fact plumbing.

### Parallel Opportunities

- Setup: T002, T003, T004, T005 touch different files and do not depend on T001's output.
- Foundational: T006–T012 and T017 touch different modules; T013–T016 touch the seam and session types, so keep T013 → T014 ordered; T015 and T016 touch different event types but both regenerate persistence docs, so run them sequentially.
- US1 tests: T021, T022, and T023 are independent files; T024 extends the `ooxml-preservation.spec.ts` file created by T021, so run it after T021.
- US1 implementation: T025 (docx.ts) and T027 (xlsx.ts) are independent; T026 shares `docx.ts` with T025, so run T026 after T025 or coordinate.
- US2 tests: T035–T037 are independent files.
- US2 implementation: T039 (tool-office), T041 (panel/locales), and T042 (client definition/icons) touch different files and can run in parallel after T038. T040 and T042 both edit `useOfficeDocument.ts`, so run T040 after T042 or coordinate the change.
- US3: T047 → T048 → T049; T045/T046 run in parallel with implementation.

---

## Parallel Example: User Story 1

```text
Task: "Add the corpus identity matrix in packages/document/office-exchange/tests/ooxml-preservation.spec.ts"
Task: "Add DOCX splice tests in packages/document/office-exchange/tests/docx-splice.spec.ts"
Task: "Add XLSX splice tests in packages/document/office-exchange/tests/xlsx-splice.spec.ts"
```

T024 then extends the T021 file. After the tests fail, T025 and T027 can proceed in parallel; T026 follows T025; T028 precedes T030.

## Implementation Strategy

### MVP First (User Story 1 only)

1. Complete Phase 1 and Phase 2.
2. Complete Phase 3 (US1).
3. **STOP and VALIDATE**: run the US1 independent test and the LibreOffice smoke.
4. Demo an edited `.docx`/`.xlsx` whose untouched bytes are unchanged.

### Incremental Delivery

1. Setup + Foundational → foundation ready.
2. US1 → byte-preserving save for ordinary files (MVP).
3. US2 → complex and macro-enabled files save; refusals become edit-scoped.
4. US3 → durable, auditable save dispositions.
5. Polish → docs, generated surfaces, amendment records, full evidence.

### Parallel Team Strategy

1. One team completes Setup + Foundational together.
2. After Foundational: Developer A takes US1, Developer B prepares US2 tests and the client surfaces, Developer C prepares US3 tests and the persistence records.
3. US2 and US3 finish their implementation once US1's plan and writer land.

---

## Notes

- [P] tasks touch different files and have no incomplete dependency.
- A story is complete only when its independent test passes against the real provider, not a stub.
- Do not claim the XLSX path preserves byte identity until T027 and T030 land; until then it is semantic-only.
- The legacy writer remains for source-less snapshots; do not delete it in this feature.
- Record only commands actually run; the coverage gate is `pnpm run test:coverage`, not `pnpm run test`.

---

## Phase 7: Convergence

- [ ] T063 Run the manual desktop Word and Excel no-repair matrix on Windows and macOS for the fixed release subset and record the builds, hosts, and outcomes in `specs/002-in-place-ooxml-save/quickstart-results.md` per SC-003/FR-012/T061 (missing)
- [ ] T064 Run the repository-wide coverage gate `pnpm run test:coverage`, record the result in `specs/002-in-place-ooxml-save/quickstart-results.md`, and classify any pre-existing gap separately per plan validation ladder/T059 (partial). **Attempted 2026-11-19: the 4-partition runner started all partitions but exceeded the 30-minute budget and produced no merged report; focused per-file coverage for the feature remains 100%. Run in CI or a longer window.**
- [x] T065 Re-run `pnpm run build:artifacts`, `pnpm run build:check`, and `pnpm run hygiene` on the final review-fixed tree and record the results per T060/plan validation ladder (partial). **Done 2026-11-19: typecheck, build:check, and duplication pass; `build:artifacts`' Desktop bundle and `hygiene`'s publint are blocked by environmental vendor-link/missing-lib issues with evidence recorded.**
- [ ] T066 Open the pull request and write its description with `dsh-pr-description`, including the constitution amendment, persistence acknowledgements, and the evidence actually run, per T062 (missing)
- [ ] T067 Review the out-of-scope `packages/document/tool-office/src/edit-cells.ts` de-duplication added for the duplication gate; either record why it belongs in this change or split it into a separate change, per plan scope/T062 (unrequested)
- [x] T068 Refuse a worksheet row or column shift when a referenced table part relationship cannot be resolved or the table has no `ref`, or document the malformed-package behavior in `packages/document/office-exchange/README.md`, per FR-011 (partial). **Done 2026-11-19: the shift now refuses with the table relationship/part named; 1,156 office tests pass and `xlsx.ts` stays at 100% coverage.**
