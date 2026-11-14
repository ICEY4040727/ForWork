---

description: "Task list for the Office Agent Framework feature"
---

# Tasks: Office Agent Framework

**Input**: Design documents from `/specs/001-office-agent-framework/`
**Prerequisites**: plan.md, spec.md, research.md, data-model.md, contracts/, quickstart.md

**Tests**: Included — the repository mandates behavior tests and a per-file coverage gate.

**Organization**: Tasks are grouped by user story so each increment is independently testable.

## Format: `[ID] [P?] [Story] Description`

- **[P]**: Can run in parallel (different files, no dependencies)
- **[Story]**: US1, US2, US3 for story phases only
- Every task names its exact file path

## Path Conventions

Monorepo: workspace packages under `packages/<group>/<pkg>/`, desktop under `apps/desktop/`, fixtures under the owning package's `tests/fixtures/`.

---

## Phase 1: Setup (Shared Infrastructure)

**Purpose**: Toolchain and package skeletons.

- [x] T001 Confirm toolchain: activate `pnpm@11.28.5` and install; evidence: `corepack pnpm install --frozen-lockfile` succeeded.
- [x] T002 [P] Create the Exchange seam package in `packages/document/office-exchange/` (`package.json`, `tsconfig.json`, `src/index.ts`, `src/types.ts`, `src/errors.ts`, `src/format.ts`, `tests/format.spec.ts`).
- [x] T003 [P] Register the Exchange package in `tsconfig.host.json`, `tsconfig.base.json` paths, and `packages/document/README.md`.
- [x] T004 [P] Create the agent tool package skeleton in `packages/document/tool-office/` (`package.json`, `tsconfig.json`, `src/index.ts`, `tests/`).
- [x] T005 [P] Create the editor package skeleton in `packages/client/ui-office-editor/` with Client and Host faces (`tsconfig.json`, `tsconfig.host.json`), mirroring `packages/client/ui-sidebar-browser/`.
- [x] T006 [P] Add Univer Apache-2.0 presets to `packages/client/ui-office-editor/package.json` and `@univerjs/core` to `packages/document/office-exchange/package.json`; record the license check in `packages/client/ui-office-editor/README.md` and `packages/document/office-exchange/README.md`.
- [x] T007 [P] Add round-trip fixtures (`.xlsx` with formulas/styles, `.docx` with headings/tables/images, one legacy `.doc`/`.xls`, one macro-bearing file) under `packages/document/office-exchange/tests/fixtures/`.

---

## Phase 2: Foundational (Blocking Prerequisites)

**Purpose**: The provider behind the seam, shared session/identity/write/privacy infrastructure. Blocks all stories.

**⚠️ CRITICAL**: No user story starts until this phase is complete.

- [x] T008 Implement the `OfficeExchange` provider's `.xlsx` mapping in `packages/document/office-exchange/src/provider/xlsx.ts` between OOXML and Univer `IWorkbookData` (Option B) with ExcelJS/SheetJS and untouched-region preservation; tests in `packages/document/office-exchange/tests/provider-xlsx.spec.ts`.
- [x] T009 [P] Implement the `.docx` run-preserving mapping in `packages/document/office-exchange/src/provider/docx.ts` between OOXML and Univer `IDocumentData` (Option B); tests in `packages/document/office-exchange/tests/provider-docx.spec.ts`.
- [x] T010 [P] Implement unsupported-feature detection and refusal in `packages/document/office-exchange/src/provider/detect.ts` (macros, tracked changes, complex shapes, v1 headers/footers); tests in `packages/document/office-exchange/tests/detect.spec.ts`.
- [x] T011 [P] Implement legacy conversion in `packages/document/office-exchange/src/provider/legacy.ts` using the LibreOffice kit; first verify target-format round-trip in `packages/document/office-exchange/tests/legacy.spec.ts`.
- [x] T012 [P] Add bounded-size/time configuration with loud validation in `packages/document/office-exchange/src/config.ts`; tests in `packages/document/office-exchange/tests/config.spec.ts`.
- [x] T013 Register the provider row in `packages/bundle/base/cordis.patch.yml` and its dependency in `packages/bundle/base/package.json`.
- [x] T014 Define `document/*` events in `packages/core/session/src/` and record the `SESSION_FORMAT_VERSION` decision per the session-format policy; tests in `packages/core/session/tests/document-events.spec.ts`.
- [x] T015 [P] Implement the stable identity registry with rebind/missing in `packages/session/session-document/src/document-identity.ts`; tests in `packages/session/session-document/tests/document-identity.spec.ts`.
- [x] T016 [P] Implement Change Record, attribution, and the version guard in `packages/session/session-document/src/document-changes.ts`; tests in `packages/session/session-document/tests/document-changes.spec.ts`.
- [x] T017 [P] Implement write-root confinement in `packages/document/office-exchange/src/authority.ts` using the filesystem seam; tests in `packages/document/office-exchange/tests/authority.spec.ts`.
- [x] T018 [P] Wire credential storage to the OS secure store in `packages/credentials/credentials/` (consumer only, no new subsystem); tests in `packages/credentials/credentials/tests/secure-store.spec.ts`.
- [x] T019 [P] Enforce telemetry-off-by-default and add the egress assertion harness in `packages/test-support/tool-testkit/`; test in `packages/test-support/tool-testkit/tests/egress.spec.ts`.
- [x] T020 Verify the foundation: `corepack pnpm exec tsc -b packages/document/office-exchange packages/core/session packages/session/session-document` plus `corepack pnpm exec vitest run` for those packages.

**Checkpoint**: Provider, events, identity, changes, write authority, credentials, and privacy harness are ready.

> Technical selections (2026-11-14, recorded in spec/plan/research/data-model/contracts): agent tools stay Host-owned and send correlated gateway commands to the single live unit (R15); runtime undo is the editor's native stack (R16); explicit save plus periodic drafts (R17); adapter-mock component tests plus a browser smoke (R18); additive operation union (R19); legacy files open as a converted copy in a new tab (R20).
>
> Consistency pass (2026-11-14): save is the only version bump; an accepted agent tool call persists immediately and refuses mixed-actor saves (`pending-edits`); save-time unsupported refusal added as T074; T025/T065 merged; T056/T057 split into audit helper vs client native undo; T040/T070 split into Host draft storage vs client capture; `Document Draft` added to the spec entities; "same file" clarified for legacy sources; SC-003 timing moved to the client suite.
>
> Progress (2026-11-14): Phase 2 complete plus US1 host prerequisites: T064 `officeDocument` Remote (version-guarded save), T066 fs `writeBytes` across local/sandbox/ssh, T067 Cordis catalog classification + subsystem page, T068 workspace-scoped legacy conversion. Verified: 343 tests across fs-local/fs-sandbox/office-exchange (office-exchange 105), 7-project tsc, oxlint, frozen install, cordis/module/config catalogs, product-use, JSDoc. **Remaining US1**: T021–T028, T052–T055, T065, T070, T072, T074 (client editor + tests).

---

## Phase 3: User Story 1 - Edit `.xlsx`/`.docx` in-app (Priority: P1) 🎯 MVP

**Goal**: Open a local file in the sidebar, edit, and save losslessly.

**Independent Test**: Fixture round-trip (S1/S2 in quickstart) plus the GUI test.

### Tests for User Story 1

- [ ] T021 [P] [US1] Editor contract test for tab registration and open/save in `packages/client/ui-office-editor/tests/editor.spec.tsx`.

### Implementation for User Story 1

- [ ] T022 [US1] Register the `office` tab kind and file patterns in `packages/client/ui-office-editor/src/client/index.ts`.
- [ ] T023 [P] [US1] Embed the Univer Sheets preset in `packages/client/ui-office-editor/src/client/SheetsEditor.tsx`.
- [ ] T024 [P] [US1] Embed the Univer Docs preset in `packages/client/ui-office-editor/src/client/DocsEditor.tsx`.
- [ ] T025 [US1] Implement `useOfficeDocument` against the `officeDocument` Remote (read/save/detect/convert) in `packages/client/ui-office-editor/src/client/useOfficeDocument.ts`.
- [x] T064 [US1] Expose a session-scoped `officeDocument` Client Remote (`read`, `save`, `detectUnsupported`, `convertLegacy`) in `packages/document/office-exchange/src/remote.ts`, persisting saves through the filesystem seam with a version guard; tests in `packages/document/office-exchange/tests/remote.spec.ts`. **Blocked on T066.**
- [x] T066 [US1] Add a version-guarded binary write to the filesystem seam and all providers: `packages/fs/fs/src/` (Service Definition), `packages/fs/fs-local/`, `packages/fs/fs-sandbox/`, `packages/fs/fs-ssh/`, reusing `FsWriteIntent` and the existing atomic-write/lock path; tests per provider. Required by T064 because OOXML is binary.
- [ ] T065 [US1] Verify the client imports no Host service directly: assert only the `officeDocument` Remote is used in `packages/client/ui-office-editor/src/client/useOfficeDocument.ts`; test in `packages/client/ui-office-editor/tests/remote-only.spec.tsx`.
- [ ] T026 [US1] Handle external change reload and unsaved-close warning in `packages/client/ui-office-editor/src/client/changes.ts`.
- [ ] T070 [US1] Implement client-side draft capture (idle-debounced, version-bound) and its `officeDocument` Remote call in `packages/client/ui-office-editor/src/client/drafts.ts`; tests in `packages/client/ui-office-editor/tests/drafts.spec.ts`.
- [ ] T074 [US1] Refuse a save when the open document carries unsupported features: re-check the read result/detection in the save hook and show the named feature in `packages/client/ui-office-editor/src/client/useOfficeDocument.ts`; test in `packages/client/ui-office-editor/tests/save-refusal.spec.tsx`.
- [ ] T027 [US1] Render the open-document panel in `packages/client/ui-office-editor/src/client/DocumentPanel.tsx`.
- [ ] T028 [US1] Verify: `corepack pnpm run test:gui`; `DSH_SNAPSHOT=replay corepack pnpm run test:web` if assembled output changed.
- [ ] T052 [US1] Exclude `.pptx` from the editable tab and assert it stays preview-only (FR-011) in `packages/client/ui-office-editor/tests/patterns.spec.ts`
- [ ] T053 [US1] Wire legacy conversion on open by calling `convertLegacy` in `packages/client/ui-office-editor/src/client/useOfficeDocument.ts`; test in `packages/client/ui-office-editor/tests/legacy-open.spec.tsx`
- [ ] T073 [US1] Open legacy conversions in a new tab with a visible "editing converted copy" notice in `packages/client/ui-office-editor/src/client/DocumentPanel.tsx`; test in `packages/client/ui-office-editor/tests/legacy-notice.spec.tsx`
- [ ] T072 [US1] Isolate Univer behind an adapter seam and mock it in component tests; add the browser smoke for real editor behavior in `packages/client/ui-office-editor/src/client/adapter.ts` and `packages/client/ui-office-editor/tests/adapter.spec.ts`
- [x] T067 [US1] Classify `officeExchange`/`officeDocument` in the Cordis catalog (`SERVICE_PAGE`, `LINK_MAP`, and a new `docs/subsystems/office-exchange.md` + `.zh.md` `cordis-surface` page) so `gen-cordis-catalog --check` passes.
- [x] T068 [US1] Make `convertLegacy` write into a workspace-scoped temp directory inside the permitted roots so the converted file is readable through the Remote; update the provider, Remote, and tests.
- [ ] T054 [US1] Bind undo/redo to the editor's standard undo surface in `packages/client/ui-office-editor/src/client/undo.ts`
- [ ] T055 [P] [US1] Add a timed open/save test for SC-001 (5 MB / 10,000 cells fixture) in `packages/client/ui-office-editor/tests/performance.spec.tsx`

**Checkpoint**: MVP is demonstrable — open, edit, save works independently.

---

## Phase 4: User Story 2 - Agent operates the open document (Priority: P1)

**Goal**: Agent edits converge with the editor on one attributed, versioned state.

**Independent Test**: Quickstart S3–S5 and S8.

- [ ] T029 [P] [US2] Register `document_read`, `document_edit_cells`, `document_edit_paragraph`, `document_convert` with model-visible schemas in `packages/document/tool-office/src/index.ts`.
- [ ] T071 [US2] Extend the `document/edit` operation union with row/column/formula/style kinds and record the same-version persistence acknowledgement in `packages/core/session/src/document-events.ts` and `docs/persistence-changes/`
- [ ] T030 [US2] Implement `document_read` in `packages/document/tool-office/src/read.ts` per `contracts/tool-office.contract.md`.
- [ ] T031 [US2] Implement `document_edit_cells` in `packages/document/tool-office/src/edit-cells.ts`: apply to the live unit, persist with one immediate version-guarded save, and reject stale versions and `pending-edits`.
- [ ] T032 [P] [US2] Implement `document_edit_paragraph` in `packages/document/tool-office/src/edit-paragraph.ts` with the same live-unit apply, immediate save, and `pending-edits` rule.
- [ ] T033 [P] [US2] Implement `document_convert` in `packages/document/tool-office/src/convert.ts`.
- [ ] T034 [US2] Return structured `conflict`, `unsupported-feature`, `out-of-scope`, `document-missing`, `document-not-open`, `pending-edits`, `bound-exceeded` failures in `packages/document/tool-office/src/errors.ts`.
- [ ] T035 [US2] Verify: session snapshots for tool output plus `corepack pnpm run test:coverage` for the tool package.
- [ ] T062 [US2] Implement the Host→renderer command bridge: the Host tool sends a correlated gateway command (local IPC on desktop, no document bytes off-device), the renderer applies it to the live unit and returns the correlated result, with timeouts and `document-not-open` in `packages/client/ui-office-editor/src/client/command-bridge.ts`; test in `packages/client/ui-office-editor/tests/command-bridge.spec.tsx`.
- [ ] T056 [US2] Expose `ChangeRecord` inverse/audit helpers only; runtime undo/redo is the editor's native stack bound in T054, in `packages/session/session-document/src/undo.ts`
- [ ] T057 [P] [US2] Test undoing an agent edit through the native undo stack (SC-012) in `packages/client/ui-office-editor/tests/undo.spec.tsx`
- [ ] T058 [P] [US2] Add a timed agent-edit visibility test for SC-003 in `packages/client/ui-office-editor/tests/performance.spec.tsx`

**Checkpoint**: US1 and US2 work; agent and user converge and conflicts are visible.

---

## Phase 5: User Story 3 - Documents persist as session entities (Priority: P2)

**Goal**: Open documents survive restart and are referenceable across turns.

**Independent Test**: Quickstart S6, S7 and the restart scenario.

- [ ] T036 [US3] Emit `document/open|edit|close|rebind|missing` facts from `packages/core/session/src/` per `contracts/document-events.contract.md`.
- [ ] T037 [US3] Implement the `document` projection and open-document system-prompt injection in `packages/session/session-document/src/document-projection.ts`; tests in `packages/session/session-document/tests/document-projection.spec.ts`.
- [ ] T038 [US3] Implement restore-after-restart of the open-document set in `packages/session/session-document/src/document-restore.ts`.
- [ ] T039 [P] [US3] Implement rename/move rebind and missing reporting in `packages/session/session-document/src/document-identity.ts`.
- [ ] T040 [P] [US3] Implement Host-side draft persistence and recovery storage in `packages/session/session-document/src/document-draft.ts` (location, version binding, cleanup); the client capture/debounce is T070.
- [ ] T041 [US3] Verify: session migration test, `corepack pnpm run test:snapshot`, and the TypeScript and Python SDK expected-output updates.

**Checkpoint**: Documents are durable, restorable, and cross-turn referenceable.

---

## Phase 6: Polish & Cross-Cutting Concerns

**Purpose**: Tests for the cross-cutting guarantees, docs, and the final independent-identity release.

- [ ] T042 [P] Add the privacy egress test proving no document bytes leave the device (SC-008) in `packages/test-support/tool-testkit/tests/egress-documents.spec.ts`.
- [ ] T043 [P] Add the write-authority test proving out-of-root writes are refused (SC-011) in `packages/document/tool-office/tests/authority.spec.ts`.
- [ ] T044 [P] Add the unsaved-work recovery test (SC-010) in `packages/session/session-document/tests/document-draft.spec.ts`.
- [ ] T045 [P] Document the packages and subsystem: `packages/document/office-exchange/README.md` Model Experience section, `packages/client/ui-office-editor/README.md`, `docs/subsystems/office-editing.md`.
- [ ] T046 [P] Add config catalog entries and run `verify-package-dependencies`, `verify-export-jsdoc`, and `verify-client-packages` for the new packages.
- [ ] T047 Run `corepack pnpm run build:check`, `corepack pnpm run doc-sync`, and `corepack pnpm run test:coverage`; fix regressions.
- [ ] T048 Run quickstart scenarios S1–S11 and record the results.
- [ ] T059 Add the upstream merge procedure (clean tree, verified recovery copy, no shared stash, recorded conflict intent) to `docs/development.md` and reference it from `specs/001-office-agent-framework/quickstart.md`
- [ ] T060 Assert MIT and Apache-2.0 notices in the packaged artifact (FR-012) in `scripts/verify-license-notices.ts`
- [ ] T061 Add an offline test proving edit and save work with no document server configured (FR-010) in `packages/document/office-exchange/tests/offline.spec.ts`
- [ ] T049 Brand/telemetry scan gate asserting zero DeepSeek-referencing shipped surfaces (FR-009, SC-006).
- [ ] T050 Rename package scope and CLI, replace brand slots/artwork/docs, remove telemetry, and replace the default model adapter (final, gated on everything above).
- [ ] T051 Configure Windows and macOS signing identities and produce signed installers (SC-007); verify both launch.
- [ ] T063 [P] Construct `SafeStorageSecureStore` in the desktop host with Electron `safeStorage` and `app.getPath('userData')`, and connect it to the model-credential path; surface `secure-storage-unavailable` in `apps/desktop/` and `packages/credentials/credentials/src/`.
- [ ] T069 [P] Classify the `officeExchange`/`officeDocument` product roles so `gen-doc-graphs --check` regenerates `docs/capability-seams.md`; update the service-role classification and re-run the doc gates.

---

## Dependencies & Execution Order

### Phase Dependencies

- **Setup (Phase 1)**: T001–T003 done; T004–T007 can start now.
- **Foundational (Phase 2)**: T008–T020 block every user story.
  - T008 (xlsx) and T009 (docx) are independent file sets and may run in parallel.
  - T014 (events) blocks T036–T041.
  - T015–T017 block T031–T034.
  - T016 (Change Record inverses) blocks T056 and T057.
  - T062 (renderer command bridge) blocks T031 and T056 for open-document edits.
- **US1 (Phase 3)**: needs T008, T009, T013, T014.
- **US2 (Phase 4)**: needs T008–T012, T016, T017; shares the provider with US1 but is independently testable.
- **US3 (Phase 5)**: needs T014–T016 and US1's editor.
- **Polish (Phase 6)**: needs all desired stories; T050 is the last destructive task (later non-destructive verification tasks may follow it).
- T059 relates to FR-019 and Constitution VIII.

### User Story Dependencies

- **US1 (P1)**: no dependency on US2/US3.
- **US2 (P1)**: no dependency on US1 (headless tools operate the same provider); integrating with the open editor is a verification step.
- **US3 (P2)**: depends on US1 for the open editor and on T014 for the event vocabulary.

### Parallel Opportunities

- Setup: T004, T005, T006, T007 are independent.
- Foundational: T009, T010, T011, T012 are separate files; T015–T019 touch different packages.
- US1: T023 and T024 are separate components.
- US2: T032 and T033 are separate files.
- US3: T039 and T040 are separate files.

### Suggested MVP

Complete Phase 1 + Phase 2 + Phase 3 (US1). Stop and validate S1/S2; that is a usable editor.

---

## Notes

- [P] tasks touch different files and can run in parallel.
- Tests are required by the repository's coverage gate; write them with each behavior.
- The destructive rename (T050) is intentionally last (Constitution V).
- Every task declares its own verification; a later task does not assume an earlier check passed.
