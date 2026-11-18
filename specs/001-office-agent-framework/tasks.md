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
- [x] T018 [P] Wire the package-level credential adapter to the OS secure store in `packages/credentials/credentials/` (consumer only, no new subsystem); tests in `packages/credentials/credentials/tests/secure-store.spec.ts`. Desktop `safeStorage` construction/wiring is T063.
- [x] T019 [P] Enforce telemetry-off-by-default and add the egress assertion harness in `packages/test-support/tool-testkit/`; test in `packages/test-support/tool-testkit/tests/egress.spec.ts`.
- [x] T020 Verify the foundation: `corepack pnpm exec tsc -b packages/document/office-exchange packages/core/session packages/session/session-document` plus `corepack pnpm exec vitest run` for those packages.

**Checkpoint**: Provider, events, identity, changes, write authority, credentials, and privacy harness are ready.

> Technical selections (2026-11-14, recorded in spec/plan/research/data-model/contracts): agent tools stay Host-owned and send correlated gateway commands to the single live unit (R15); runtime undo is the editor's native stack (R16); explicit save plus periodic drafts (R17); adapter-mock component tests plus a browser smoke (R18); additive operation union (R19); legacy files open as a converted copy in a new tab (R20).
>
> Consistency pass (2026-11-14): save is the only version bump; an accepted agent tool call persists immediately and refuses mixed-actor saves (`pending-edits`); save-time unsupported refusal added as T074; T025/T065 merged; T056/T057 split into audit helper vs client native undo; **T077 owns the Host draft store, T040 orchestrates recovery, T070 captures in the client**; `Document Draft` added to the spec entities; "same file" clarified for legacy sources; SC-003 timing moved to the client suite.
>
> Progress (2026-11-15): **US3 implemented and committed** — facts (`open`/`save`/`edit`/`close`/`rebind`/`missing`), `document` projection + dynamic prompt section, restore + draft recovery, Host lifecycle endpoints. Verified: session-document 40, client 175, office-exchange 141, tool-office 80; aggregates 0; migration/catalog 1448 pass. **T041/T035 remain:** snapshot replay needs re-recording with a key (5 office tool-list deltas; Windows env diffs), and no `document_*` session fixture exists. **Lockfile refreshed separately.**
>
> Identity landing (2026-11-16): packages and vendor re-scoped to `@forwork/*`, CLI/brand/scan landed; `verify-forwork-identity` passes and T049/T050 are closed. Residual: repository URL, protocol tokens `dsh.*`/`DSH_*`, external `libreoffice-kit` (T089–T091), signing (T051); post-rename gates (T092–T094).

---

## Phase 3: User Story 1 - Edit `.xlsx`/`.docx` in-app (Priority: P1) 🎯 MVP

**Goal**: Open a local file in the sidebar, edit, and save losslessly.

**Independent Test**: Fixture round-trip (S1/S2 in quickstart) plus the GUI test.

### Tests for User Story 1

- [x] T021 [P] [US1] Editor contract test for tab registration and open/save in `packages/client/ui-office-editor/tests/editor.spec.tsx`.

### Implementation for User Story 1

- [x] T022 [US1] Register the `office` tab kind and file patterns in `packages/client/ui-office-editor/src/client/index.ts`.
- [x] T075 [US1] Register `ui-office-editor` in the web-app bundle: add the `dsh.client` row in `packages/bundle/web-app/cordis.patch.yml` and the dependency in `packages/bundle/web-app/package.json` (its `tsconfig.client.json` reference already exists).
- [x] T023 [P] [US1] Embed the Univer Sheets preset in `packages/client/ui-office-editor/src/client/editor/sheets.ts`.
- [x] T024 [P] [US1] Embed the Univer Docs preset in `packages/client/ui-office-editor/src/client/editor/docs.ts`.
- [x] T025 [US1] Implement `useOfficeDocument` against the `officeDocument` Remote (read/save/detect/convert) in `packages/client/ui-office-editor/src/client/useOfficeDocument.ts`.
- [x] T064 [US1] Expose a session-scoped `officeDocument` Client Remote (`read`, `save`, `detectUnsupported`, `convertLegacy`) in `packages/document/office-exchange/src/remote.ts`, persisting saves through the filesystem seam with a version guard; tests in `packages/document/office-exchange/tests/remote.spec.ts`. T066 supplies the binary write.
- [x] T066 [US1] Add a version-guarded binary write to the filesystem seam and all providers: `packages/fs/fs/src/` (Service Definition), `packages/fs/fs-local/`, `packages/fs/fs-sandbox/`, `packages/fs/fs-ssh/`, reusing `FsWriteIntent` and the existing atomic-write/lock path; tests per provider. Required by T064 because OOXML is binary.
- [x] T065 [US1] Verify the client imports no Host service directly: assert only the `officeDocument` Remote is used in `packages/client/ui-office-editor/src/client/useOfficeDocument.ts`; test in `packages/client/ui-office-editor/tests/remote-only.spec.ts`.
- [x] T026 [US1] Handle external change reload and unsaved-close warning in `packages/client/ui-office-editor/src/client/changes.ts`.
- [x] T070 [US1] Implement client-side draft capture (idle-debounced, version-bound) and its `officeDocument` Remote call in `packages/client/ui-office-editor/src/client/drafts.ts`; tests in `packages/client/ui-office-editor/tests/drafts.client.spec.ts`.
- [x] T077 [US1] Add `writeDraft`/`readDraft`/`clearDraft` to the `officeDocument` Remote and Host storage under `<workspaceRoot>/.dsh/office-tmp/drafts/` (version-bound, authority-checked); tests in `packages/document/office-exchange/tests/drafts.spec.ts`.
- [x] T074 [US1] Refuse a save when the open document carries unsupported features: re-check the read result/detection in the save hook and show the named feature in `packages/client/ui-office-editor/src/client/useOfficeDocument.ts`; test in `packages/client/ui-office-editor/tests/use-office-document.spec.tsx` and `tests/editor.spec.tsx`.
- [x] T027 [US1] Render the open-document panel in `packages/client/ui-office-editor/src/client/DocumentPanel.tsx`.
- [ ] T028 [US1] Verify: `corepack pnpm run test:gui`; then `DSH_SNAPSHOT=replay corepack pnpm run test:web` — mandatory because US1 changed assembled output. **Office-caused failures resolved:** `test:gui` green after rebrand-drift test fixes (761/763, 2 residual load-timeout flakes pass in isolation); `build:artifacts` 8/8; focused keyless web replay of the two office scenarios passes — `document-preview` 6/7 (1 pre-existing Windows CRLF clipboard) and `excel-opc` 1/1 (routed through the read-only fallback). **Full 237-file lane not run:** multi-hour serial on Windows; remaining known failures are pre-existing Windows env diffs (~36) and the shipped tool roster (T041, needs a key).
- [x] T078 Fix the host-aggregate `TS2375` in `packages/document/office-exchange/tests/provider-docx.spec.ts` and triage the remaining `tsc -p tsconfig.host.json` errors as pre-existing vs ours; rerun the host aggregate so T028 can attempt `test:web`. **Host aggregate 177→0 after moving client tests to the `.client.` lane; web lane unblocked.**
- [x] T052 [US1] Exclude `.pptx` from the editable tab and assert it stays preview-only (FR-011) in `packages/client/ui-office-editor/tests/patterns.client.spec.ts`
- [x] T053 [US1] Wire legacy conversion on open by calling `convertLegacy` in `packages/client/ui-office-editor/src/client/useOfficeDocument.ts`; test in `packages/client/ui-office-editor/tests/legacy-open.spec.tsx`
- [x] T072 [US1] Isolate Univer behind an adapter seam and mock it in component tests in `packages/client/ui-office-editor/src/client/adapter.ts` and `packages/client/ui-office-editor/tests/adapter.spec.ts`; the real browser smoke is T028.
- [x] T073 [US1] Open legacy conversions in a new tab with a visible "editing converted copy" notice in `packages/client/ui-office-editor/src/client/DocumentPanel.tsx`; test in `packages/client/ui-office-editor/tests/legacy-notice.spec.tsx`
- [x] T067 [US1] Classify `officeExchange`/`officeDocument` in the Cordis catalog (`SERVICE_PAGE`, `LINK_MAP`, and a new `docs/subsystems/office-exchange.md` + `.zh.md` `cordis-surface` page) so `gen-cordis-catalog --check` passes.
- [x] T068 [US1] Make `convertLegacy` write into a workspace-scoped temp directory inside the permitted roots so the converted file is readable through the Remote; update the provider, Remote, and tests.
- [x] T054 [US1] Bind undo/redo to the editor's standard undo surface in `packages/client/ui-office-editor/src/client/undo.ts`
- [x] T055 [P] [US1] Add a timed open/save test for SC-001 (5 MB / 10,000 cells fixture) in `packages/client/ui-office-editor/tests/performance.spec.tsx`

**Checkpoint**: MVP is demonstrable — open, edit, save works independently.

---

## Phase 4: User Story 2 - Agent operates the open document (Priority: P1)

**Goal**: Agent edits converge with the editor on one attributed, versioned state.

**Independent Test**: Quickstart S3–S5 and S8.

- [x] T029 [P] [US2] Register `document_read`, `document_edit_cells`, `document_edit_paragraph`, `document_convert` with model-visible schemas in `packages/document/tool-office/src/index.ts`.
- [x] T071 [US2] Extend the `document/edit` operation union with row/column/formula/style kinds and record the same-version persistence acknowledgement in `packages/core/session/src/document-events.ts` and `docs/persistence-changes/`
- [x] T030 [US2] Implement `document_read` in `packages/document/tool-office/src/read.ts` per `contracts/tool-office.contract.md`.
- [x] T031 [US2] Implement `document_edit_cells` in `packages/document/tool-office/src/edit-cells.ts`: apply to the live unit, persist with one immediate version-guarded save, and reject stale versions and `pending-edits`.
- [x] T032 [P] [US2] Implement `document_edit_paragraph` in `packages/document/tool-office/src/edit-paragraph.ts` with the same live-unit apply, immediate save, and `pending-edits` rule.
- [x] T033 [P] [US2] Implement `document_convert` in `packages/document/tool-office/src/convert.ts`.
- [x] T034 [US2] Return structured `conflict`, `unsupported-feature`, `out-of-scope`, `document-missing`, `document-not-open`, `pending-edits`, `bound-exceeded` failures in `packages/document/tool-office/src/errors.ts`.
- [ ] T035 [US2] Verify: consume the T041-recorded session snapshot for tool output plus `corepack pnpm run test:coverage` for the tool package. **Tool package coverage is 100%; no session-level `document_*` tool-output fixture exists yet — it requires recording with a key. Deferred (2026-11-16): macOS/Linux snapshot re-recording out of scope for now.**
- [x] T062 [US2] Implement the Host→renderer command bridge: the Host tool sends a correlated gateway command (local IPC on desktop, no document bytes off-device), the renderer applies it to the live unit and returns the correlated result, with timeouts and `document-not-open` in `packages/client/ui-office-editor/src/client/command-bridge.ts`; test in `packages/client/ui-office-editor/tests/command-bridge.spec.tsx`.
- [x] T056 [US2] `ChangeRecord` inverse/audit helpers are shipped as `undoChange` in `packages/session/session-document/src/document-changes.ts` (no `src/undo.ts`); runtime undo/redo is the editor's native stack (T054/T057).
- [x] T057 [P] [US2] Test undoing an agent edit through the native undo stack (SC-012) in `packages/client/ui-office-editor/tests/undo.spec.tsx`
- [x] T058 [P] [US2] Add a timed agent-edit visibility test for SC-003 in `packages/client/ui-office-editor/tests/performance.spec.tsx`

**Checkpoint**: US1 and US2 work; agent and user converge and conflicts are visible.

---

## Phase 5: User Story 3 - Documents persist as session entities (Priority: P2)

**Goal**: Open documents survive restart and are referenceable across turns.

**Independent Test**: Quickstart S6, S7 and the restart scenario.

- [x] T036 [US3] Emit `document/open|edit|close|rebind|missing` facts from `packages/core/session/src/` per `contracts/document-events.contract.md`.
- [x] T037 [US3] Implement the `document` projection and open-document system-prompt injection in `packages/session/session-document/src/document-projection.ts`; tests in `packages/session/session-document/tests/document-projection.spec.ts`.
- [x] T038 [US3] Implement restore-after-restart of the open-document set in `packages/session/session-document/src/document-restore.ts`. **Client restore runs from the first mounted Office body; a session with no Office tab has no trigger (layout persistence normally mounts them).**
- [x] T039 [P] [US3] Implement rename/move rebind and missing reporting in `packages/session/session-document/src/document-identity.ts`. **Host endpoints and missing trigger are wired; the client move/rebind trigger still awaits a move signal.**
- [x] T040 [P] [US3] Implement recovery/restore orchestration over the T077 draft Remote (read draft, compare `contentVersion`, offer restore, clear on save) in `packages/session/session-document/src/document-recovery.ts`; the draft store itself is T077 and client capture is T070.
- [ ] T041 [US3] Verify: session migration test, `corepack pnpm run test:snapshot`, and the TypeScript and Python SDK expected-output updates. **Migration/catalog pass (1448 tests). Snapshot replay: 5 office-caused tool-list deltas need re-recording with `DEEPSEEK_API_KEY` on macOS/Linux; ~190 Windows-only env diffs; SDK expected outputs show no office delta. Lockfile refreshed separately. Deferred (2026-11-16): macOS/Linux re-recording out of scope for now.**

**Checkpoint**: Documents are durable, restorable, and cross-turn referenceable.

---

## Phase 6: Polish & Cross-Cutting Concerns

**Purpose**: Tests for the cross-cutting guarantees, docs, and the final independent-identity release.

- [x] T042 [P] Add the privacy egress test proving no document bytes leave the device (SC-008) in `packages/test-support/tool-testkit/tests/egress-documents.spec.ts`.
- [x] T043 [P] Add the write-authority test proving out-of-root writes are refused (SC-011) in `packages/document/tool-office/tests/authority.spec.ts`.
- [x] T044 [P] Add the unsaved-work recovery test (SC-010) in `packages/session/session-document/tests/document-recovery.spec.ts`.
- [x] T045 [P] Document the packages and subsystem: `packages/document/office-exchange/README.md` Model Experience section, `packages/client/ui-office-editor/README.md`, `docs/subsystems/office-editing.md`.
- [x] T046 [P] Add config catalog entries and run `verify-package-dependencies`, `verify-export-jsdoc`, and `verify-client-packages` for the new packages. **Done: config catalog up to date; all three verify gates pass for the office packages.**
- [x] T047 Run `corepack pnpm run build:check`, `corepack pnpm run doc-sync`, and `corepack pnpm run test:coverage`; fix regressions. **Done: build:check passes; doc-sync office-caused failures fixed; office packages at 100% per-file coverage.**
- [ ] T048 Run quickstart scenarios S1–S11 and record the results. **Status: [quickstart-results.md](./quickstart-results.md) records S1, S2, and S4–S10 as automated evidence from 45 test files / 539 passing tests, the 3-test document egress spec, `verify-forwork-identity`, `verify-license-notices`, and a 43/43 `doc-sync` run; S3 remains unrun because no model provider is configured (`FORWORK_API_KEY` unset, no root `.env`) and S11 because signing identities are absent (T051). No GUI observation was performed. Stale generated catalogs and three renamed-package note imports were repaired to restore the documentation gates.**
- [x] T059 Add the selective upstream-follow policy (decision table, high-conflict paths, cherry-pick/recovery rules, sync log) to `docs/development.md#upstream-sync`, cross-reference it from `quickstart.md`, and amend Constitution VIII to v2.0.0.
- [x] T060 Assert MIT and Apache-2.0 notices in the packaged artifact (FR-012) in `scripts/verify-license-notices.ts`
- [x] T061 Add an offline test proving edit and save work with no document server configured (FR-010) in `packages/document/office-exchange/tests/offline.spec.ts`
- [x] T049 Brand/telemetry scan gate asserting zero `deepseek` references in shipped surfaces: add `scripts/verify-forwork-identity.ts` and wire it into `doc-sync`/gates (FR-009, SC-006).
- [x] T050 Execute the scripted identity-landing phase in Constitution V's step order, each step one revertible commit: (1) rename `@deepseek-ai/dsh-<name>` → `@forwork/<name>` (drop the `dsh-` segment; keep `/client`, `/types`, `/remote`, `/typert` subpaths) across package.json, imports, `tsconfig.base.json` paths, cordis.yml rows, and generated catalogs; (2) rename `dsh` CLI and `dsh-*` artifacts → `forwork` / `forwork-*`, and desktop brand → ForWork; (3) re-scope vendor `@deepseek-ai/*` → `@forwork/*` and update `vendor/README.md` manifests/upstream SHAs; (4) replace DeepSeek telemetry and default-model-adapter defaults; (5) run the T049 scan. Gate each step on green `build`/`typecheck`/package tests. **Landed: scope, CLI/brand/scan. Remaining deepseek references are third-party-notice, protocol, or external-dependency references, tracked by T089; signing is T051.**
- [ ] T051 Configure Windows and macOS signing identities and produce signed installers (SC-007); verify both launch. **macOS signing deferred (2026-11-16: no macOS/Linux adaptation for now); Windows signing still needs certificates.**
- [x] T089 [P] Replace `github.com/deepseek-ai/deepseek-harness` repository URLs with the ForWork fork once it exists; keep legally required upstream third-party notices. **Done: repository, homepage, issues, and discussions URLs now use the ForWork fork across 389 manifests, READMEs, pyproject, website config, the native packing test, and the assembled fixture; legally required upstream source URLs stay upstream. Repository URLs use `main`, which the fork now advertises as its GitHub default branch — projected docs links, the tutorial badges, and the i18n switcher examples all follow it. The docs-site hostname/base remain unchanged pending a domain decision.**
- [x] T090 [P] Migrate the protocol-stable `dsh.*`/`DSH_*` tokens (`dsh.bundle`, `dsh.client`, `dsh.profile`, `$DSH_HOME`, `dsh-app://`) only if a later session-format amendment justifies the risk; otherwise record a documented exception. **Done: documented exception — [ForWork identity exceptions](../../.agents/notes/implemented/process/2026-11-17-forwork-identity-exceptions.md) keeps the tokens as the shipped protocol vocabulary; migration needs a session-format or persistence amendment plus a home migration.**
- [x] T091 [P] Replace the external `@deepseek-ai/libreoffice-kit*` dependency with a ForWork-published or vendored equivalent, or record a documented exception in the identity scan's allowlist. **Done: documented exception — [ForWork identity exceptions](../../.agents/notes/implemented/process/2026-11-17-forwork-identity-exceptions.md) keeps the external kit and its exact scan allowlist.**
- [x] T063 [P] Construct `SafeStorageSecureStore` in the desktop host with Electron `safeStorage` and `app.getPath('userData')`, and connect it to the model-credential path; surface `secure-storage-unavailable` in `apps/desktop/` and `packages/credentials/credentials/src/`.
- [x] T069 [P] Classify the `officeExchange`/`officeDocument` product roles so `gen-doc-graphs --check` regenerates `docs/capability-seams.md`; update the service-role classification and re-run the doc gates.
- [x] T076 Classify the four `officeDocument` Remote endpoints (`read`, `save`, `detectUnsupported`, `convertLegacy`) in `packages/experimental/liveshare/src/client/endpoint-classification.ts` so the client aggregate typecheck passes.

---

## Dependencies & Execution Order

### Phase Dependencies

- **Setup (Phase 1)**: T001–T003 done; T004–T007 can start now.
- **Foundational (Phase 2)**: T008–T020 block every user story.
  - T008 (xlsx) and T009 (docx) are independent file sets and may run in parallel.
  - T014 (events) blocks T036–T041.
  - T015–T017 block T031–T034.
  - T016 (Change Record inverses) blocks T056 and T057.
  - T062 (renderer command bridge) blocks T031 for open-document edits (T056 was folded into the shipped `undoChange`).
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
- The identity-landing phase landed (T050); residual identity work is T089–T091 and signing is T051 (Constitution V).
- Every task declares its own verification; a later task does not assume an earlier check passed.

---

## Phase 7: Convergence

Convergence on 2026-11-16 found the remaining unbuilt work below; open tasks T028, T035, T041–T048, T051, T089–T091, and the Phase 8 gates stay authoritative for the items they already name.

- [x] T079 [P] [US1] Update `apps/web/tests/document-preview.e2e.ts` and `snapshots/web/document-preview/document.expected.md` so `.docx`, `.xlsx`, `.doc`, `.xls`, and `.ppt` open the `office` tab and only `.pptx` stays in the read-only preview, then replay the web lane named in T028, leaving the shipped roster and SDK deltas to T041 per FR-001/FR-011 (partial)
- [x] T080 [P] [US3] Call `officeDocument.rebindDocument` from the `workspaceFiles` change feed when an open document's path disappears and the same version reappears, recording `document/rebind`, in `packages/client/ui-office-editor/src/client/changes.ts` with tests in `packages/client/ui-office-editor/tests/changes.client.spec.ts`, so `rebindDocument` gains its production caller per FR-015/SC-009 (missing)
- [x] T081 [P] [US3] Trigger the one restore pass from Session activation as well as from a mounted Office body in `packages/client/ui-office-editor/src/client/restore.ts` and `packages/client/ui-office-editor/src/client/index.ts`, with coverage in `packages/client/ui-office-editor/tests/restore.client.spec.tsx`, so a Session with no Office tab open still restores its open set per FR-007/SC-005 (partial)
- [x] T082 [P] [US1] Adopt the Host-reported content version when the Host's own save reloads a tab in `packages/client/ui-office-editor/src/client/useOfficeDocument.ts` instead of treating it as an external change, so the tab and the Host keep one version state; test in `packages/client/ui-office-editor/tests/use-office-document.spec.tsx` per FR-005/FR-020 (partial)
- [x] T083 [US1] Add the FR-011 acceptance scenario to `specs/001-office-agent-framework/spec.md` and a fixture-backed render check that opens a `.pptx` fixture in the read-only preview and asserts no `office` tab or edit surface in `apps/web/tests/document-preview.e2e.ts`, after T079 per FR-011 (missing)
- [x] T084 [P] Correct the accepted-save and command wording: add `save` to FR-006 and the Document Session Fact entity in `specs/001-office-agent-framework/spec.md`, document `document/save` and the ignorable `document/command` transport event in `specs/001-office-agent-framework/contracts/document-events.contract.md`, and fix the entity overview and command entry in `specs/001-office-agent-framework/data-model.md` per FR-006 (partial)
- [x] T085 [P] Update the stale plan and quickstart text: add `contracts/command-bridge.contract.md` and the real dependencies (ExcelJS, fast-xml-parser, fflate, libreoffice-kit; no SheetJS or standalone docx library), and drop the nonexistent `session-projection` fold and `undo` entries in `specs/001-office-agent-framework/plan.md`; reference the command-bridge contract in `specs/001-office-agent-framework/quickstart.md` per plan: technology and structure decisions (contradicts)
- [x] T086 [P] Add the bilingual twins: `packages/document/office-exchange/README.zh.md` with `README.i18n.yaml` (its English README already links the missing counterpart), `packages/document/tool-office/README.zh.md` with `README.i18n.yaml`, `packages/session/session-document/README.zh.md` with `README.i18n.yaml`, and after T045 `docs/subsystems/office-editing.zh.md` with `.i18n.yaml`; re-record with `verify-translation-pairing --write` per plan: documentation (partial)
- [x] T087 [P] [US2] Extend the unsaved-work coverage in T044 with the abnormal-exit and dirty-close branches of SC-010 (a killed process offers its draft on relaunch and a dirty close warns) across `packages/session/session-document/tests/document-recovery.spec.ts` and `packages/client/ui-office-editor/tests/draft-recovery.spec.tsx` per SC-010/T044 (partial)
- [x] T088 [P] Close or rewrite T056 to its shipped form (`undoChange` in `packages/session/session-document/src/document-changes.ts`; `src/undo.ts` does not exist) and correct the stale "T062 blocks T031 and T056" US2 dependency note in `specs/001-office-agent-framework/tasks.md` per T056 (contradicts)

---

## Phase 8: Identity-landing follow-up

The identity landing surfaced three repository gates that are not yet tracked; they are cross-cutting, not story work.

- [x] T092 Fix the client bundle manifest so `check-workspace-constraints` accepts the declared entry: `packages/client/ui-office-editor` currently declares `lib/client*.js` where the policy expects `lib/client.js`; adjust `packages/client/ui-office-editor/package.json` and `packages/client/ui-office-editor/tsdown.config.ts`.
- [x] T093 Make `rescope-vendor:check` green: fix the postcondition text mismatch in `scripts/rescope-vendor.ts` and remove the four residue files it reports, so the vendored re-scope gate passes.
- [x] T094 Rename `scripts/verify-dsh-package-licenses.ts` and any remaining `dsh-*` script/artifact names to `forwork-*`, updating their references in `scripts/run-gates.ts` and the root `package.json`.

## Phase 9: Convergence

Convergence on 2026-11-17 re-assessed every requirement, success criterion, acceptance scenario, plan decision, and constitution principle against the current code; checkbox state and earlier completion claims were not treated as evidence. Open T028, T035, T041, T048, and T051 stay authoritative for the verification they already name.

- [x] **CRITICAL** T095 Make default-on egress opt-in: `@forwork/session-log-deepseek` ships `enabled` defaulting to `true` in `base` and `sdk-minimal` and uploads the canonical log — including `document/edit` operation data and inverses — with official provider requests, while `@forwork/client-product-analytics` ships `enabled: true` for desktop; default both to off, exclude `document/*` facts from the log field, and add executed tests per FR-013, SC-008, and Constitution II (contradicts). **Done (2026-11-17): session-log and desktop analytics default off; `document/*` facts are never uploaded.**
- [x] **CRITICAL** T096 Refuse or preserve the OOXML constructs a save currently drops: `.docx` images (`word/media/`), hyperlinks, lists and numbering, run fonts, colors, sizes, alignment, merged-cell spans, and nested tables, and `.xlsx` conditional formatting, data validation, freeze panes, tab colors, and hyperlink targets are neither detected nor preserved; extend `detectUnsupportedFeatures` or the save hook, and add one refusal test per kind plus a preservation round-trip per FR-002, FR-008, SC-002, SC-004, and Constitution IV (contradicts). **Done (2026-11-17): detection now refuses docx images/hyperlinks/numbering/unsupported run formatting/merged spans/nested tables and xlsx conditional formatting/data validation/freeze panes/tab colors/hyperlink targets.**
- [x] **CRITICAL** T097 Reconcile Constitution I with the approved Office-kit exception: Principle I calls any upstream scope in a shipped manifest a defect while `apps/desktop-host`, `bundle/web-app`, `skill/skill-office`, `document/office-to-pdf`, and `document/office-exchange` declare `@deepseek-ai/libreoffice-kit` under the FR-009 and T091 allowlist; amend the principle with the recorded rationale and migration note, or vendor the dependency, per Constitution I and its Governance section (contradicts). **Done (2026-11-17): Constitution v3.1.0 records the external-dependency exception in Principle I; see `.agents/notes/implemented/process/2026-11-17-forwork-identity-exceptions.md`.**
- [x] T098 Keep the editor's native undo history across the save that follows an accepted agent edit: the self-caused reload and remount dispose the stack holding the agent edit, and `undoChange` has no production caller; add a test that applies an agent edit through the live bridge, saves, undoes, and reads the original value back per FR-014 and SC-012 (missing). **Done (2026-11-17): the agent-command save is classified as self-caused so the unit is not remounted and the native undo stack survives.**
- [x] T099 Attribute an agent-triggered save to the agent and keep one document identity: the tool saves through `officeDocument.save`, which always records `actor: 'user'` and opens a second path-spelt entry when the tool's absolute path differs from the editor tab's identity, and the declared `turn` field is never populated; fix the identity and actor in `office-exchange/src/remote.ts` and `tool-office/src/edit.ts` with covering tests per FR-006 and FR-014 (contradicts). **Done (2026-11-17): one documentId per file, save carries the real actor, and turn is recorded when a turn exists.**
- [x] T100 Make the open-document prompt section addressable: `session-document/src/plugin.ts` renders only `basename (format)` while the document tools require `<sessionId>:<path>`, so a document opened from the sidebar has no id the model can pass; include a stable bounded id or path and test that a later turn resolves an open document from the prompt text per US3/AC2 and FR-006 (partial). **Done (2026-11-17): each line renders its `<sessionId>:<path>` id so the model can pass it to the document tools.**
- [x] T101 Extend `verify-forwork-identity` to the surfaces FR-009 names — `python/` distributions and modules, the root READMEs, `docs/`, `website/`, `apps/*/config`, `apps/desktop/scripts`, lowercase `deepseek-harness`, and standalone `DeepSeek` — with rejection tests, then fix the defects it reports in the request identity (`llm/src/attribution.ts` product and URL), the README documentation host, and the website wordmark per FR-009 and SC-006 (partial). **Done (2026-11-17): scanner covers python/README/docs/website/apps configs, adds the deepseek-harness and standalone DeepSeek rules with a rejection spec, and fixes the request User-Agent, README tagline/host, and website wordmark.**
- [ ] T102 Rename or record an amendment for the DeepSeek-named shipped identifiers FR-009 does not allow: the PyPI distributions `deepseek-harness-sdk` and `deepseek-harness-runtime-bin`, the modules `deepseek_harness` and `deepseek_harness_runtime`, and the `@forwork/*deepseek*` package names with their locale copy per FR-009 (contradicts).
- [x] T103 Replace the DeepSeek installer and company identity so a signed build can satisfy SC-007: the desktop `author`, the `com.deepseek.dsh` and `com.deepseek.harness` app ids, the `com.deepseek.harness` notarization cache paths, and the `download.deepseek.com` origin per SC-007, FR-009, and Constitution I; T051 still owns the certificates (contradicts). **Done (2026-11-17): ForWork author, `com.forwork.*` app ids, and GitHub-releases origin.**
- [x] T104 Re-address the live unit on rebind: the change tracker updates its own path but the mounted tab keeps the old address and saves to it, and restore reopens by path so a duplicate identity entry can appear; add edit-after-move and rebind-then-restart tests per FR-015 and SC-009 (partial). **Done (2026-11-17): the mounted tab re-addresses to the new path under the same documentId and restore reopens by documentId.**
- [x] T105 Flush pending drafts on quit and unload: the draft `flush()` has no production caller and capture is idle-debounced by two seconds, so an edit inside that window on crash or quit is neither warned about nor recoverable; wire the flush and add quit, unload, and process-kill coverage per FR-016 and SC-010 (partial). **Done (2026-11-17): drafts flush on tab close and page unload, and the first edit writes immediately.**
- [x] T106 Return `document-not-open` when no renderer is attached instead of the current `timeout`, and add an end-to-end tool to broker to client bridge to runtime test per FR-020 (partial). **Done (2026-11-17): no registered target returns `document-not-open` immediately; a registered-but-silent target keeps `timeout`.**
- [x] T107 Consume `Config.maxConversionMs` in legacy conversion: the bound is validated and deployed but never passed to the conversion call, so an override has no effect per SC-001 (partial). **Done (2026-11-17): the bound now aborts in-flight conversion with `too-large` and leaves the source unchanged.**
- [ ] T108 Store user credentials in the OS secure store on every shipped surface or record an explicit exception: CLI and Web keep them in a plaintext `$DSH_HOME/.credentials.yaml` at `0600` per FR-018 (partial).
- [ ] T109 Fix restore-pass bookkeeping: the first mounted Office body claims the pass even when the projection is empty, so a projection that arrives later in that window never restores per FR-007 and SC-005 (partial).
- [ ] T110 Require a mounted editor unit before saving: `canSave` checks only the snapshot, and `persist` can call `save(undefined)` in the gap before `mountEditor` resolves per FR-002 (partial).
- [ ] T111 Add a spec for `verify-license-notices` that covers the MIT and Apache-2.0 acceptance and rejection paths, so the FR-012 and Constitution VII evidence is executed per FR-012 and Constitution VI (missing).
- [ ] T112 Reconcile the `.ppt` acceptance wording: US1/AC4 says `.ppt` opens an editable office tab while FR-003 and FR-011 make it a read-only converted `.pptx`, which the code follows; restate the scenario's expected outcome per US1/AC4, FR-003, and FR-011 (contradicts).
- [x] T113 Correct the Constitution Technology Constraints to the shipped conversion stack: `.docx` is hand-written OOXML over `fast-xml-parser` and `fflate` with no docx library, and SheetJS CE is used only for legacy `.xls` preview per the Constitution Technology Constraints section (contradicts). **Done (2026-11-17): Constitution v3.1.0 corrects the conversion stack.**
- [ ] T114 Extend legacy conversion immutability coverage from `.xls` to `.doc` and `.ppt`, asserting the source hash is unchanged per FR-003 and SC-002 (partial).
- [ ] T115 Surface the converted-copy notice on the `.ppt` converted preview tab, which never reads the notice store per FR-003 (partial).
- [ ] T116 Assert the SC-003 visible-edit bound: no test asserts apply-to-visible latency, and the default `commandTimeoutMs` is five seconds rather than the stated two per SC-003 (partial).

## Phase 10: Convergence Merge (2026-11-17)

- [x] T117 Execute the declared `formula`, `row`, `column`, and `style` steps in the live editor: `editor/sheets.ts:67` and `editor/docs.ts:63` accept only `cell` and `paragraph`, so an accepted operation with any other kind returns `unsupported-feature` even though the operation union and the tool schemas declare it; implement those steps in the Univer adapter or remove them from the tool schema and record the limitation, per FR-004 and FR-020 (contradicts). **Done (2026-11-17): the Univer sheets adapter now applies formula/row/column/style steps atomically; no kind was removed from the tool schema, and the docs adapter still refuses sheet-only kinds.**
- [ ] T118 Align the legacy-conversion tab behaviour: `OfficePanel.tsx:121` calls `openResource(address, { replaceTab: true })` while the clarification and T073 say the converted copy opens in a new tab; choose one and update the code or the spec and task wording, with a test asserting the chosen behaviour, per FR-003 and US1 (contradicts).
- [x] T119 Enforce the `officeDocument` write postcondition: `contracts/office-exchange.contract.md:26` requires a save to refuse when the source's unsupported set is non-empty, but the provider save path checks only the content version; add the refusal at the save boundary plus a contract test, extending T096 (partial). **Done (2026-11-17): the Host save re-detects the source and refuses with `unsupported-feature` before writing.**
- [x] **HIGH** T120 Replace the DeepSeek brand artwork and gate it: `FishLogo` and `BrandWordmark`, `website/public/wordmark.svg`, `website/public/favicon.svg`, `apps/desktop/resources/{icon,icon-windows,icon-macos,tray-windows}`, `apps/desktop/renderer/assets/welcome-brand.svg`, `apps/desktop/installer/assets/*`, and the `home-whale` pixel animation carry DeepSeek marks that FR-009 and SC-006 do not allow; produce ForWork-owned marks at the existing sizes, decide replace-or-remove for the home animation, add `verify-brand-assets` with an asset manifest, upstream-path fingerprints, and a rejection spec, and update the affected desktop and brand snapshots per FR-009, SC-006, and Constitution I (contradicts). **Done (2026-11-17): the door mark, lockup, platform icons, favicons, welcome/installer art, `verify-brand-assets`, and 13 finished-product screenshots are landed; the whale atlas is removed.**
- [ ] T121 Correct task and plan references: T036 and T040 name the wrong emitting files, T066 writes `packages/fs/fs-ssh` for `packages/ssh/fs-ssh`, `provider/detect.ts` is `src/detect.ts`, and `command-bridge.spec.tsx` is `command-bridge.client.spec.ts`; plan.md omits `apps/desktop-host`, `apps/web`, `bundle/web-app`, `ssh/fs-ssh`, and `experimental/liveshare` and describes conversion and credentials roles that no longer match (partial).
- [ ] T122 Correct the remaining specification and contract prose: the office-exchange operation union omits `read` and `timeout`, `document_convert` omits the `.ppt` refusal, the office-exchange README describes a superseded surface, the data-model entity and field list drifted from the events, SC status still reads implemented while T028/T035/T041/T048/T051 are open, contract signatures omit `outputDir`, and the snapshot label calls a converted `.ppt` an editable tab per FR-003 and the current contracts (partial).
- [ ] T123 Decide the default workspace directory name: `packages/api/workspace-controller/src/default-workspace.ts:11` and `default-directory.ts:77` still create a `deepseek-harness` folder, a user-visible brand leak that FR-009 and SC-006 do not allow; rename to a ForWork-owned path with an in-place migration for existing workspaces (or record the migration decision), plus a test per FR-009 and SC-006 (contradicts).

## Phase 11: Convergence

- [x] T124 Detect and refuse the remaining `.docx` constructs the writer still drops without detection — run and paragraph formatting beyond font, color, size, and alignment (strike, highlight, spacing, indentation, borders), footnotes and endnotes, sections and page layout, theme styles, and images that only reference an external target — with one reject and one no-false-positive test per construct, per FR-008, SC-002, and SC-004 (partial). **Done (2026-11-17): strike, highlight, spacing, indentation, borders, footnotes, endnotes, external images, and multiple sections are now detected; single-section page setup and theme styles remain reported limitations.**
- [x] T125 Expose the `row`, `column`, and `style` operation kinds the live adapter now applies: `document_edit_cells` declares no row-insert/delete, column-insert/delete, or style steps, so the agent cannot issue them; either add those steps to the tool schema or formally descope them in the spec and data model, per FR-004 and FR-020 (partial). **Done (2026-11-18): row/column/style steps are now exposed in document_edit_cells and applied atomically to the live editor.**
- [x] T126 Make the web office fixture lossless and cover the refusal path: the docx branch of apps/web/tests/office-fixture.ts emits rFonts/sz/color/jc that the run-formatting detector now refuses, so regenerate it as a minimal editable document (Chinese text, a w:br page break, a single default sectPr) and add a positive assertion that a docx carrying <w:rPr><w:rFonts/></w:rPr> opens into the unsupported state with Save disabled, plus README and quickstart wording that explicit run/paragraph formatting is refused rather than lost, per FR-008, SC-004, and T096 (partial). **Done (2026-11-17): the docx fixture is now lossless and a positive refusal assertion covers the run-formatting path.**
- [x] T127 Add and mount the Univer sheet plugins for conditional formatting, data validation, hyperlinks, and images, then map those xlsx features (ExcelJS → IWorkbookData resource hooks / plugin models) so they round-trip and are editable; retire the corresponding `conditional-formatting`, `data-validation`, `hyperlink` (worksheet branch), and `image` (worksheet `xl/media/` branch) detection once each is preserved, per FR-022 (partial). **Done (2026-11-18): conditional formatting, data validation, hyperlinks, and images now round-trip and render via the four sheet plugins.**

## Phase 12: ForWork Brand Artwork Design (T120 design half)

**Goal**: Produce ForWork-owned vector masters and every derived bitmap at the existing sizes, then adopt them in components and static surfaces so T120's gate has real assets to enforce.

**Independent Test**: Render every asset from its master, run `corepack pnpm run verify-brand-assets`, and inspect the mark at 16, 24, 50, 192, 512, and 1024 px on light and dark backgrounds; no DeepSeek contour, letterform, or blue-whale favicon remains.

**Design baseline**: A 24-unit grid half-open door mark — frame at `x=4..20, y=3..21`, leaf hinged at `(7,7)`, handle hole at `(14,14)` — on a flat `#F4F2EC` tile with `#141414` ink and a `#D8D3C8` hairline. The wordmark pairs the mark with "ForWork" at weight 600. See `assets/brand/README.md`.

### Design Foundations

- [x] T127 Write the ForWork brand design brief in `assets/brand/README.md`: door construction grid, wordmark spacing, flat palette, clear space, minimum sizes, platform safe zones, derivation commands, and the recorded removal decision for the home animation.
- [x] T128 [P] Author the canonical monochrome door mark in `assets/brand/forwork-mark.svg` with a 24×24 viewBox and `currentColor` fills, so sidebar, hero, tray, and favicon derive from one geometry.
- [x] T129 [P] Author the 1024×1024 app tile master in `assets/brand/forwork-tile.svg` with the `#F4F2EC` field, `#141414` door, handle cutout, and a 1024-unit artwork extent that survives the 0.4 safe-radius maskable circle.
- [x] T130 [P] Author the lockup master in `assets/brand/forwork-lockup.svg` with the door mark and "ForWork" lettering, preserving the existing `BrandWordmark` contract: 182 units with the mark at `viewBox="0 0 182 24"` and 156 units without it at `viewBox="26 0 156 24"`.

### Surface Derivation

- [x] T131 [P] Derive the desktop platform sources from T129 — `apps/desktop/resources/icon.svg` (1104-unit canvas with the 40/32 inset and shadow groups), `apps/desktop/resources/icon-windows.svg` (192-radius tile plus `id="tray-glyph"` around the artwork), and `apps/desktop/resources/icon-macos.svg` (230-radius tile) — by adding a derivation script at `scripts/render-brand-assets.ts` and committing its generated output.
- [x] T132 [P] Derive the website assets from T128 and T130: `website/public/wordmark.svg` (mark plus ForWork lettering at 182×24, `currentColor` fills) and `website/public/favicon.svg` (tile-backed mark, not a bare glyph), keeping the `dsh-wordmark` class hook in `website/.vitepress/config.ts:187` working.
- [x] T133 [P] Derive the welcome and installer bitmaps from T129 and T130 at their existing pixel sizes: `apps/desktop/renderer/assets/welcome-brand.svg` (472×40), `apps/desktop/installer/assets/brand.png` and `brand-2x.png` (light), and `brand-dark.png` and `brand-dark-2x.png` (dark), then verify each in `apps/desktop/tests/welcome-renderer.client.spec.tsx` and `apps/desktop/tests/installer-packaging.spec.ts`.
- [x] T134 Regenerate the Web install icons from the T129-derived `apps/desktop/resources/icon.svg` by running `corepack pnpm --filter @forwork/web-frontend run render:app-icons`, committing `apps/web/public/icon-192.png`, `icon-512.png`, `icon-maskable-512.png`, and `apple-touch-icon.png`, and keeping `apps/web/tests/pwa-manifest.e2e.ts` within the 0.4 safe-radius assertion.
- [x] T135 Regenerate `apps/desktop/resources/tray-windows.ico` from the T131 Windows source by running `corepack pnpm run render:tray-icon` in `apps/desktop`, confirm the mark stays legible at 16 px, and update the door wording in `apps/desktop/scripts/render-tray-icon.ts:2`.
- [x] T136 [P] Remove the DeepSeek-derived home animation: delete `packages/client/ui-brand-official/src/client/home-whale/` and `breach-atlas.webp`, drop the `PixelWhale` import and `conversation.hero.frame` registration, delete its unit/e2e tests, and update the package metadata, slot catalog, and docs; keep the slot declared for other bundles.

### Component Adoption

- [x] T137 Replace `FishLogo` with the T128 mark in `packages/client/ui-primitives/src/BrandMark.tsx` (exporting `BRAND_MARK_PATH`, `BRAND_MARK_FRAME_PATH`, `BRAND_MARK_LEAF_PATH`, and `BRAND_MARK_VIEWBOX`), update `packages/client/ui-primitives/src/index.ts`, and update every consumer: `packages/client/ui-sidebar/src/client/SidebarRoot.tsx:297`, `packages/client/ui-brand-official/src/client/Brand.tsx:11`, and `packages/experimental/liveshare/src/client/viewer-brand.tsx:61`.
- [x] T138 Replace `BrandWordmark` with the T130 lockup in `packages/client/ui-primitives/src/BrandWordmark.tsx`, preserving the `includeMark` prop and the 182/156 width and viewBox contract, and update `packages/client/ui-primitives/tests/icons.client.spec.tsx`.
- [x] T139 Replace the fish morph in `packages/client/ui-conversation/src/client/skeleton/EmptyHero.tsx:74` with a door-leaf rotation around its hinge that keeps the reduced-motion fallback and the `conversation.hero.brand.mark` slot, and remove the now-unused `HERO_SWIM_UP_PATH` and `HERO_SWIM_DOWN_PATH` constants.

### Gate and Evidence

- [x] T140 [P] Add `scripts/verify-brand-assets.ts`, `scripts/brand-assets.manifest.json`, `scripts/verify-brand-assets.spec.ts`, and `scripts/fixtures/brand/`: the manifest records every shipped brand asset path, SHA-256, generator, dimensions, kind, and the 13 reviewed evidence rows; the verifier rejects a missing or stale manifest entry, an unreferenced asset, a dimension mismatch, missing evidence, and any fingerprint match against the retired DeepSeek whale, wordmark, blue favicon, or home atlas.
- [x] T141 Wire `verify-brand-assets` into `package.json` and `scripts/run-gates.ts` beside `forwork-identity`, then update the affected tests and snapshots: delete the whale unit/e2e tests, keep `packages/client/ui-brand-official/tests/browser-plugin.client.spec.tsx` covering the empty frame, and update `apps/desktop/tests/welcome-renderer.client.spec.tsx`, `apps/desktop/tests/installer-packaging.spec.ts`, `apps/web/tests/app-icons.spec.ts`, and `apps/web/tests/pwa-manifest.e2e.ts`.
- [x] T142 [P] Update the surface documentation and record the design decision: `packages/client/ui-brand-official/README.md`, `README.zh.md`, `apps/desktop/README.md:88`, `README.zh.md`, `docs/subsystems/slots.md`, `slots.zh.md`, and `assets/brand/README.md`, replacing whale/fish artwork descriptions with the door mark, the removed frame layer, and the flat palette; link the design brief from T127.

### Checkpoint

`verify-brand-assets`, `verify-forwork-identity`, the icon and brand unit tests, and the Web/Desktop snapshot lanes pass; every listed asset derives from `assets/brand/` and no retired DeepSeek contour remains.

### Dependencies

- T127 blocks T128–T130.
- T128 and T129 block T131–T136.
- T130 blocks T132 and T133.
- T128 and T130 block T137 and T138.
- T137 blocks T139.
- T140 blocks T141.
- T142 can run after T127 and in parallel with T140/T141.

## Phase 13: Finished-Product Brand Evidence (T120 evidence half)

**Goal**: Prove the shipped door mark in the running Web app, the desktop welcome renderer, the installer artwork, and the built website, with committed screenshots and a manifest that binds each image to the assets it depicts.

**Independent Test**: `corepack pnpm exec vitest run --config vitest.brand.config.ts` and `corepack pnpm run verify-brand-assets`; inspect the 13 committed images and confirm the door mark, ForWork label, and no whale.

- [x] T143 Add `vitest.brand.config.ts` with the four evidence suites, serial file execution, and the host-plane tsconfig entries for `apps/web/tests/brand-*.e2e.ts`.
- [x] T144 Add `apps/web/tests/brand-artwork.e2e.ts` and commit `home-desktop-{light,dark}.png`, `home-hover.png`, `home-mobile.png`, `rail-collapsed.png`, and `icon-sheet.png` under `apps/web/tests/expected/brand/`.
- [x] T145 Add `apps/web/tests/brand-welcome.e2e.ts` and commit `welcome-{light,dark}.png` under `apps/desktop/tests/expected/brand/`.
- [x] T146 Add `apps/web/tests/brand-installer.e2e.ts` and commit `installer-{light,dark}.png` and `uninstaller.png` under `apps/desktop/tests/expected/brand/`.
- [x] T147 Add `apps/web/tests/brand-website.e2e.ts` and commit `nav-{desktop,mobile}.png` under `website/tests/expected/brand/`.
- [x] T148 Extend `scripts/brand-assets.manifest.json` with the 13 evidence rows, record their hashes with `verify-brand-assets --write`, and verify the evidence inventory and dimensions.
- [x] T149 Rename the local-build label from `DSH Local Build` to `ForWork Local Build` in the locale dictionaries, Vite title, tests, snapshots, and packaging docs so the screenshots show the ForWork product.
- [x] T150 Record `brand-home-hover.gif` from the real Web server with Playwright video and ffmpeg; the 960×600, 5.8 s artifact is at `D:\kilo\output\forwork-brand\brand-home-hover.gif` for the PR attachment.
- [x] T151 Update the plan, contract, and quickstart to the consolidated `apps/web/tests/brand-*.e2e.ts` lane and the final evidence paths.

### Checkpoint

`vitest.brand.config.ts` passes, all 13 evidence images exist with matching manifest hashes, `verify-brand-assets` passes, and the Web screenshot shows the door mark and the ForWork label in both themes.

## Phase 12: Convergence

Convergence on 2026-11-18 records that the mapping-fidelity expansion (Waves 1-6, T125, and T127) superseded the earlier detect-and-refuse work recorded in T096 and T124: run and paragraph formatting, borders, sections, page setup, headers and footers, numbering, hyperlinks, tables, images, and footnotes are now mapped into the Univer document model and editable or losslessly preserved, so those completed tasks' Done notes describe a refusal state the expansion intentionally reversed. The refusals that remain are the unrepresentable set in FR-008: `.docx` macros, tracked changes, comments, charts, drawings/shapes, and embedded objects, and `.xlsx` charts, drawings, comments, pivot tables, and macros.
