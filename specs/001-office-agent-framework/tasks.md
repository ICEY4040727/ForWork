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
