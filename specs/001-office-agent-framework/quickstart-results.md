# Quickstart Validation Results: Office Agent Framework

**Recorded**: 2026-11-16, worktree `H:\ForWork\.kilo\worktrees\protective-talk`, branch `protective-talk`, at the branch head reported by `git rev-parse HEAD`. The commit identifier is omitted from this file because `verify-repository-references` rejects commit hashes in maintained files; the same command reported a single clean head with no local modifications before the run.

**Scope**: [quickstart.md](./quickstart.md) scenarios S1–S11 and its "Automated checks" block. The checks in this record are the automated evidence available in this worktree; no GUI interaction was performed.

## Environment

| Item | Value |
| --- | --- |
| OS build | Windows 10.0.26100.0 |
| Node.js | v24.15.0 |
| pnpm (corepack) | 11.28.5 |
| `FORWORK_API_KEY` | not set |
| Root `.env` | absent |
| Model provider configured | no |

The worktree had a clean tree at HEAD before the run. Dependency installation was already present; no `corepack pnpm install` was executed.

## Commands and results

| # | Command | Exit | Observed summary |
| --- | --- | --- | --- |
| 1 | `corepack pnpm run build:artifacts` | 0 | 8 gates passed; 80.46s; 525 client artifacts recorded |
| 2 | `corepack pnpm run build:check` | 0 | 2 gates passed (host and client aggregate checks); 15.50s |
| 3 | `corepack pnpm run doc-sync` (first run) | 1 | 40 passed, 3 failed: cordis catalog, client catalog, doc-typecheck |
| 4 | `corepack pnpm run gen-client-catalog` | 0 | wrote `packages/extensions/cordis-client-runner/src/client/slot-catalog.ts` |
| 5 | `corepack pnpm run gen-cordis-catalog` | 0 | 131 artifacts computed, 3 written |
| 6 | `corepack pnpm run verify-translation-pairing --write <three notes>` | 0 | 0 records written; pairs validate |
| 7 | `corepack pnpm run doc-sync` (second run) | 1 | 42 passed, 1 failed: CLI help reference (30s child timeout under parallel gate load) |
| 8 | `corepack pnpm run verify-cli-help` | 0 | 7 help pages; outputs current |
| 9 | `corepack pnpm run doc-sync` (third run) | 0 | 43 passed, 0 failed; 89.22s |
| 10 | `corepack pnpm exec vitest run packages/document/office-exchange/tests packages/document/tool-office/tests packages/client/ui-office-editor/tests packages/session/session-document/tests` | 0 | 45 test files passed, 539 tests passed; 17.50s |
| 11 | `corepack pnpm exec vitest run packages/test-support/tool-testkit/tests/egress-documents.spec.ts` | 0 | 1 file passed, 3 tests passed |
| 12 | `corepack pnpm run verify-forwork-identity` | 0 | shipped surfaces carry ForWork identity |
| 13 | `corepack pnpm run verify-license-notices` | 0 | 9 MIT base and 4 Apache-2.0 Univer entries checked |
| 14 | `corepack pnpm run test:docs` | 0 | 21 gates passed; 19.85s |
| 15 | `corepack pnpm run hygiene` | 0 | 25 gates passed; 190.95s |

Suite file counts inside command 10: `packages/document/office-exchange/tests` 15, `packages/document/tool-office/tests` 4, `packages/client/ui-office-editor/tests` 19, `packages/session/session-document/tests` 7.

### Failures found and repaired

The first `doc-sync` run exposed three stale states; none required a policy change.

- **CLI help reference** failed once with exit 1 because the `acp --help` child process printed its usage text and was then killed at the 30-second subprocess timeout while eight gate workers ran beside it. The same gate passed standalone (command 8) and in the third `doc-sync` run (command 9). Its checked output was current throughout.
- **Cordis catalog** failed because `EVENT_SCOPE_PAGE` in `scripts/gen-cordis-catalog.ts` still mapped `account`, while the projection renders the `deepseek-account` event scope documented in [docs/subsystems/credentials.md](../../docs/subsystems/credentials.md). The mapping now names `deepseek-account`; regeneration updated its generated regions.
- **Client catalog** reported `packages/extensions/cordis-client-runner/src/client/slot-catalog.ts` stale; `gen-client-catalog` regenerated it.
- **doc-typecheck** failed on three Agent Note TypeScript fences that imported packages under their pre-rename names. The fences now import `@forwork/web`, `@forwork/session`, and `@forwork/brand`; the English and Chinese copies carry the same code.

Regenerated or repaired files:

- `scripts/gen-cordis-catalog.ts` (event-scope mapping)
- `docs/subsystems/core.md`, `docs/subsystems/core.zh.md` (generated API region)
- `packages/extensions/tool-cordis/src/api-catalog.ts` (generated catalog)
- `packages/extensions/cordis-client-runner/src/client/slot-catalog.ts` (generated catalog)
- `.agents/notes/implemented/architecture/2026-06-24-web-capability-seam.md` and `.zh.md`
- `.agents/notes/implemented/architecture/2026-08-27-adjacent-agent-steer-messaging.md` and `.zh.md`
- `.agents/notes/proposed/feature/2026-08-04-task-surface.md` and `.zh.md`

## Scenario results

### S1 — Edit and save `.xlsx` losslessly (US1)

- **Status**: automated evidence
- **Coverage**: `packages/document/office-exchange/tests/provider-xlsx.spec.ts` — "reads workbook values, formulas, styles, and merges from the fixture" and "preserves supported semantics through a write/read round-trip"; `packages/document/office-exchange/tests/remote.spec.ts` — "saves a modified workbook at the read version and reads the change back".
- **Observed**: the fixture's formulas, styles, merges, and layout survive a write/read round-trip, and a saved workbook reads back the modified cell at the next content version, in the passing suite run (command 10).

### S2 — Edit and save `.docx` losslessly (US1)

- **Status**: automated evidence
- **Coverage**: `packages/document/office-exchange/tests/provider-docx.spec.ts` — "reads the fixture heading, bold run, and 2x2 table" and "preserves supported semantics through a write/read round-trip"; `packages/document/office-exchange/tests/remote.spec.ts` — "saves a modified document at the read version and reads the change back".
- **Observed**: headings, runs, and tables survive a write/read round-trip, and a saved document reads back its paragraph edit, in the passing suite run (command 10).

### S3 — Agent edit converges (US2)

- **Status**: not run
- **Coverage**: no automated substitute for the live editor-and-model observation.
- **Reason**: the scenario needs a configured model provider. `FORWORK_API_KEY` is not set and no root `.env` exists in this worktree, so no model turn could be issued.

### S4 — Conflict is rejected (US2/FR-005)

- **Status**: automated evidence
- **Coverage**: `packages/document/office-exchange/tests/remote.spec.ts` — "refuses a stale save as a conflict and leaves the file unchanged" and "refuses a save whose file disappeared as a conflict"; `packages/document/tool-office/tests/document-tools.spec.ts` — "maps conflict and unsupported answers".
- **Observed**: a stale save is refused with `office-document/conflict` and the file is unchanged; a save whose file disappeared is refused the same way; the tool layer maps the refusal to `conflict` (command 10).

### S5 — Unsupported feature is refused (FR-008)

- **Status**: automated evidence
- **Coverage**: `packages/document/tool-office/tests/document-tools.spec.ts` — "refuses a document carrying unsupported features before dispatch"; `packages/document/office-exchange/tests/provider-service.spec.ts` — "reports a macro feature on read and through detection"; `packages/document/office-exchange/tests/remote-edges.spec.ts` — "maps an unsupported feature refusal".
- **Observed**: a document reporting a macro is refused with `unsupported-feature` before any command is dispatched, so no write reaches the exchange; macro detection names the feature; the Remote layer maps the refusal (command 10).

### S6 — Rename rebinds identity (FR-015)

- **Status**: automated evidence
- **Coverage**: `packages/document/office-exchange/tests/remote.spec.ts` — "rebinds a moved document without changing its identity" and "marks a document missing once while keeping it open"; `packages/session/session-document/tests/document-identity.spec.ts` — the `rebindDocument` and `markDocumentMissing` cases; `packages/client/ui-office-editor/tests/changes.client.spec.ts` — "rebinds the document when its version reappears after its path disappears".
- **Observed**: a moved document keeps its document id and records the new path, a missing document stays open with its status, and the client rebinds when the same version reappears at an observed path (command 10).

### S7 — Unsaved work is not lost (FR-016)

- **Status**: automated evidence
- **Coverage**: `packages/client/ui-office-editor/tests/draft-recovery.spec.tsx` — "offers the draft an interrupted process left when the document reopens" and "warns instead of closing a dirty tab and keeps the edits"; `packages/session/session-document/tests/document-recovery.spec.ts` — "offers the draft an interrupted process left behind on relaunch" and its stale-version case.
- **Observed**: a dirty tab warns before closing and keeps its edits, and a draft left by an interrupted process is offered on relaunch, with drafts from another file version labelled stale (command 10).

### S8 — Write authority (FR-017)

- **Status**: automated evidence
- **Coverage**: `packages/document/tool-office/tests/authority.spec.ts` — "refuses cell and paragraph edits outside the workspace root", "refuses a relative path that escapes the workspace root", and "refuses an out-of-root write attempted through the Remote save seam"; `packages/document/office-exchange/tests/authority.spec.ts` — "rejects a target outside every root and names the target and roots" and the symlink escape case.
- **Observed**: out-of-root edits, relative escapes, and Remote-seam writes are refused with `out-of-scope`, naming the target and permitted roots, in the passing suite run (command 10).

### S9 — Local-first privacy (FR-013)

- **Status**: automated evidence
- **Coverage**: `packages/test-support/tool-testkit/tests/egress-documents.spec.ts` — "reads a workbook without reaching a network destination", "applies and saves a cell edit without reaching a network destination", and "applies and saves a paragraph edit without reaching a network destination"; `packages/document/office-exchange/tests/offline.spec.ts` — "configures no document server endpoint", "edits and saves a workbook with no external transport", and "edits and saves a document with no external transport".
- **Observed**: document reads and saves complete with outbound transports refusing connections and no document-server endpoint, so no document bytes leave the process; the egress spec passed all 3 tests (command 11) and the offline spec passed inside command 10.
- **Limit**: this is transport-level evidence. Telemetry default-off behavior was not separately exercised in this record; the message-level egress case lives in `packages/session/session-telemetry-otel/tests/egress.spec.ts`, outside the suites run here.

### S10 — Independent identity (final phase, FR-009)

- **Status**: automated evidence
- **Coverage**: `scripts/verify-forwork-identity.ts` via `corepack pnpm run verify-forwork-identity` (also a `hygiene` gate); `scripts/verify-license-notices.ts` via `corepack pnpm run verify-license-notices` (also a `hygiene` gate).
- **Observed**: the identity scan reports shipped surfaces carry ForWork identity under its documented exceptions; the notice scan checks 9 MIT base entries and 4 Apache-2.0 Univer entries in `THIRD_PARTY_NOTICES.md` (commands 12, 13, and 15).

### S11 — Signed desktop installers (final phase)

- **Status**: not run
- **Coverage**: none in this record.
- **Reason**: the scenario needs Windows and macOS signing identities and an installed, signed build. macOS signing is deferred and Windows signing still needs certificates, tracked by T051 in [tasks.md](./tasks.md).

## Not run

- **S3** — no configured model provider (`FORWORK_API_KEY` unset, no root `.env`).
- **S11** — no signing identities or signed installers; T051 owns the remaining work.
- **GUI observations** — no browser, desktop, or editor session was launched, so the quickstart's visual checks (editor refresh timing, recovery dialogs, installer launch) have no manual confirmation here.
- **`test:web`, `test:snapshot`, `test:e2e`, and the full `pnpm run test`** — excluded from this validation by instruction; T041 owns the remaining snapshot replay work.

## Gates after regeneration

`doc-sync` passed 43/43 on its final run (command 9). `test:docs` passed 21/21 (command 14) and `hygiene` passed 25/25 (command 15) after the regenerated catalogs and note repairs.
