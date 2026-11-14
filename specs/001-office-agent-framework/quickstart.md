# Quickstart Validation: Office Agent Framework

**Purpose**: Prove the feature end to end without reading implementation code.

## Prerequisites

- Node >= 24.9; pnpm `11.28.5` (activate with `corepack pnpm --version`).
- A workspace checkout; dependencies installed.
- Fixture files: at least one `.xlsx` with formulas/styles, one `.docx` with headings/tables/images, one `.doc` or `.xls` legacy file, and one file containing an unsupported feature (for example a macro).
- For US2, a configured model provider; without one, use the headless tool path only.

## Setup

```sh
corepack pnpm install --frozen-lockfile
corepack pnpm exec tsc -b packages/document/office-exchange
```

## Scenario S1 — Edit and save `.xlsx` losslessly (US1)

1. Open the fixture workbook, change one cell, save, reopen.
2. Compare against the original with `document_read`.

**Expected**: only the edited cell differs; formulas, styles, merged cells, and images are unchanged. Untouched bytes/features are preserved.

## Scenario S2 — Edit and save `.docx` losslessly (US1)

1. Open the fixture document, replace one paragraph's text, save, reopen.

**Expected**: the target paragraph changes; all other paragraphs, tables, and images are preserved.

## Scenario S3 — Agent edit converges (US2)

1. Open a document; ask the agent to set a cell or replace a paragraph.
2. Observe the open editor and the saved file.

**Expected**: the editor shows the change within 2 seconds without a manual reload; the file contains it; one Change Record with `actor = agent` is logged.

## Scenario S4 — Conflict is rejected (US2/FR-005)

1. Make an external change to the file while it is open.
2. Have the agent save using the stale version.

**Expected**: the write is rejected with `conflict`; the external change is not overwritten.

## Scenario S5 — Unsupported feature is refused (FR-008)

1. Open the macro-bearing fixture and attempt a save.

**Expected**: the save is refused with `unsupported-feature` naming the feature; the file is not rewritten.

## Scenario S6 — Rename rebinds identity (FR-015)

1. Open a document, then rename or move the file outside the app.
2. Trigger a reload/restore.

**Expected**: the document keeps its identity and rebinds to the new path; if the file is gone, it is reported missing and stays open.

## Scenario S7 — Unsaved work is not lost (FR-016)

1. Edit a document without saving.
2. Kill the app process, then relaunch.

**Expected**: the app warns before an intentional close, and after an abnormal exit a recoverable draft is offered; no silent loss.

## Scenario S8 — Write authority (FR-017)

1. Instruct the agent to write outside the permitted roots.

**Expected**: the write is refused with `out-of-scope`; no file is created or modified.

## Scenario S9 — Local-first privacy (FR-013)

1. Capture network egress while opening, editing, and saving with the model call disabled.

**Expected**: no document bytes or edit history leave the device; telemetry sends nothing by default.

## Scenario S10 — Independent identity (final phase, FR-009)

1. Build and scan the packaged product for brand/telemetry references.

**Expected**: zero DeepSeek-referencing shipped surfaces.

## Scenario S11 — Signed desktop installers (final phase)

1. Build Windows and macOS installers with the product's signing identity and install.

**Expected**: both launch under the product's own publisher identity with no unsigned-publisher warning.

## Automated checks

```sh
corepack pnpm exec vitest run packages/document/office-exchange/tests
corepack pnpm run test:gui
corepack pnpm run build:check
corepack pnpm run doc-sync
```

**Expected**: all pass. Session-format changes also require `test:snapshot` and the TypeScript/Python SDK expected-output updates.

## References

- Entity and state rules: [data-model.md](./data-model.md)
- Service contract: [contracts/office-exchange.contract.md](./contracts/office-exchange.contract.md)
- Session contract: [contracts/document-events.contract.md](./contracts/document-events.contract.md)
- Tool contract: [contracts/tool-office.contract.md](./contracts/tool-office.contract.md)
