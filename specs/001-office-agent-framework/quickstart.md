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
- Command bridge contract: [contracts/command-bridge.contract.md](./contracts/command-bridge.contract.md)
- Tool contract: [contracts/tool-office.contract.md](./contracts/tool-office.contract.md)
- Upstream sync policy: [docs/development.md#upstream-sync](../../docs/development.md#upstream-sync)

## Scenario S14 — Brand assets derive and pass the gate

1. Run `corepack pnpm run verify-brand-assets`.
2. Change one SVG in `assets/brand/` intentionally and rerun.

**Expected**: the unchanged tree passes; the changed tree fails naming the stale manifest row, and the printed refresh command updates the manifest after review.

## Scenario S15 — Finished Web screenshots

1. Build the Web app and run `corepack pnpm run test:brand:built`.
2. Inspect `apps/web/tests/expected/brand/` and the fresh copies under `.artifacts/brand/`.
3. After an intentional visual change, refresh with `UPDATE_BRAND_SCREENSHOTS=1 corepack pnpm run test:brand:built`, then record the hashes with `corepack pnpm run verify-brand-assets --write`.

**Expected**: desktop light, desktop dark, hover, mobile, rail, and icon-sheet images exist at their recorded dimensions; the door is visible in the sidebar and hero; the hover image shows the leaf rotated; the manifest hashes match the assets.

## Scenario S16 — Desktop welcome screenshot

1. Run `corepack pnpm run test:brand:built`.
2. Inspect `apps/desktop/tests/expected/brand/welcome-{light,dark}.png`.

**Expected**: the built welcome renderer shows the ForWork door lockup at 472×40 in both themes; the mocked preload exposes only the bridge the page reads.

## Scenario S17 — Website navigation screenshot

1. Run `corepack pnpm run website:build`.
2. Run `corepack pnpm run test:brand:built`.
3. Inspect `website/tests/expected/brand/nav-{desktop,mobile}.png`.

**Expected**: the nav shows the door lockup and the current release tag; the favicon tile renders; no DeepSeek whale is visible.

## Scenario S18 — Retired artwork is rejected

1. Run the rejection spec, which builds a temp tree with a manifest row and an SVG containing a retired whale path fragment.
2. Run `corepack pnpm run verify-brand-assets` on the real tree.

**Expected**: the rejection case exits 1 naming `deepseek-whale-mark` and the temp path; the real tree passes. The test exercises the fingerprint branch rather than only the unreferenced-file branch.

## Scenario S19 — Home no longer plays the whale

1. Open the built Web app on the blank home.
2. Click empty space in the conversation frame.

**Expected**: no whale canvas appears, no `[data-home-whale]` node exists, and the door mark remains in the sidebar and hero. The `conversation.hero.frame` slot is still declared in the slot catalog.

## Automated checks

```sh
corepack pnpm run build:artifacts
corepack pnpm run build:check
corepack pnpm run verify-brand-assets
corepack pnpm run verify-forwork-identity
corepack pnpm run test:brand:built
corepack pnpm run test:docs
corepack pnpm run website:build
```

**Expected**: all pass, and every screenshot named in the evidence matrix exists with a matching manifest hash.