# Implementation Plan: Office Agent Framework

**Branch**: `001-office-agent-framework` | **Date**: 2026-11-13 | **Spec**: [spec.md](./spec.md)

**Input**: Feature specification from `/specs/001-office-agent-framework/spec.md`

## Summary

Fork the MIT-licensed harness into **ForWork**, a Windows/macOS desktop product published under the `@forwork/*` npm scope with the `forwork-` artifact prefix and `forwork` CLI, and add a first-class document capability: an embedded Univer (Apache-2.0 core) editor for `.docx`/`.docm`/`.xlsx`/`.xlsm` (the macro-enabled pair widened by [002](../002-in-place-ooxml-save/plan.md)), a self-owned Exchange layer between OOXML files and Univer unit snapshots, agent tools that operate the same snapshot, durable attributable `document/*` session facts, and local-first safety (no upload, opt-in telemetry, write authority, credential storage, stable identity). Delivery is additive first; the identity-landing phase (scope, CLI, artifacts, brand, vendored re-scope, telemetry/default adapter, scan, signing) is scripted and gated per Constitution V.

## Technical Context

**Language/Version**: TypeScript (strict), Node >= 24.9, ESM

**Primary Dependencies**: Cordis (harness DI); Univer open-source presets (`@univerjs/presets`, `@univerjs/preset-sheets-core`, docs presets) and `@univerjs/core` (native snapshot types, also a Host dependency of `office-exchange`); ExcelJS for `.xlsx`; `fast-xml-parser` and `fflate` for the `.docx` OOXML path; `@deepseek-ai/libreoffice-kit` for legacy conversion; Electron for desktop packaging and OS secure storage

**Storage**: Local workspace files through the filesystem seam with version guards; the session log for document facts and change records; OS secure store for model credentials

**Testing**: Vitest (unit/behavior); repo session snapshots; `test:coverage` per-file gate; `test:gui` for client packages; `test:web` replay snapshots; an egress assertion for privacy

**Target Platform**: Windows and macOS Electron desktop (Web profile remains available in the fork)

**Project Type**: Desktop application plus workspace packages (monorepo)

**Performance Goals**: Editor open under 30 s for ordinary files; agent edit visible under 2 s; conversion off the renderer main thread (OOXML mapping runs in the Host's Node process via `office-exchange`, reached through the renderer command bridge)

**Constraints**: Offline editing and saving; no Linux document server or cloud account; bounded conversion memory and time (see `office-exchange` `Config`: `maxFileBytes`, `maxCells`, `maxConversionMs`); lossless-by-default saves; writes confined to permitted roots; agent tools dispatch correlated commands over local IPC (no document bytes off-device); runtime undo is the editor's native stack; user saves are explicit with periodic local drafts, while an accepted agent tool call persists with one immediate save

**Scale/Scope**: v1 edits `.docx`, `.docm`, `.xlsx`, and `.xlsm` (widened by [002](../002-in-place-ooxml-save/plan.md)); `.pptx` read-only; legacy conversion; one local user plus the agent; documents persist and restore per session

### Decision: Univer Sheet Feature Plugins (FR-022)

`.xlsx` conditional formatting, data validation, hyperlinks, and cell images require adding and vendoring `@univerjs/sheets-conditional-formatting`, `@univerjs/sheets-data-validation`, `@univerjs/sheets-hyper-link`, and `@univerjs/sheets-drawing` (Apache-2.0/MIT), mounted in the sheets preset and registered as `IResourceHook`s, so those features round-trip through the workbook snapshot and render in the live editor. Keeping the FR-008 refusal for these four features is the alternative; adding the plugins is chosen because FR-022 makes them preserved and editable, and the refusal narrows to worksheet features the sheet model cannot represent.

## Constitution Check

*GATE: Must pass before Phase 0 research. Re-check after Phase 1 design.*

| Principle | Gate | Status |
|---|---|---|
| I. Independent Identity | No shipped surface retains `deepseek`; `@forwork/*` scope, `forwork` CLI, `forwork-` artifacts; a scan gate asserts it | PASS for scope/CLI/brand; T051 signing pending |
| II. Offline, Local-First Desktop | Edit/save local; no upload; telemetry opt-in; no AGPL/cloud engine | PASS |
| III. Documents First-Class | `document/*` facts, projection, restore | PASS |
| IV. Lossless, Non-Destructive Editing | Unsupported-feature detection, refusal, attribution, undo, conflict | PASS |
| V. Incremental, Reversible Delivery | Additive packages/events before destructive rename | PASS |
| VI. Evidence Over Claims | Per-phase check ladder; egress/write-authority tests | PASS |
| VII. License and Dependency Hygiene | MIT + Apache-2.0 notices; license check on new dependencies | PASS |
| VIII. Selective Upstream Follow | Additive seams; cherry-pick by the decision table, no merge-forward claim | PASS |

No violations require a Complexity Tracking justification beyond the three already recorded.

**Post-design re-evaluation (after Phase 1)**: PASS. The design adds no AGPL/cloud/paid dependency; privacy is enforced by an egress test, writes by the filesystem seam and root confinement, credentials by the OS secure store, and upstream compatibility by additive seams. The stable identity registry is the only new component and is justified above.

## Project Structure

### Documentation (this feature)

```text
specs/001-office-agent-framework/
├── spec.md
├── plan.md
├── research.md
├── data-model.md
├── quickstart.md
├── contracts/
│   ├── office-exchange.contract.md
│   ├── document-events.contract.md
│   ├── command-bridge.contract.md
│   └── tool-office.contract.md
├── tasks.md
└── checklists/
    └── requirements.md
```

### Source Code (repository root)

```text
packages/document/
├── office-to-pdf/       # existing: LibreOffice -> PDF (NOT used for editable legacy conversion)
├── office-exchange/     # this feature's seam: OOXML <-> unit snapshot, feature detection (built by T002/T008–T012)
└── tool-office/         # NEW: model-facing structured document tools (agent)
packages/client/
└── ui-office-editor/    # NEW: embedded Univer editor tab + document panel
packages/core/session/                 # CHANGED: document/* SessionEventMap
packages/session/session-document/     # NEW: identity, change records, document facts, projection, restore, recovery
packages/bundle/base/                  # CHANGED: mount office capability as non-removable core
packages/credentials/credentials/      # CHANGED (consumer only): OS secure store
packages/test-support/tool-testkit/    # CHANGED: egress assertion helpers
apps/desktop/            # CHANGED: branding + signing identity (final phase)
```

**Structure Decision**: Follow the existing package-group layout. The Exchange layer is a Host service seam; the editor is a Client plugin package registered into the existing right-sidebar tab system and owns the single live unit per open document; agent tools are a Host tool package that dispatch correlated commands over the gateway into the editor's live unit (falling back to `document-not-open` when no renderer is attached); session changes extend the core event map; privacy, write authority, and credentials reuse existing seams rather than new subsystems. Legacy conversion uses `@deepseek-ai/libreoffice-kit` directly; `office-to-pdf` only renders PDF and is not the conversion path.

## Complexity Tracking

| Violation | Why Needed | Simpler Alternative Rejected Because |
|-----------|------------|-------------------------------------|
| Self-built Exchange layer | Open-source Univer excludes OOXML import/export | Buying the commercial Server SDK adds a paid dependency and contradicts the offline, independent goal |
| New session event family | Documents must survive reload and be model-visible | Tool-argument-only documents cannot be cross-turn referenced or reconstructed from the log |
| Stable document identity registry | Files move/rename; path identity breaks restore | Path-keyed identity loses the document on rename and cannot rebind safely |

## Next Step: Brand Artwork Completion and Visual Evidence (2026-11-17)

**Scope**: Finish the T120 design half that is still open after the door mark, app icons, wordmark, favicons, welcome art, installer art, and component adoption landed. This step removes the last DeepSeek-derived raster, makes the desktop SVG derivation reproducible, adds the asset gate, and produces committed end-to-end screenshots of the shipped surfaces.

**Inputs**: [research R21-R24](./research.md#r21-home-animation-removal-accepted-2026-11-17), [brand asset and evidence model](./data-model.md#brand-asset-and-evidence), [asset contract](./contracts/brand-artwork.contract.md), [quickstart S14-S19](./quickstart.md#scenario-s14--brand-assets-derive-and-pass-the-gate).

### Technical Context

**Language/Version**: TypeScript (strict), Node >= 24.9, ESM

**Primary Dependencies**: Playwright Chromium for browser screenshots; sharp for deterministic raster composition; VitePress build output for website evidence; existing `launchWebScaffold` and desktop welcome renderer; no new runtime dependency

**Storage**: Repository asset masters under `assets/brand/`; committed review screenshots under each owner's `tests/expected/brand/`; CI working copies under the gitignored `.artifacts/brand/`; `scripts/brand-assets.manifest.json` records the committed screenshot hashes and dimensions

**Testing**: Focused Vitest unit tests, Web scaffold e2e with Playwright screenshots, desktop welcome renderer e2e, website static-server e2e, `verify-brand-assets` acceptance and rejection specs, existing `test:docs` and `test:web` lanes

**Target Platform**: Web (Chromium), Desktop welcome renderer (Chromium with the preload bridge mocked), website build; Windows/macOS installer pages are asset-verified here and captured in the Windows packaging lane

**Project Type**: Desktop application plus workspace packages (monorepo)

**Performance Goals**: `verify-brand-assets` under 2 s on the asset corpus; screenshot suites under 60 s each after the build; no product-path cost

**Constraints**: Screenshots are deterministic in viewport, locale, theme, and motion but do not gate on cross-platform pixel equality; every committed screenshot records the asset hashes it depicts; the home animation removal is a deliberate user-visible behavior change

**Scale/Scope**: About 40 shipped brand assets across seven roots; one removed feature (`conversation.hero.frame` whale occupant); four screenshot suites plus one GIF; one gate and one manifest

### Constitution Check

| Principle | Gate | Status |
|---|---|---|
| I. Independent Identity | Remove the last DeepSeek-derived whale atlas; manifest and fingerprints reject reintroduction | PASS after R21 |
| II. Offline, Local-First Desktop | Screenshots and the gate run locally; no upload, model call, or telemetry change | PASS |
| III. Documents First-Class | No document behavior changes | N/A |
| IV. Lossless, Non-Destructive Editing | No file-write or undo behavior changes | N/A |
| V. Incremental, Reversible Delivery | Asset and component work already landed; this step removes one isolated frame layer and adds checks | PASS |
| VI. Evidence Over Claims | Committed screenshots, asset hashes, live manifest, rejection spec, and gate output | PASS |
| VII. License and Dependency Hygiene | Removes an upstream-derived raster; adds no dependency | PASS |
| VIII. Selective Upstream Follow | No upstream sync or seam rewrite | N/A |

No Complexity Tracking exception is required.

**Post-design re-evaluation**: PASS. The plan removes rather than adds a shipped decoration, keeps the slot contract available, and makes every remaining brand asset derivable and fingerprinted.

### Phase 0 Research Output

- R21: Remove the click-triggered whale frame layer instead of generating a replacement atlas.
- R22: Commit reviewed screenshots and assert DOM, geometry, dimensions, and hashes; do not use cross-platform pixel diff as a blocking gate.
- R23: One manifest owns shipped assets, screenshot inventory, and retired fingerprints; the gate runs in hygiene and doc-adjacent lanes.
- R24: A derivation script owns the desktop SVG variants so `icon.svg` remains the square crop source and platform sources stay reproducible.

### Phase 1 Design Output

- [Brand Asset and Evidence model](./data-model.md#brand-asset-and-evidence): manifest rows, evidence records, and validation rules.
- [Brand artwork contract](./contracts/brand-artwork.contract.md): asset paths, manifest and screenshot schemas, fingerprint rejection, and the frame-layer removal promise.
- [Quickstart S14-S19](./quickstart.md#scenario-s14--brand-assets-derive-and-pass-the-gate): runnable asset, screenshot, removal, and rejection scenarios.

### E2E Visual Evidence Matrix

| Surface | Test | Viewport / theme | Committed artifact |
|---|---|---|---|
| Web blank home, sidebar expanded | `apps/web/tests/brand-artwork.e2e.ts` | 1440×900, light | `apps/web/tests/expected/brand/home-desktop-light.png` |
| Web blank home, sidebar expanded | same | 1440×900, dark | `apps/web/tests/expected/brand/home-desktop-dark.png` |
| Web home, door leaf hover open | same | 1440×900, light, motion allowed | `apps/web/tests/expected/brand/home-hover.png` |
| Web home, mobile drawer | same | 390×844, light | `apps/web/tests/expected/brand/home-mobile.png` |
| Web sidebar collapsed rail | same | 1440×900, light | `apps/web/tests/expected/brand/rail-collapsed.png` |
| Web install icon sizes | same | 1024×320 contact sheet | `apps/web/tests/expected/brand/icon-sheet.png` |
| Desktop welcome window | `apps/web/tests/brand-welcome.e2e.ts` | 600×700 CSS, 2× image, light and dark | `apps/desktop/tests/expected/brand/welcome-{light,dark}.png` |
| Website nav lockup | `apps/web/tests/brand-website.e2e.ts` | 1440×240 and 390×720 | `website/tests/expected/brand/nav-{desktop,mobile}.png` |
| Installer and uninstaller art | `apps/web/tests/brand-installer.e2e.ts` | 600×196 and 164×314 | `apps/desktop/tests/expected/brand/installer-{light,dark}.png`, `uninstaller.png` |
| Home hover demo GIF | `record-browser-gif` workflow | 960×600, light | PR attachment `brand-home-hover.gif` |

Each screenshot records locale `en-US`, `deviceScaleFactor: 1`, and the manifest hash of every asset it depicts. The e2e suite always writes fresh screenshots to `.artifacts/brand/`; `UPDATE_BRAND_SCREENSHOTS=1` refreshes the committed review set, and the normal lane asserts that the committed set exists, has the recorded dimensions, and matches the manifest hashes. Pixel equality against the live render is deliberately not asserted because font rasterization differs by platform.

### Project Structure

```text
assets/brand/                              # masters, bilingual brief, manifest source
scripts/
├── render-brand-assets.ts                 # desktop SVG variants from the tile master
├── verify-brand-assets.ts                 # manifest, dimensions, fingerprints, evidence inventory
└── verify-brand-assets.spec.ts            # acceptance plus rejection fixtures
vitest.brand.config.ts                     # includes the four evidence suites
apps/web/tests/
├── brand-artwork.e2e.ts                   # Web screenshots and door geometry
└── expected/brand/                        # committed review screenshots
apps/desktop/tests/
├── brand-welcome.e2e.ts                   # welcome renderer screenshots
├── brand-installer.e2e.ts                 # installer/uninstaller art evidence
└── expected/brand/
website/tests/
├── brand-website.e2e.ts                   # built-site nav screenshots
└── expected/brand/
packages/client/ui-brand-official/         # whale frame layer removed; sidebar slots retained
```

**Structure Decision**: Keep screenshots owner-local because they are reviewed expected output, not session-driven snapshots. Keep the gate and derivation script at the repository scripts layer so every app shares one manifest and one rejection corpus. Keep the `conversation.hero.frame` slot declared; only the official whale occupant is removed.

### Risks and Mitigations

- Removing the whale changes a decorative behavior: land it as its own commit, update the package README and slot catalog, and keep the slot available for future occupants.
- Screenshots can vary by font and OS: assert DOM and dimensions, commit reviewed images, and upload fresh `.artifacts/brand/` from CI rather than comparing pixels.
- Gate hashes churn on legitimate asset edits: the failure message names the asset and prints the command to refresh the manifest; review includes the regenerated screenshots.
- Installer pages are platform-specific: asset-level evidence is committed here; the Windows packaging lane captures the real NSIS pages when it runs.
- The welcome renderer needs the preload bridge: the e2e injects a typed mock before load and asserts the brand image is visible, not the real IPC behavior already covered by host tests.

### Validation Ladder

```sh
corepack pnpm run build:artifacts
corepack pnpm run build:check
corepack pnpm run lint
corepack pnpm run verify-brand-assets
corepack pnpm run verify-forwork-identity
corepack pnpm run test:brand:built
corepack pnpm run test:docs
corepack pnpm run website:build
```

### Out of Scope

- A new animated home decoration; the minimal brand uses the static door mark and its hover leaf.
- Cross-platform pixel-diff baselines.
- Real NSIS page capture outside the Windows packaging lane.
- Provider-backed model conversations in the screenshot suites; the blank-home surface is the evidence target.
