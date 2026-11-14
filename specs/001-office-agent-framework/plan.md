# Implementation Plan: Office Agent Framework

**Branch**: `001-office-agent-framework` | **Date**: 2026-11-13 | **Spec**: [spec.md](./spec.md)

**Input**: Feature specification from `/specs/001-office-agent-framework/spec.md`

## Summary

Fork the MIT-licensed harness into an independently branded Windows/macOS desktop product and add a first-class document capability: an embedded Univer (Apache-2.0 core) editor for `.docx`/`.xlsx`, a self-owned Exchange layer between OOXML files and Univer unit snapshots, agent tools that operate the same snapshot, durable attributable `document/*` session facts, and local-first safety (no upload, opt-in telemetry, write authority, credential storage, stable identity). Delivery is additive first; the scope/brand/signature rename is the final gated phase.

## Technical Context

**Language/Version**: TypeScript (strict), Node >= 24.9, ESM

**Primary Dependencies**: Cordis (harness DI); Univer open-source presets (`@univerjs/presets`, `@univerjs/preset-sheets-core`, docs presets) and `@univerjs/core` (native snapshot types, also a Host dependency of `office-exchange`); ExcelJS; SheetJS CE; a docx library; `@deepseek-ai/libreoffice-kit` for legacy conversion; Electron for desktop packaging and OS secure storage

**Storage**: Local workspace files through the filesystem seam with version guards; the session log for document facts and change records; OS secure store for model credentials

**Testing**: Vitest (unit/behavior); repo session snapshots; `test:coverage` per-file gate; `test:gui` for client packages; `test:web` replay snapshots; an egress assertion for privacy

**Target Platform**: Windows and macOS Electron desktop (Web profile remains available in the fork)

**Project Type**: Desktop application plus workspace packages (monorepo)

**Performance Goals**: Editor open under 30 s for ordinary files; agent edit visible under 2 s; conversion off the main thread

**Constraints**: Offline editing and saving; no Linux document server or cloud account; bounded conversion memory and time; lossless-by-default saves; writes confined to permitted roots; agent tools dispatch correlated commands over local IPC (no document bytes off-device); runtime undo is the editor's native stack; user saves are explicit with periodic local drafts, while an accepted agent tool call persists with one immediate save

**Scale/Scope**: v1 edits `.docx` and `.xlsx`; `.pptx` read-only; legacy conversion; one local user plus the agent; documents persist and restore per session

## Constitution Check

*GATE: Must pass before Phase 0 research. Re-check after Phase 1 design.*

| Principle | Gate | Status |
|---|---|---|
| I. Independent Identity | No shipped surface retains `deepseek`; a scan gate asserts it | PASS with planned final rename phase |
| II. Offline, Local-First Desktop | Edit/save local; no upload; telemetry opt-in; no AGPL/cloud engine | PASS |
| III. Documents First-Class | `document/*` facts, projection, restore | PASS |
| IV. Lossless, Non-Destructive Editing | Unsupported-feature detection, refusal, attribution, undo, conflict | PASS |
| V. Incremental, Reversible Delivery | Additive packages/events before destructive rename | PASS |
| VI. Evidence Over Claims | Per-phase check ladder; egress/write-authority tests | PASS |
| VII. License and Dependency Hygiene | MIT + Apache-2.0 notices; license check on new dependencies | PASS |
| VIII. Fork Stewardship | Additive seams; merge-forward with recovery copy and intent | PASS |

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
│   └── tool-office.contract.md
├── tasks.md
└── checklists/
    └── requirements.md
```

### Source Code (repository root)

```text
packages/document/
├── office-to-pdf/       # existing: LibreOffice -> PDF (NOT used for editable legacy conversion)
├── office-exchange/     # EXISTING (seam landed): OOXML <-> unit snapshot, feature detection
└── tool-office/         # NEW: model-facing structured document tools (agent)
packages/client/
└── ui-office-editor/    # NEW: embedded Univer editor tab + document panel
packages/core/session/                 # CHANGED: document/* SessionEventMap
packages/session/session-document/     # NEW: identity, change records, projection, restore, undo
packages/session/session-projection/   # CHANGED: fold document facts
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
