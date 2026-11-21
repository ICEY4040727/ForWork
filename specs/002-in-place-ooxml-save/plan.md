# Implementation Plan: In-Place OOXML Save

**Branch**: `002-in-place-ooxml-save` | **Date**: 2026-11-18 | **Spec**: [spec.md](./spec.md)

**Input**: Feature specification from `/specs/002-in-place-ooxml-save/spec.md`

## Summary

Replace the whole-package regeneration in `@forwork/office-exchange` with a two-layer write path. A Host-only package layer indexes the source ZIP, copies every untouched local record byte-for-byte, and rebuilds only the central directory; an XML span layer splices only the edited unit (paragraph, table cell, worksheet cell) into the parts that change. The save operation uses the source bytes read under the existing version guard, and returns a durable disposition fact. `.docm`/`.xlsm` become editable under the same rules, and the whole-file unsupported-feature refusal is replaced by a construct-intersection refusal.

The fallback writer stays for snapshots and drafts that carry no source identity, so old drafts keep working and the new path ships as an additive module.

## Technical Context

**Language/Version**: TypeScript (strict, ESM), Node >= 24.9; Host process only for the new writer

**Primary Dependencies**: existing `fflate` 0.8.3 (read helper and tests), `fast-xml-parser` 5.11.1 (model mapping), `fast-xml-validator` 1.4.2 (well-formedness precondition), `exceljs` 4.4.0 (`.xlsx` read mapping), `@univerjs/core` 1.4.0 (types), and built-in `node:zlib` (`crc32`, `deflateRawSync`, `inflateRawSync`). No new dependency.

**Storage**: Local filesystem through the `Fs` seam; the visible file is published by `writeBytes(..., { kind: 'replaceIfVersion' })`.

**Testing**: Vitest unit/behavior specs; a new OPC/byte-diff harness under `packages/document/office-exchange/tests/`; LibreOffice kit load smoke; Web keyless snapshot; per-file 100% coverage on `packages/*/*/src`.

**Target Platform**: Electron desktop for Windows and macOS; the writer is Host-side and must not enter a Client bundle.

**Project Type**: pnpm monorepo, desktop application and capability-seam packages.

**Performance Goals**: A 5 MB / 10,000-cell package opens, edits one region, and saves in under 30 s on the reference desktop; preservation cost is linear in package entries plus the edited parts, and no part is re-serialized when it has no edited span.

**Constraints**: Offline only; no new dependencies; existing `maxFileBytes`/`maxCells`/`maxConversionMs` and entry-count bounds, plus two new validated Config fields for the anchor count and anchor payload byte budgets; Host code must not import `node:zlib` from a Client package; `SESSION_FORMAT_VERSION` stays 4 with an optional event property; the constitution is amended to 3.2.0 for the format set and write mechanism.

**Scale/Scope**: `packages/document/office-exchange` (new `src/package/*`, provider rewrite), `packages/core/session` + `packages/session/session-document` (save fact), `packages/client/ui-office-editor` (format and refusal surface), `packages/document/tool-office`, `apps/web` tests, docs and persistence records. Corpus floor: 30 fixtures.

## Constitution Check

*GATE: Must pass before Phase 0 research. Re-check after Phase 1 design.*

| Principle | Gate | Status |
| --- | --- | --- |
| I. Independent Identity | New modules live under `@forwork/office-exchange`; no brand, scope, or artifact change | PASS |
| II. Offline, Local-First Desktop | `node:zlib` and the local filesystem only; LibreOffice smoke runs locally; no network path added | PASS |
| III. Documents Are First-Class Session Entities | The save disposition is a durable `document/save` fact; nothing model-visible is added without a log entry | PASS |
| IV. Lossless, Non-Destructive Editing | This feature is the enforcement: untouched local records are copied raw; splices are unit-scoped; coupled constructs refuse loudly | PASS |
| V. Incremental, Reversible Delivery | New `src/package/*` layer and a feature-gated splice path; the existing regeneration writer remains as the fallback for snapshots without source identity | PASS |
| VI. Evidence Over Claims | Corpus with byte diffs, OPC validator, LibreOffice load smoke, manual Word/Excel matrix, per-file coverage | PASS (planned evidence, not yet executed) |
| VII. License and Dependency Hygiene | No new dependency; `node:zlib` is built in; ExcelJS and Univer licenses unchanged | PASS |
| VIII. Selective Upstream Follow | Additive fork-local modules in `office-exchange`; no diverged upstream file rewritten | PASS |

**Constitution amendment recorded** (3.1.0 → 3.2.0, 2026-11-18; see Complexity Tracking):

- The technology-constraint sentence "v1 edits `.docx` and `.xlsx`" is widened to `.docx/.docm/.xlsx/.xlsm`.
- The conversion-stack sentence now records the Host-only package writer over `node:zlib`, with ExcelJS as the `.xlsx` read mapper, `fast-xml-parser` as the `.docx` read mapper, and the legacy regeneration writer retained for snapshots without source identity.

## Project Structure

### Documentation (this feature)

```text
specs/002-in-place-ooxml-save/
├── spec.md              # Feature specification
├── plan.md              # This file
├── research.md          # Phase 0 decisions R1-R16
├── data-model.md        # Source package, anchors, splice plan, disposition
├── quickstart.md        # Runnable validation guide
├── contracts/
│   ├── office-exchange-write.contract.md
│   └── document-save-disposition.contract.md
└── checklists/
    └── requirements.md
```

### Source Code (repository root)

```text
packages/document/office-exchange/
├── src/
│   ├── package/                    # NEW: Host-only ZIP/OPC layer
│   │   ├── zip-format.ts           # signatures, flags, methods, DOS time, bounds
│   │   ├── zip-reader.ts           # entry index, raw local/central records
│   │   ├── zip-writer.ts           # raw copy + central directory rebuild
│   │   └── ooxml-package.ts        # part map, dirty set, no-op fast path
│   ├── source/                     # NEW: anchors, scanner, splice plan
│   │   ├── source-state.ts         # sidecar validation and digests
│   │   ├── xml-spans.ts            # namespace/comment/CDATA-aware byte scanner
│   │   ├── unit-digest.ts          # canonical model digest per unit
│   │   └── splice.ts               # edit script -> span operations
│   ├── provider/
│   │   ├── docx.ts                 # splice writer + legacy fallback
│   │   ├── xlsx.ts                 # cell-span writer + ExcelJS read path
│   │   └── index.ts                # write(snapshot, format, source)
│   ├── detect.ts                   # preserved vs coupled classification
│   ├── remote.ts                   # source read, version ordering, disposition
│   └── types.ts                    # formats, source state, disposition
└── tests/
    ├── ooxml-package.ts            # OPC validator + byte-diff report helper
    ├── ooxml-preservation.spec.ts  # corpus and identity matrix
    ├── ooxml-engine.spec.ts        # LibreOffice load smoke
    └── fixtures/                   # extended generator and binaries

packages/core/session/src/document-events.ts        # optional disposition field
packages/session/session-document/src/document-facts.ts
packages/client/ui-office-editor/src/client/useOfficeDocument.ts
packages/document/tool-office/src/edit.ts
apps/web/tests/document-preview.e2e.ts              # save click and read-back
```

**Structure Decision**: The new code is additive under the existing `office-exchange` package: a generic ZIP/OPC layer (`src/package/`), a source-span layer (`src/source/`), and provider integration. Session, client, and tool changes are thin and follow existing seams. No new package is created.

## Phase 0: Research Output

See [research.md](./research.md). Decisions R1-R16 cover the container layer, span scanner, anchor transport, changed-unit detection, structural edits, auxiliary-part splices, `.xlsx` cell representation, source read ordering, legacy fallback, refusal classification, the session fact, format widening, and the validation ladder.

Key conclusions:

- No installed ZIP library can copy a compressed entry raw; a small `node:zlib`-backed reader/writer is required (R1).
- `fast-xml-parser`'s `captureMetaData` normalizes CRLF before recording offsets, so it cannot back a byte guarantee; a small namespace-aware byte scanner is required (R3).
- Anchors must live in `data.officeExchange.source`; the snapshot wrapper, `resources`, `body.payloads`, and per-paragraph fields all lose them (R5).
- The source package must reach the writer; the snapshot must not carry it (R10).

## Phase 1: Design Output

- [data-model.md](./data-model.md): `SourcePackage`, `OfficeSourceState`, `UnitAnchor`, `SplicePlan`, `PartDisposition`, `SaveDisposition`, `ConstructIntersection`, `PreservationFailure`; validation rules, bounds, and state transitions.
- [contracts/office-exchange-write.contract.md](./contracts/office-exchange-write.contract.md): the new write operation, its preconditions, postconditions, errors, and invariants.
- [contracts/document-save-disposition.contract.md](./contracts/document-save-disposition.contract.md): the optional save-fact property and its compatibility rule.
- [quickstart.md](./quickstart.md): fixture regeneration, preservation matrix, refusal postcondition, LibreOffice smoke, focused coverage, Web save snapshot, manual Office matrix.

## Constitution Check (post-design)

| Principle | Re-check |
| --- | --- |
| I. Independent Identity | PASS — all new code stays in `@forwork/office-exchange`; no shipped surface renamed |
| II. Offline | PASS — `node:zlib` + local files; the engine smoke is local and skippable |
| III. Session entities | PASS — the disposition is an optional `document/save` field, not a new untracked surface |
| IV. Lossless | PASS — the write contract makes byte identity and named refusal postconditions |
| V. Incremental | PASS — the legacy writer remains for source-less snapshots and old drafts |
| VI. Evidence | PASS — the corpus is the acceptance basis; no claim without a fixture |
| VII. License | PASS — no new dependency |
| VIII. Upstream | PASS — additive fork-local modules |

**Technology Constraints**: PASS — the 3.2.0 amendment records the widened format set (`.docx/.docm/.xlsx/.xlsm`) and the package-level write stack; no other constraint changes.

## Complexity Tracking

| Violation | Why Needed | Simpler Alternative Rejected Because |
| --- | --- | --- |
| `.docm`/`.xlsm` added to the editable set (constitution `Technology Constraints`) | The refusal case users actually have is a macro-enabled file; preserving the VBA part is free once entries are raw-copied | Keeping `.docx`/`.xlsx` only leaves the headline macro case refused and makes US2's acceptance scenario artificial |
| `node:zlib` ZIP layer instead of fflate for writing | fflate cannot reuse a compressed stream, cannot preserve local/central metadata, and cannot write an archive comment; all are required for byte identity | Writing with fflate makes a no-op save produce different bytes, so SC-001/SC-002 cannot pass |
| Hand-written cell XML for an edited `.xlsx` cell | ExcelJS serializes a whole workbook, not a single cell span, and cannot be asked to preserve the rest of the package | Re-serializing through ExcelJS defeats the feature and drops unknown parts |
| New `document/save` optional property | FR-017 requires a durable, restart-readable disposition | A tool result only is not durable; a log file is not reconstructable from the session |

The first two rows are recorded in the 3.2.0 amendment landed with this change, per Governance.

## Risks and Mitigations

| Risk | Mitigation |
| --- | --- |
| ZIP writer mis-copies a descriptor or offset and produces an unreadable package | Refuse ambiguous descriptors and ZIP64; validate the output through the OPC validator and LibreOffice smoke; add negative fixtures |
| Scanner selects a wrong span (comments, CDATA, prefix rebinding, `mc:AlternateContent`) | Fail closed on any ambiguity; never splice at a guessed offset; cross-check spans against parser metadata on LF-only fixtures |
| Editor round trip drops anchors | Anchors live in `data.officeExchange.source`; add a real Univer round-trip test (read → wire → mount → edit → save → write) and a draft round trip |
| Structural edits (insert/delete/resize) cannot be expressed as one span | Sequence diff produces replace/insert/delete spans; table row/column resize refuses with the construct named |
| Sharing a string shared by several cells | Convert the edited cell to `t="inlineStr"`; refuse rich-text/shared/array formula cells as coupled |
| Anchor payload too large for 2M cells | Bounded anchor budget; above it, re-map the source at save and diff, or refuse loudly; never silently drop anchors |
| A second save uses stale anchors | Compare the source package digest; refuse on mismatch before writing |
| Existing tests encode the old whole-file refusal | Update the tests with the behavior change and document why in the PR |

## Validation Ladder

1. `pnpm exec vitest run packages/document/office-exchange/tests/ooxml-preservation.spec.ts` — corpus identity and splice matrix.
2. `pnpm exec vitest run packages/document/office-exchange/tests/write-postcondition.spec.ts` — refusal and fact postconditions.
3. `pnpm exec vitest run packages/document/office-exchange/tests/ooxml-engine.spec.ts` — LibreOffice load smoke (skips when unavailable).
4. Focused coverage: `pnpm exec vitest run <specs> --coverage --coverage.include='packages/document/office-exchange/src/**/*.ts'`.
5. `pnpm run typecheck`, `pnpm run lint`, `pnpm run build:artifacts`, `pnpm run build:check`.
6. `pnpm run test:coverage` (CI gate), `pnpm run doc-sync`, `pnpm run hygiene` when manifests/exports/docs change.
7. Manual Word/Excel open matrix on the fixed release subset; Web save snapshot refresh.

## Out of Scope

- Editing macros, charts, pivot tables, comments, or embedded objects.
- Legacy `.doc/.xls/.ppt` binary write-back; `.pptx` editing.
- ZIP64, encrypted packages, split archives, and unsigned non-OOXML payloads.
- Structural table row/column resize beyond refusal.
- Any cloud or network path.
