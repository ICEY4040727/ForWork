# Research: Office Agent Framework

**Input**: [spec.md](./spec.md) and [plan.md](./plan.md)
**Date**: 2026-11-13

This file records the technical decisions that resolve the plan's open questions. Each entry states the decision, why, and what was rejected.

## R1. Editor core and license boundary

- **Decision**: Use Univer's open-source Web SDK (Apache-2.0) with the Sheets and Docs presets, embedded in React. Use no commercial Univer module in v1.
- **Rationale**: It is the only evaluated core that is browser-native (no server), Apache-2.0, embeddable without an iframe, and exposes a programmatic facade plus headless content operations.
- **Alternatives**: OnlyOffice (full fidelity but AGPL and a Linux document server); SpreadJS (paid, spreadsheets only); AFFiNE/BlockSuite (not OOXML); Luckysheet/FortuneSheet (spreadsheet only, superseded).

## R2. XLSX to and from the editable model

- **Decision**: Map `.xlsx` with ExcelJS (already pinned and patched in this repo) and use SheetJS CE for `.xls`. Preserve untouched parts by editing the parsed workbook structures rather than regenerating the file.
- **Rationale**: ExcelJS is already vetted and patched for part-relationship handling; SheetJS covers legacy `.xls`.
- **Alternatives**: Rebuild the workbook from scratch (loses styles, drawings, comments); a commercial converter (paid, violates offline independence).

## R3. DOCX to and from the editable model (highest risk)

- **Decision**: Parse the OOXML text-body into the document snapshot, preserving run structure, and write back by targeted replacement of changed runs; detect and refuse features the mapping cannot represent (macros, tracked changes, complex shapes, headers/footers in v1).
- **Rationale**: Run-preserving replacement keeps untouched formatting. Refusing unrepresentable features satisfies the lossless principle.
- **Alternatives**: A wholesale docx rebuild (loses formatting); embedding an office engine (server dependency); converting through PDF (not editable).
- **Open risk**: full docx fidelity for complex layouts is not guaranteed; the supported operation set and refusal path are mandatory, not optional.

## R4. PPTX

- **Decision**: Read-only preview in v1.
- **Rationale**: Univer's Slides capability is commercial; a slide unit and its conversion are out of v1 scope.
- **Alternatives**: Buy the commercial slide module (rejected by the license principle); no preview at all (worse).

## R5. Legacy conversion (`.doc/.xls/.ppt`)

- **Decision**: Convert to the modern format before opening, using the existing LibreOffice kit; verify that the kit exposes target-format conversion (not only PDF) as the first Phase 1 task. If it does not, extend the independent kit or add a dedicated converter; never alter the source file.
- **Rationale**: Reuses an existing, license-compatible engine and keeps the original intact.
- **Alternatives**: Parse binary formats in-process (high effort, low fidelity); refuse legacy files (worse).

## R6. Stable document identity and rebind

- **Decision**: Assign each opened document a stable id scoped to the session, persisted in the session log with its current path and content version. On rename or move, rebind the id to the resolved path; if the file is missing, keep the id and report it.
- **Rationale**: Satisfies FR-015 and keeps restore working across path changes.
- **Alternatives**: Path as identity (breaks on rename); content hash as identity (collides on identical copies, recomputes on every save).

## R7. Attribution and undo

- **Decision**: Every applied change is a Change Record carrying actor, document version, and the inverse operation. User and agent edits go through the same command path so both are undoable at the document level.
- **Rationale**: Satisfies FR-014 and gives a single audit/reversal model.
- **Alternatives**: Editor-native undo only (does not cover agent edits or cross-actor changes); log-only attribution (not reversible).

## R8. Write authority and conflict

- **Decision**: All writes go through the filesystem seam with a version guard and sandbox roots. A stale version is rejected and reported as a conflict; no path silently overwrites.
- **Rationale**: Satisfies FR-005, FR-017 and matches the existing filesystem contract.
- **Alternatives**: Direct file writes from the editor (bypasses sandbox); last-write-wins (violates non-destructive editing).

## R9. Local-first privacy and egress

- **Decision**: No document or edit history leaves the device; the only expected network call is the model provider. Telemetry is off by default; any opt-in is explicit.
- **Rationale**: Satisfies FR-013 and the local-first principle.
- **Alternatives**: Background sync or crash uploads (violates privacy); default-on telemetry (violates the principle).

## R10. Credential storage

- **Decision**: Store model credentials in the operating system's secure store through the desktop shell; never commit them or write them to logs.
- **Rationale**: Satisfies FR-018.
- **Alternatives**: Plaintext config files (unsafe); environment-only (poor desktop UX).

## R11. Upstream stewardship

- **Decision**: Extend seams or add packages instead of rewriting upstream modules; merge upstream with a preserved-work recovery copy, a clean tree, no shared stash, and recorded conflict intent.
- **Rationale**: Satisfies FR-019 and Constitution VIII; keeps upstream fixes mergeable.
- **Alternatives**: Divergent rewrites (unmergeable); no upstream tracking (loses framework fixes).

## R12. Performance and bounds

- **Decision**: Parse and convert in a worker or child process with byte, cell, and time caps; fail loudly past the caps rather than truncating.
- **Rationale**: Prevents UI stalls and unbounded memory for large files.
- **Alternatives**: Main-thread parsing (freezes the app); silent truncation (data loss).

## R13. UnitSnapshot model ownership (Option B, accepted 2026-11-14)

- **Decision**: `UnitSnapshot.data` is the editor core's native snapshot type directly (`IWorkbookData` for workbooks, `IDocumentData` for documents). No project-owned neutral model. The editor and the agent share this model; the session log persists attributable change operations, not raw snapshots.
- **Rationale**: Eliminates an adapter hop, gives the highest fidelity, and lets the agent use the same runtime model and command/undo semantics as the editor. Writing raw snapshots to the log would couple the durable session format to a fast-moving third-party type, so operations stay the durable vocabulary (Constitution III).
- **Alternatives**: A project-owned neutral DTO with an editor adapter (extra mapping, fidelity loss, duplicated schemas); persisting raw snapshots (session-format instability, heavy events).
- **Consequence**: `packages/document/office-exchange` (Host) takes a dependency on `@univerjs/core`; T006 adds it and T008/T009 target the native types.

## R14. One live document instance (accepted 2026-11-14)

- **Decision**: The editor owns the only live unit for an open document. Agent document tools dispatch operations to that unit through a renderer-side command bridge, so both actors share one state, one version, and one undo stack. An operation against a document with no live instance is refused (`document-not-open`) rather than applied to a detached copy.
- **Rationale**: Directly satisfies FR-020, SC-003 (visible in the open editor), and SC-012 (agent edits undoable at the document level). It is the natural consequence of Option B: one model means one instance.
- **Alternatives**: A separate headless instance writing the file and forcing a reload (loses shared undo, harder 2 s visibility); the Host owning the unit with the editor as a remote view (fights Univer's DOM/command model).
- **Consequence**: Agent tool execution routes into the renderer for open documents; the bridge enforces the same version guard and Change Record path. A headless/CLI run without an editor refuses document operations.
