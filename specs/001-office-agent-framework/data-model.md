# Data Model: Office Agent Framework

**Input**: [spec.md](./spec.md), [plan.md](./plan.md)
**Date**: 2026-11-13

## Overview

Seven entities carry the feature: the **Document**, the **Document Session Fact**, the **Document Command**, the **Change Record**, the **Unit Snapshot**, the **Conversion Record**, and the **Document Draft**. The Document Session Fact records `open`, `edit`, `save`, `rebind`, `missing`, and `close`; the Document Command is the transport request that carries no snapshot and appends no content. A **Document Identity Binding** links a stable document id to its current path. All durable state is reconstructable from the session log (Principle III).

## Document

An open office file and its editable state.

| Field | Type | Rules |
|---|---|---|
| `documentId` | opaque string | Unique within the session; stable across rename/move (FR-015) |
| `sessionId` | string | Authorizing session; a document never crosses sessions |
| `path` | absolute path | Current locator; may be rebound; never used as identity |
| `format` | `docx`/`xlsx`/`pptx` | Editable or read-only per v1 scope (FR-011) |
| `contentVersion` | monotonic version | Bumped only by an accepted save; guards conflicts (FR-005) |
| `status` | enum | See state transitions |
| `unsupportedFeatures` | list | Non-empty blocks a lossy save (FR-008) |
| `dirty` | boolean | True when in-memory differs from `contentVersion` |

**Relationships**: one Document has many Document Session Facts and many Change Records; it holds one current Unit Snapshot.

## Document Session Fact

A logged, durable fact about a document. The log is the source of truth.

| Field | Type | Rules |
|---|---|---|
| `event` | `open`/`edit`/`save`/`close`/`missing`/`rebind` | Closed set; a new member follows the repository format decision and must be safely refused by older readers |
| `documentId` | string | References a Document |
| `sessionId` | string | Authorizing session |
| `turn` | turn id | Owning turn for model-visible reconstruction |
| `contentVersion` | version | Document version at the time of the fact |
| `actor` | `user`/`agent`/`system` | Who caused the fact |
| `at` | timestamp | Recorded by the log, not the caller |

**Validation**: an `edit` fact is only written after its Change Record is accepted; a `save` fact records an accepted save whose in-memory changes stayed in the editor and therefore carries no operation steps or inverse; a `close` fact does not delete history.

## Document Command

A correlated request from the Host tool layer to the live editor unit (FR-020). It is transport, not an accepted change.

| Field | Type | Rules |
|---|---|---|
| `requestId` | opaque string | Correlation id chosen by the Host; the editor answers with it exactly once |
| `documentId` | string | References an open Document; another document's command is ignored |
| `operation` | read or ordered edit steps | Names the work to perform, never a unit snapshot |
| `fromVersion` | version | The version the caller believes the live unit holds |
| `turn` | turn id | Owning turn for model-visible reconstruction |

**Validation**: a command appends no content fact by itself; a refused, unanswered, or timed-out command writes nothing, and only the accepted save appends `document/edit` (agent steps with inverse) or `document/save` (an editor-side user save).

## Change Record

An attributable, reversible edit (FR-014).

| Field | Type | Rules |
|---|---|---|
| `changeId` | opaque string | Unique and ordered |
| `documentId` | string | References a Document |
| `actor` | `user`/`agent` | Never `system` for content changes |
| `fromVersion` / `toVersion` | version | `toVersion = fromVersion + 1` |
| `operation` | structured edit | Cell/row/column/paragraph operation |
| `inverse` | structured edit | Reverses `operation`; required for undo |
| `at` | timestamp | Logged |

**Validation**: a Change Record is accepted only if `fromVersion` matches the current version; otherwise it is rejected as a conflict (FR-005). The inverse must restore the prior content exactly. A Change Record is created when a save is accepted; in-memory commands become facts only at that point. An agent save never includes unsaved user edits (it fails with `pending-edits` instead).

## Unit Snapshot

The editable content model shared by the editor and the agent.

| Field | Type | Rules |
|---|---|---|
| `kind` | `workbook`/`document` | Matches `format` |
| `data` | `IWorkbookData` (workbook) or `IDocumentData` (document) | The editor core's native snapshot type; validated at the file and model boundaries; never trusted raw |

**Validation**: a snapshot is bound to a `contentVersion`; a snapshot derived from a stale version is discarded rather than written. The session log stores change operations (Change Records), never raw snapshots.

## Conversion Record

The result of mapping a file to and from a snapshot.

| Field | Type | Rules |
|---|---|---|
| `format` | office format | Detected from the file, not assumed |
| `sourcePath` | absolute path | The file read |
| `sourceHash` | hash | Detects out-of-band change between read and write |
| `unsupported` | list | Same shape as `Document.unsupportedFeatures` |
| `bounded` | boolean | True when the conversion stopped at a cap (fail loud) |

**Validation**: a write re-checks `sourceHash`; a mismatch means an external change and forces a reload/conflict rather than a blind overwrite.

## Document Identity Binding

| Field | Type | Rules |
|---|---|---|
| `documentId` | string | Stable id |
| `path` | absolute path | Current locator |
| `missing` | boolean | True when the path no longer resolves |
| `reboundAt` | timestamp? | Set when the path changed |

**Validation**: rebinding preserves `documentId`, `sessionId`, and history; a missing binding keeps the document open and reports it (edge case: rename/move).

## Document Draft

A recoverable local draft for unsaved edits (FR-016).

| Field | Type | Rules |
|---|---|---|
| `documentId` | string | References the open Document |
| `contentVersion` | version | The version the draft was derived from |
| `path` | absolute path | Draft location; never the document's own path |
| `savedAt` | timestamp | Last draft write |
| `snapshot` | Unit Snapshot | Native editor snapshot; not written to the session log |

**Validation**: a draft never overwrites the original file; a successful save clears it; on restart a draft is offered only when its `contentVersion` matches the file or the user explicitly chooses it.

## State Transitions

```text
closed --open--> opened --edit--> dirty --save(accepted)--> saved
dirty --save(stale version)--> conflict --reload/merge--> dirty
dirty --close--> close-warned --discard--> closed
opened --external change--> reload-prompted
opened|dirty --path moved--> rebind --> opened|dirty
opened|dirty --path missing--> missing (document retained)
```

**Rules**

- No transition from `dirty` to `closed` may discard edits without a warning or a recoverable draft (FR-016).
- `conflict` never overwrites; the user or agent must resolve it (FR-005).
- `missing` keeps the document and its history; it never silently closes.
