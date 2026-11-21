# Data Model: In-Place OOXML Save

**Input**: Feature specification from `/specs/002-in-place-ooxml-save/spec.md`

**Date**: 2026-11-18

## Overview

The model separates three concerns: the **source package** (immutable bytes and entry records), the **editable model** (Univer snapshot that crosses the wire and the editor), and the **splice plan** (the minimal set of byte-range operations that turn one into the other). Source identity and unit anchors travel with the editable model in a plain-JSON sidecar; package bytes never do.

### Mapping to the specification entities

The implementation entities below refine the spec's Key Entities; they are not additional product concepts:

| Spec entity | Implementation entities |
| --- | --- |
| Source Package | `SourcePackage`, `ZipEntryRecord` |
| Editable Part Set | `OfficePackage` part map and dirty set |
| Edited Range | `UnitAnchor`, `SheetSummary`, the `splice` operations |
| Construct Intersection | `ConstructIntersection`, `PreservationFailure` |
| Part Disposition | `PartDisposition` |
| Preservation Evidence | `SaveDisposition` plus the byte-diff report |
| Conflicting Constraint | `ConstructIntersection` returned as a refusal |

### Terminology

An **entry** is one ZIP record. A **part** is an OPC part name carried by an entry. The two coincide for every file part; directory entries ending `/` are carried as entries but have no part and no disposition. `PartDisposition.name` is the ZIP entry name as stored in the central directory.

## SourcePackage

The file as opened.

| Field | Type | Meaning |
| --- | --- | --- |
| `bytes` | `Uint8Array` | The complete source file, as read at `expectedVersion` |
| `entries` | `readonly ZipEntryRecord[]` | One record per central-directory entry, in central order |
| `localOrder` | `readonly ZipEntryRecord[]` | The same records ordered by local header offset |
| `comment` | `Uint8Array` | Archive comment, preserved verbatim |
| `digest` | `string` | SHA-256 hex of `bytes` |

`ZipEntryRecord`: `name`, `nameBytes`, `method` (`0 | 8`), `flags`, `crc32`, `compressedSize`, `uncompressedSize`, `localOffset`, `localHeaderLength`, `dataOffset`, `descriptorLength`, `localRecord`, `centralRecord`.

**Validation**:
- Reject a CFB/OLE wrapper, split archives, ZIP64, encrypted entries, unsupported methods, ambiguous descriptors, out-of-bounds offsets, local/central disagreement, duplicate names, and names equal under ASCII case folding.
- Directory entries ending `/` are carried but excluded from the part map.
- `localRecord` must be a contiguous view ending at the descriptor end.

## OfficeSourceState (sidecar)

Stored at `data.officeExchange.source` inside `IDocumentData`/`IWorkbookData`.

| Field | Type | Meaning |
| --- | --- | --- |
| `schemaVersion` | `1` | Sidecar format version |
| `format` | `'docx' \| 'docm' \| 'xlsx' \| 'xlsm'` | Format as opened |
| `packageDigest` | `string` | SHA-256 hex of the source bytes at read |
| `byteLength` | `number` | Source byte length at read |
| `anchorMode` | `'per-unit' \| 'sheet-summary'` | Which anchor representation is present |
| `anchors` | `readonly UnitAnchor[]` | Present for `per-unit` |
| `sheets` | `readonly SheetSummary[]` | Present for `sheet-summary` |
| `sourceParts` | `readonly string[]` | Parts that carry editable units |

**Validation**: a malformed sidecar fails with `conversion-failed`; a missing `packageDigest` means "legacy snapshot" and selects the fallback writer (R11). `anchors` and `sheets` are mutually exclusive with `anchorMode`.

## UnitAnchor

One editable unit in one source part.

| Field | Type | Meaning |
| --- | --- | --- |
| `kind` | `'paragraph' \| 'tableCell' \| 'sheetCell'` | Unit kind |
| `part` | `string` | Decoded part name, e.g. `word/document.xml`, `xl/worksheets/sheet1.xml` |
| `key` | `string` | Model identity: `paragraphId`; `${tableId}/${row}/${cell}`; `${sheetId}!${row}:${column}` |
| `start` | `number` | Inclusive byte offset in the uncompressed part bytes |
| `end` | `number` | Exclusive byte offset |
| `digest` | `string` | SHA-256 hex of the canonical unit projection at read |

**Validation**:
- `part` exists in the source package.
- `0 ≤ start < end ≤ partByteLength`.
- Spans for one part do not overlap.
- `key` is unique within `(part, kind)`.
- `digest` is 64 lowercase hex characters.
- Total anchor count does not exceed `maxSourceAnchors` (default 100,000, configurable).
- The anchor payload byte length does not exceed `maxSourceAnchorBytes` (default 8 MiB).

## SheetSummary (large-sheet fallback)

| Field | Type | Meaning |
| --- | --- | --- |
| `sheetId` | `string` | Model sheet id |
| `part` | `string` | Resolved worksheet part |
| `digest` | `string` | Read-time digest of the canonical sheet projection |
| `cellCount` | `number` | Non-empty cells at read |

In `sheet-summary` mode the writer re-reads the source at save, rebuilds the same anchors, and diffs the current model against the source projection by address. If the diff cannot be completed within the bounds, the save fails with `bound-exceeded` rather than splicing a partial plan.

## Canonical unit projection

The digest input is a deterministic JSON value with sorted keys, excluding positional fields:

- Paragraph: `dataStream` slice covered by the paragraph, `textRuns` inside it (style id and properties), `paragraphStyle`, `bullet`, and the list id.
- Table cell: the paragraphs and nested tables inside the cell, plus the cell's modelled properties.
- Sheet cell: `v`, `t`, `f`, `p` (hyperlink), and the modelled style object.

`startIndex`, array positions, and snapshot ids are excluded so an editor round trip that shifts offsets does not look like a change.

## SplicePlan

The writer's output plan, derived from the anchors and the current model.

| Field | Type | Meaning |
| --- | --- | --- |
| `operations` | `readonly SpliceOperation[]` | Ordered, non-overlapping per part |
| `touchedParts` | `readonly string[]` | Parts that will differ from the source |
| `entries` | `readonly PartDisposition[]` | One entry per affected package entry |
| `refusal` | `ConstructIntersection \| undefined` | Present only when the plan cannot be completed |

`SpliceOperation` variants:

- `{ kind: 'carry'; part; entry }` — no byte change.
- `{ kind: 'replace'; part; start; end; bytes }` — replace one edited unit span.
- `{ kind: 'insert'; part; at; bytes }` — insert a unit at an anchor boundary.
- `{ kind: 'delete'; part; start; end }` — remove a unit span.
- `{ kind: 'append'; part; at; bytes }` — auxiliary-part append before the closing tag (style, list, relationship, content type).

**Invariants**:
- Operations for one part are sorted and non-overlapping.
- Every source byte outside `replace`/`delete` spans and every inserted byte is accounted for.
- A plan with a refusal has no operations.
- The union of `touchedParts` is minimal for the edit.

## PartDisposition

| Field | Type | Meaning |
| --- | --- | --- |
| `name` | `string` | Package entry name |
| `status` | `'verbatim' \| 'spliced' \| 'added' \| 'removed'` | What happened |
| `spans` | `readonly { start; end }[]` | Present for `spliced` |
| `reason` | `string` | Present for `added`/`removed` and for any non-obvious splice |

`verbatim` means the local record bytes are copied unchanged, including compressed data, CRC, timestamps, extra fields, and any descriptor.

## SaveDisposition (session fact payload)

| Field | Type | Meaning |
| --- | --- | --- |
| `schemaVersion` | `1` | Payload version |
| `verbatimCount` | `number` | Entries copied unchanged |
| `spliced` | `{ count; names }` | Entries with at least one span operation |
| `added` | `{ count; names }` | New entries |
| `removed` | `{ count; names }` | Removed entries |
| `nameLimit` | `number` | Cap applied to each `names` array (256) |

**Validation**: counts are exact; `names` are sorted ASCII and truncated to `nameLimit`; the payload is bounded and never fails the save.

## ConstructIntersection

| Field | Type | Meaning |
| --- | --- | --- |
| `kind` | `UnsupportedFeatureKind` | Reused vocabulary: `macro`, `tracked-changes`, `comment`, `chart`, `drawing`, `shape`, `embedded-object`, `other` |
| `detail` | `string` | The construct and its location, e.g. `w:ins in word/document.xml at 412-530` |
| `part` | `string` | Part name |
| `start` | `number` | Span start or anchor-relative position |

**Rule**: the intersection is computed from the source span of the edited units, not from the whole file. Independent unmapped properties are carried and do not produce an intersection.

## PreservationFailure

A typed reason a plan cannot be completed: `ambiguous-span`, `unsupported-coupling`, `digest-mismatch`, `anchor-budget-exceeded`, `unsupported-container`, `encoding-unsupported`.

Failures map to `conversion-failed` (malformed state), `conflict` (source digest/version mismatch), `bound-exceeded` (budget), or `unsupported-feature` (coupled construct) at the Remote boundary.

## Relationships

```
SourcePackage 1 ── * ZipEntryRecord
OfficeSourceState 1 ── * UnitAnchor
OfficeSourceState 1 ── * SheetSummary
OfficeSourceState 1 ── 1 SourcePackage (by packageDigest)
SplicePlan 1 ── * SpliceOperation
SplicePlan 1 ── * PartDisposition
SaveDisposition 1 ── * PartDisposition (counts and names only)
ConstructIntersection 0..1 ── 1 SplicePlan (refusal)
```

## State transitions

| From | Event | To | Guard |
| --- | --- | --- | --- |
| Closed | `read` | Anchored | Source digest recorded; anchors built |
| Anchored | editor mutation | Anchored (stale model) | Anchors unchanged |
| Anchored | `save` | Saved | Version and digest match; plan complete |
| Anchored | `save` with changed source | Conflict | Version mismatch or digest mismatch |
| Anchored | `save` with coupled edit | Refused | Intersection non-empty; source unchanged |
| Anchored | draft write | Drafted | Whole snapshot including sidecar |
| Drafted | restore | Anchored | Draft sidecar present; else Legacy |
| Legacy | `save` | Saved (legacy writer) | Fallback path; no span guarantee |
| Saved | `read` | Anchored | New source digest and anchors |

## Persistence and compatibility

- `OfficeSourceState` lives inside the editable data object, so it survives `unit.save()`, the JSON wire, drafts, and recovery without a session-format change.
- The `document/save` disposition property is optional and additive. `SESSION_FORMAT_VERSION` stays 4. Older readers ignore it and keep the save fact; new readers treat absence as a pre-disposition save.
- Widening the persisted `DocumentFormat` with `docm`/`xlsm` is a separate persistence-type change and requires its own acknowledgement. Because an older client's projection schema union rejects the new literals, `documentProjectionDefinition.stateVersion` bumps from 1 to 2; the projection schema, `sameEntry`, and their tests change with it.
