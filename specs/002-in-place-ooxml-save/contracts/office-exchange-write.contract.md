# Contract: Office Exchange Write

**Feature**: 002-in-place-ooxml-save

**Consumers**: `OfficeDocumentRemote.save`, `document_*` tools through the Remote, direct provider callers in tests.

## Terminology

An **entry** is one ZIP record. A **part** is an OPC part name carried by an entry. `PartDisposition.name` is the ZIP entry name from the central directory. Directory entries ending `/` are carried but have no part and no disposition.

## Operations

### `OfficeExchange.read(path, signal?): Promise<ExchangeReadResult>`

Unchanged behavior, plus one postcondition: the returned native snapshot carries `data.officeExchange.source` when the provider can build anchors for the format and package. A package the writer cannot preserve is still readable; the refusal, if any, happens at save.

### `OfficeExchange.write(snapshot, format, source, signal?): Promise<OfficeWriteResult>`

- `snapshot`: the native unit snapshot (document or workbook).
- `format`: `docx | docm | xlsx | xlsm`.
- `source`: the `SourcePackage` bytes read at the caller's `expectedVersion`. Required when the source file exists.
- `signal`: optional cancellation.

Returns `{ bytes, disposition }`:

- `bytes`: the complete output file.
- `disposition`: a `SaveDisposition` describing what was carried, spliced, added, and removed.

### `OfficeExchange.detectPreserved(bytes, config?): PreservedConstructs`

Pure, byte-based classification used at save. It reports constructs that are preserved verbatim and constructs that would be coupled to an edit. It never refuses on its own: refusal is decided against the edited ranges.

### `OfficeExchange.detectUnsupported(path, signal?)`

Retained for callers that need the source-level feature report. It no longer blocks a save by itself.

## Preconditions

1. `source` is the exact byte sequence the caller stat-verified at `expectedVersion`.
2. `snapshot.officeExchange.source.packageDigest` equals `sha256(source)` and `byteLength` matches, unless the snapshot is legacy (no sidecar).
3. The format matches the snapshot kind (`document` ↔ `docx/docm`, `workbook` ↔ `xlsx/xlsm`).
4. The source is an OOXML package the container layer supports: methods 0 and 8, no ZIP64, no encryption, no split archive, no ambiguous descriptor, unique OPC part names under ASCII case folding.
5. Every anchor is inside its part and non-overlapping.

## Postconditions

1. **No-op identity**: when the model produces no span operation, `bytes` equals `source` byte-for-byte, including entry order, compressed streams, CRCs, timestamps, extra fields, and archive comment.
2. **Untouched entry identity**: for every entry without a span operation, the output local record equals the source local record byte-for-byte.
3. **Span identity**: inside an entry that changes, every byte outside the reported span operations equals the source byte at the same position.
4. **Minimality**: no entry, span, or auxiliary append exists that the edit and its required reference updates do not justify.
5. **Validity**: the output is a well-formed OPC package; every relationship target resolves, every part has a content type, no declaration points at a missing part, and every authored XML part is well-formed.
6. **Determinism**: the same `(snapshot, format, source)` produces identical bytes.
7. **Disposition truth**: the disposition's counts equal the actual entry statuses and its names equal a deterministic prefix of the affected entries.
8. **Format preservation**: a `.docm`/`.xlsm` output keeps the macro-enabled main content type and the VBA project entry.

## Errors

| Code | Condition | Source effect |
| --- | --- | --- |
| `conflict` | `packageDigest`/`byteLength` mismatch, or source changed since read | None |
| `unsupported-feature` | A coupled construct overlaps an edited span: tracked change, comment range, content control wrapping the text, field owning the text, rich-text shared string, shared/array formula, ambiguous alternative-content branch | None |
| `bound-exceeded` | Anchor count/bytes, entry count, expanded size, output size, or cell count over bound | None |
| `conversion-failed` | Malformed sidecar, unlocatable span, unsupported container structure, or post-splice validation failure | None |
| `unsupported-format` | Format outside `docx/docm/xlsx/xlsm` | None |

A refusal never writes bytes and never records a save fact.

## Guarantees

- The writer never reconstructs an entry from the editable model when the entry is not in the edited set.
- The writer never re-encodes a carried entry.
- The writer never normalizes, reorders, or drops a source part it does not edit.
- A snapshot without `officeExchange.source` uses the legacy regeneration writer; that path does not claim span identity.
- No part is read into the editor model merely to copy it.

## Invariants

- `SplicePlan` operations per part are sorted and non-overlapping.
- A plan with a `ConstructIntersection` has zero operations.
- `touchedParts` equals the set of parts named by span operations plus auxiliary appends.
- The disposition is computed from the plan, never from a caller-supplied summary.

## Compatibility

- `OfficeExchange.write` gains a parameter and a richer return type. Every caller is updated in the same change (Remote, provider, tests, generated catalog, subsystem docs, `001` contract).
- Snapshots produced before this change keep working through the legacy writer; drafts carrying the old `officeExchange.docx` sidecar are still accepted.
- Read results gain the `source` sidecar; consumers that ignore unknown fields are unaffected.
- No `SESSION_FORMAT_VERSION` change is implied by this contract.
