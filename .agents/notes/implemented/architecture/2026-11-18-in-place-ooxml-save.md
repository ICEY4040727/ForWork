# Agent Note: In-place OOXML save preserves source package bytes

Status: implemented

English | [中文](2026-11-18-in-place-ooxml-save.zh.md)

## Problem

Rebuilding an Office package from the editable model loses or refuses every construct the model does not represent. A real `.docx` or `.xlsx` commonly carries a chart, a custom theme, a pivot table, a tracked change in another clause, or a VBA project; a regeneration round-trip drops those parts, and the previous refusal treated their presence anywhere in the file as a reason not to save at all. The user-visible cost is either silent loss or an unsaveable document. Preservation has to start from the file the user opened, not from what the editor can serialize.

## Decision

`@forwork/office-exchange` writes through two layers. A Host-only ZIP/OPC layer indexes the source package, copies every untouched entry from its source local record, and rebuilds the central directory; an XML span layer replaces only the edited unit inside a part that changes. A no-op save returns the source bytes. `.docm` and `.xlsm` are editable under the same rules, so the VBA project and the macro-enabled main content type survive. A source-less snapshot still uses the regeneration writer.

The [specification](../../../../specs/002-in-place-ooxml-save/spec.md), [plan](../../../../specs/002-in-place-ooxml-save/plan.md), [write contract](../../../../specs/002-in-place-ooxml-save/contracts/office-exchange-write.contract.md), and [data model](../../../../specs/002-in-place-ooxml-save/data-model.md) own the complete acceptance criteria and entity definitions. The [package README](../../../../packages/document/office-exchange/README.md) owns the consumer contract.

### Why raw ZIP local-record copy

No installed ZIP library can pass a compressed entry through unchanged. `fflate.unzipSync` returns decompressed payloads, and `zipSync` recomputes the CRC, recompresses, sorts entries, writes a fixed mtime, and cannot reproduce an archive comment. `yauzl` exposes compressed bytes but has no writer, and `jszip`, `archiver`, and `unzipper` all rebuild or extract. A no-op save that rebuilt the archive would therefore fail the byte-identity requirement before any edit was involved.

The package writer copies each untouched `localRecord` verbatim, including its compression method, compressed data, CRC, timestamps, extra fields, and optional data descriptor. It patches only the offset field of a carried central-directory record and writes the end-of-central-directory with the original archive comment. A changed entry is re-encoded with `node:zlib` `crc32` and `deflateRawSync`, while a stored entry stays stored. This keeps the change confined to the edited entry and keeps the dependency count at zero; [R1](../../../../specs/002-in-place-ooxml-save/research.md) records the library investigation.

### Why a custom byte scanner

The scanner must report offsets into the original UTF-8 part, because a splice written at a parser-normalized offset corrupts the file. `fast-xml-parser` 5.11.1 has `captureMetaData`, but it normalizes `\r\n` and `\r` to `\n` before recording `startIndex` and `endIndex`, so every offset after the first CRLF is wrong for the source bytes. `fast-xml-validator` exposes only line and column and strips the BOM and declaration before scanning. The existing regex catalog is not quote-, comment-, or CDATA-aware and assumes a literal `w:` prefix.

`src/source/xml-spans.ts` scans the part once, resolves namespace URIs, tracks quotes, comments, CDATA, processing instructions, and DOCTYPE, and handles self-closing tags and whitespace before `>`. It preserves CRLF and a BOM and fails closed on an ambiguous span instead of guessing an offset. The parser remains a test oracle on LF-only fixtures; it cannot be the production source, as [R3](../../../../specs/002-in-place-ooxml-save/research.md) records.

### Why anchors live at `data.officeExchange.source`

The snapshot must carry source identity and per-unit anchors through `unit.save()`, the JSON wire, the editor round trip, and a recovery draft; a place that any of those steps drops or rewrites makes a later save unable to find its spans. The `OfficeDocumentSnapshot` wrapper has a declared shape whose unknown fields the codec strips, `resources` is replaced on save, and `body.payloads` is documented as not saved to disk. A per-paragraph field has no durable slot, and a cell's `custom` field is deleted by auto-fill and duplicated by copy/paste.

A provider-owned, plain-JSON namespace inside the editable `data` object survives all four paths. It holds the schema version, opened format, source digest and byte length, anchor mode, source parts, and either `UnitAnchor` entries or per-sheet summaries. Each anchor carries a model key, byte span, and canonical unit digest, so save-time change detection compares identities and content rather than positional indices. [R5](../../../../specs/002-in-place-ooxml-save/research.md) records the transport alternatives.

### Coupled versus independent constructs

Presence of an unmodeled construct in the source is not a refusal. The writer classifies the edited source span. A property whose meaning does not reference the changed text or node identity is independent and is merged into the regenerated unit: a cell's `s` style index, `w:tcPr` and `w:trPr` children, unmapped `w:pPr` and `w:rPr` children, cell width, and fit-to-text. A construct whose meaning or anchor depends on the changed content is coupled, and the save refuses with that construct named: a tracked change or comment range overlapping the unit, a content control wrapping the edited text, a field whose result is the text, a rich-text shared string, a shared or array formula, a changed worksheet hyperlink whose target lives in the relationship part, and an ambiguous `mc:AlternateContent` branch.

The classification is derived from the source span, not from a file-level unsupported list, so a chart in another paragraph never blocks a text edit while a tracked change around the edited sentence does. [R12](../../../../specs/002-in-place-ooxml-save/research.md) owns the refusal rule.

### Re-anchoring on save

A successful save rewrites the file, so the snapshot's read-time offsets and digests no longer describe the bytes on disk. The editor adopts the published version without re-reading, which keeps the live unit and its undo history. A second save therefore supplies the new source bytes together with a sidecar recorded against the old ones.

The writer compares the sidecar's `packageDigest` and `byteLength` with the supplied source. On a mismatch it rebuilds the anchors from that source before planning the splice, then diffs the current model against those anchors by the same model keys. A source that cannot be re-anchored is refused as `conversion-failed`. The filesystem `replaceIfVersion` guard remains authoritative for a concurrent external change, so a stale source is a conflict, not a re-anchor.

## Alternatives considered

**Whole-package rebuild with raw blocks for unmodeled parts.** The regeneration writer already carried some parts opaquely, but it still rebuilt every entry, reordered or recompressed the archive, and lost any part the model did not name. It cannot satisfy a byte-identical no-op save or span-level identity, and it is the behavior this decision replaces.

**Write through an installed ZIP library.** `fflate` cannot reuse a compressed stream, preserve local and central metadata, or write an archive comment; other installed libraries extract or recompress. Adding a dependency for a writer that still recompressed would not preserve bytes, as [R1](../../../../specs/002-in-place-ooxml-save/research.md) records.

**Use `fast-xml-parser` `captureMetaData` for span offsets.** Its CRLF normalization makes every offset after the first CRLF wrong. Keeping it as the test oracle on LF-only fixtures preserves its value without making it authoritative.

**Put anchors in the snapshot wrapper, `resources`, or per-unit fields.** The wrapper strips unknown fields, `resources` is replaced, `body.payloads` is not persisted, paragraphs have no custom slot, and cell `custom` is not stable across copy/paste or auto-fill.

**Refuse every unmapped property inside an edited unit.** This keeps the refusal set large and rejects ordinary files whose edited paragraph carries a property the model does not name. Merging independent properties and refusing only coupled ones is the smaller residual set; refusing everything was the simpler alternative and lost the feature's user-visible benefit.

**Always carry every construct inside the edited span.** A carried tracked change or shared formula can point at content the edit removed, yielding an invalid or semantically wrong document. Those constructs are coupled and must refuse.

**Require a re-read after every save instead of re-anchoring.** A re-read discards the live unit and its undo history, and it forces the editor to rebuild content the writer can already diff against the new source under the existing version guard.

## Consequences

The package grows a Host-only ZIP/OPC layer and a source-span layer, each with its own tests and failure modes; `node:zlib` stays the only new runtime primitive. A no-op save and every untouched entry are byte-identical, and the fixed release subset of the corpus imports, renders, and reads back in the system LibreOffice on Windows. LibreOffice can repair a package silently and reports no repair diagnostics, so that matrix is not a no-repair guarantee for desktop Word or Excel; the limitation is disclosed instead of claimed. A snapshot without a source sidecar, including an old draft, keeps the earlier regeneration behavior, so the package's contract now has two write paths with different preservation guarantees.

Auxiliary parts are append-only: a required style or numbering definition is added without regenerating the part, and an edited worksheet cell style reuses a matching `<xf>` or appends only the needed number-format, font, fill, border, and `<xf>` entries to `xl/styles.xml` while updating that collection's count. A Word table geometry change or a unit reorder is refused because it cannot be expressed as span operations, and a worksheet style component the writer cannot express refuses with the component named. The save disposition is durable but audit-only; the `document` projection does not expose it. The [fixture corpus](../../../../packages/document/office-exchange/tests/fixtures/README.md) and the [LibreOffice load smoke](../../../../packages/document/office-exchange/tests/ooxml-engine.spec.ts) are the evidence base, and the byte-diff harness reconciles each reported disposition with the actual output.
