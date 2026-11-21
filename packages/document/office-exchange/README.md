---
description: "Office Exchange capability seam: OOXML-to-unit-snapshot mapping, in-place source-package saves, preservation and coupled-construct refusal, and legacy conversion shared by the office editor and agent document tools."
kind: "package-reference"
---

# @forwork/office-exchange

English | [中文](README.zh.md)

## Summary

Define the contract for moving Office files into the editable unit model and writing them back into the package they came from. The seam names conversion, preservation, refusal, and save disposition; providers own the ExcelJS, docx, and LibreOffice mechanics. The editor and agent tools consume `ctx.officeExchange` without importing a conversion library.

## Table of Contents

- [Use this package](#use-this-package)
- [Understand the implementation](#understand-the-implementation)
- [Further Exploration](#further-exploration)
- [Model Experience](#model-experience)
- [Known Limitations and Deferred Work](#known-limitations-and-deferred-work)
- [Dev Note](#dev-note)

-----

<a id="use-this-package"></a>
## Use this package

The package default-exports `LocalOfficeExchange`, a Cordis plugin that registers `ctx.officeExchange`. Mount it with deployment bounds, then read and write through the seam:

```ts
import LocalOfficeExchange from '@forwork/office-exchange'

declare const ctx: import('@forwork/cordis').Context
declare const path: string
declare const signal: AbortSignal

await ctx.plugin(LocalOfficeExchange, {
  maxFileBytes: 25 * 1024 * 1024,
  maxCells: 2_000_000,
  maxConversionMs: 60_000,
  commandTimeoutMs: 5_000,
  maxSourceAnchors: 100_000,
  maxSourceAnchorBytes: 8 * 1024 * 1024,
})

const { format, snapshot, unsupported } = await ctx.officeExchange.read(path, signal)
```

`read(path)` returns the detected format, an editable unit snapshot, and every feature a whole-file regeneration would drop. `write(snapshot, format, source?, signal?)` encodes the snapshot back to file bytes without writing them and returns `{ bytes, disposition }`; `source` is the package read at the caller's expected version, and its presence with a snapshot sidecar selects the in-place writer. `detectUnsupported(path)` reports loss-prone features without reading content into the editor. `detectPreserved(bytes)` classifies package bytes the caller already holds into preserved construct classes and coupled-construct markers. `convertLegacy(path, signal?, outputDir?)` turns a `.doc/.xls/.ppt` file into its editable equivalent; an omitted `outputDir` writes into a fresh temporary directory outside any workspace.

Supported formats:

- `.docx`, `.docm`, `.xlsx`, and `.xlsm` are editable: `read` maps them to Univer document and workbook snapshots, and `write` splices the edited units back into the source package. An in-place save keeps a macro-enabled file's VBA project and macro-enabled main content type.
- `.doc`, `.xls`, and `.ppt` must be converted with `convertLegacy` before `read`; reading one directly raises `unsupported-format`.
- `.pptx` is preview-only: `detectUnsupported` inspects it, while `read` and `write` raise `unsupported-format`.

Every failure is an `OfficeExchangeError` with a stable `code`; bounds resolve once at mount and invalid values fail loudly.

`formatOf(path)` classifies a path by extension; `isEditableFormat` separates directly editable `.docx/.docm/.xlsx/.xlsm` from formats that require conversion; `isLegacyFormat` names the legacy set.

### In-place save

With a snapshot that carries `data.officeExchange.source` and the source bytes, `write` assembles the output from the source package. Every entry the edit does not change is copied from its source ZIP local record, so compressed data, CRC, timestamps, extra fields, compression method, entry order, and data descriptors stay identical. A save that changes no unit returns the source bytes unchanged.

Only the edited unit's XML span changes inside a part: a paragraph, a table cell, or a worksheet `<c>` element. A body-level content control anchors to its whole `w:sdt` element, so an edit of that unit is refused while a sibling edit leaves it untouched. Every byte outside the edited spans is copied from the source. A worksheet row or column insert or delete is expressed as row or cell span operations rather than a whole-sheet rewrite; a shift on a sheet that carries a merged range, conditional-formatting range, data-validation range, autoFilter range, or Excel-table address range is refused with that construct named, because the range coordinates would not move with the shifted rows or columns, and a declared table part whose relationship, target part, or `ref` cannot be resolved to a range refuses the shift instead of leaving a stale table reference. A required style, numbering definition, relationship, or content-type declaration is appended before its part's closing tag, and an edited cell style appends only the style components it needs; no auxiliary part is regenerated, `word/settings.xml` is never written, and a changed worksheet string becomes an inline string so `xl/sharedStrings.xml` stays untouched.

The source sidecar records `schemaVersion: 1`, the opened `format`, the source `packageDigest` and `byteLength`, the `anchorMode`, the `sourceParts`, and either per-unit `anchors` or per-sheet `sheets`. Each anchor names the unit kind, its part, a model key, the byte span in the uncompressed part, and a digest of the canonical unit. It survives the editor's snapshot serialization, the JSON wire, and a recovery draft. `maxSourceAnchors` (100,000 by default) and `maxSourceAnchorBytes` (8 MiB) bound the payload; a workbook above the per-cell count stores sheet summaries and rebuilds cell anchors at save, while a source that cannot be anchored is read without a sidecar.

When a later save supplies a source whose digest or length differs from the sidecar, the writer re-anchors that source before planning, so saving twice without a re-read works. A source that cannot be re-anchored is refused. A snapshot without a sidecar, or a caller that supplies no source, uses the legacy regeneration writer, which rebuilds the package from the model and reports a disposition that claims no verbatim entries; a `.docm` or `.xlsm` target refuses that path because regeneration would write a macro-free package without the VBA project.

### Preservation and refusal

`read` reports preserved construct classes in `preserved` and whole-file regeneration losses in `unsupported`; neither list blocks a save. The in-place writer carries these classes verbatim outside an edited unit: macros, charts, pivot tables, Excel tables, defined names, calc chains, comments, embedded objects, drawings, shapes, themes, font tables, web settings, custom XML, external links, fields, bookmarks, content controls, and document properties.

A save is refused when the edit intersects a construct whose meaning or anchor depends on the changed content: a tracked change or comment range over the unit, a content control wrapping the edited text, a field that owns the text, a rich-text shared string, a shared or array formula, a changed worksheet hyperlink, or an ambiguous `mc:AlternateContent` branch. A span the writer cannot locate or express, and a style component it cannot represent, also refuses. The refusal names the cause and its location, writes nothing, and records no save fact. An unmapped property that is independent of the change is re-emitted inside the spliced unit instead of refusing, including a cell style index, table-cell and row properties, unmapped paragraph and run properties, cell width, and fit-to-text.

A structural row or column shift is refused only when the sheet carries one of the address-bearing ranges named above, and a macro-enabled target without its anchored source refuses rather than falling back to the regeneration writer. A preserved construct elsewhere in the file never refuses an edit.

### Save disposition

`write` returns a `SaveDisposition`: `schemaVersion: 1`, an exact `verbatimCount`, and one `{ count, names }` group each for `spliced`, `added`, and `removed` entries. Counts are exact; every `names` list is sorted ASCII and capped at 256 entries. The Remote records the same value on `document/save` as an optional property, so an older save or a legacy regeneration reads as absent rather than invalid.

### Unsupported-feature detection

`read` and `detectUnsupported` report every construct a whole-file regeneration would drop. The list is informational for an in-place save: carrying the construct does not by itself cause a refusal, and only a coupled construct intersecting the edit, or a span the writer cannot preserve, does.

- Documents: macros, tracked changes, comments, charts, drawings, shapes, and embedded objects.
- Workbooks: charts, comment and non-picture drawing parts, macros, and embedded objects.

Run formatting (font family, size, color, highlight or shading, strike, caps, small caps, vertical alignment, underline variants, and run borders), paragraph formatting (alignment, spacing, line spacing, indentation, tab stops, outline level, shading, and paragraph borders), `word/styles.xml` document defaults, bold, italic, underline, heading styles, section page setup (page size, orientation, margins, header/footer distances, columns, and break type), header and footer text with their default, first-page, and even-page references, lists with their `word/numbering.xml` definitions and per-paragraph `w:numPr`, hyperlinks with their external relationship target or `w:anchor` bookmark, table structure (merged cells, cell borders, shading, vertical alignment and margins, row height and header repeat, table width, alignment, indentation and layout, and tables nested in a cell), number formats, fills, workbook cell borders, horizontal and vertical alignment, wrap and shrink-to-fit, text rotation, strike, superscript and subscript, frozen panes, sheet zoom, right-to-left view, header visibility, tab colors, the very-hidden sheet state, row and column styles, merged ranges, row heights, and column widths round-trip, so detection does not report them.

Workbook conditional formatting, data validation, cell hyperlinks, cell images, and worksheet background images map into the sheet plugin models and write back. Conditional-formatting rules, validation rules, and drawings persist in the workbook snapshot through the `SHEET_CONDITIONAL_FORMATTING_PLUGIN`, `SHEET_DATA_VALIDATION_PLUGIN`, and `SHEET_DRAWING_PLUGIN` resources; a hyperlink target persists in its cell's rich-text custom range. Detection reports none of them.

Document images and footnotes or endnotes are preserved rather than mapped: each drawing, picture, or note-reference run is captured byte-for-byte with its media part, relationship, note part, and content-type declaration, and written back unchanged. They do not render or become editable in the editor; the surrounding text does.

A file opens with every preserved construct carried, not refused. Common run and paragraph formatting is mapped into the document model and written back on save, so an edit keeps it without a refusal; an edit that intersects a construct whose meaning or anchor depends on the changed content is refused with that construct named.

### Remote surface

The plugin also mounts `ctx.officeDocument`, a session-scoped Remote surface over the same seam. Its generated Client namespace is `remote.officeDocument`; mount the package in the Host composition and add `@forwork/office-exchange/remote` to the Client assembly.

The Remote resolves the live Session from the header-derived scope and records `document/*` facts there; the `session-document` plugin folds them into the open set.

- `read(workspaceFileScope, path)` resolves the path from the Session's workspace root, confines it with `assertWithinRoots`, reads it through the filesystem seam under the provider's `maxFileBytes` bound, and returns `{ format, snapshot, preserved, unsupported, version }`. The first read of a document in this process appends `document/open` once; later reads update no fact. The open fact records the workspace-relative locator, so an absolute and a relative spelling of one file share one document identity.
- `save(workspaceFileScope, path, snapshot, format, expectedVersion, attribution)` requires the file to still hold `expectedVersion`, reads its bytes under the byte bound, re-checks the version, and writes from those bytes through the provider. A write that intersects a coupled construct rejects as `office-document/unsupported-feature` with the construct named and appends no fact; a file that changed since the read rejects as `office-document/conflict`, and the write never falls back to an unconditional save. The output is published with a filesystem `replaceIfVersion` guard. An accepted save appends `document/save` with `attribution.actor` (`user` for the editor's save, `agent` for a document tool), the next content version, and the writer's disposition; an agent save inside a turn also records `attribution.turn`. In-memory changes stay in the editor, so no operation steps or inverse are recorded.
- `closeDocument(workspaceFileScope, documentId, reason)` appends `document/close` and drops the process-local open entry. A tab closed after a Host restart still leaves the durable fact.
- `rebindDocument(workspaceFileScope, documentId, toPath)` appends `document/rebind` and keeps the identity; both the old and new paths must resolve inside the workspace root.
- `markDocumentMissing(workspaceFileScope, documentId)` appends `document/missing` once and keeps the document open; a later save, edit, or rebind marks it available again.
- `detectUnsupported(workspaceFileScope, path)` reports loss-prone features without returning content.
- `convertLegacy(workspaceFileScope, path)` returns the converted file path inside `<workspaceRoot>/.dsh/office-tmp/convert-<hash>-<random>/`, so the same Remote can `read` it under the unchanged root check. Each call gets its own directory.
- `writeDraft(workspaceFileScope, documentId, documentPath, contentVersion, snapshot)` persists one recovery draft and returns `{ savedAt }`. The draft lives at `<workspaceRoot>/.dsh/office-tmp/drafts/<sha256(documentId)>.json`; the digest keeps a caller-chosen id from naming a path outside the draft directory. It replaces an existing draft and never writes the document path itself.
- `readDraft(workspaceFileScope, documentId)` returns `{ documentId, path, contentVersion, snapshot, savedAt }`, or `null` when the document has no draft. A malformed draft rejects as `office-document/conversion-failed`.
- `clearDraft(workspaceFileScope, documentId)` removes the draft; clearing a document without a draft is a no-op.
- `completeCommand(workspaceFileScope, requestId, result)` resolves one pending correlated command with the live editor's answer. The provider also mounts `ctx.officeCommandBroker`, the service that owns pending requests, their timeout (`commandTimeoutMs`), and the last observed numeric content version per document. A result for an unknown, late, or other-Session request is ignored and logged, never applied. `registerCommandTarget(workspaceFileScope, documentId, registrationId)` and `unregisterCommandTarget(workspaceFileScope, documentId, registrationId)` record which live tabs can answer: a command for a document with no registered target is answered `document-not-open` immediately, while only a registered but silent target reaches the timeout.

`completeCommand` answers a `document/command` Session event appended by the Host tool layer; this package does not append that event itself. The result union is `read`, `applied`, `pending-edits`, `conflict`, `unsupported-feature`, or `document-not-open`; a read or applied answer carries the live snapshot and numeric content version, and a conflict carries `currentVersion`. The Host writes files only after an `applied` answer.

Drafts are version-bound: the record carries the `contentVersion` the snapshot was captured from, so a restore path can compare it with the file's current version before offering recovery. Both the draft directory and the resolved draft file pass `assertWithinRoots`, so a link that points outside the workspace root is refused rather than written through.

`OfficeDocumentSnapshot` is a JSON projection tagged with the snapshot kind: Univer's native interfaces admit unconstrained custom fields, which a JSON wire cannot carry. The Host validates the projection before serializing and maps every failure onto `office-document/*` Remote codes (`unsupported-format`, `unsupported-feature`, `out-of-scope`, `bound-exceeded`, `conversion-failed`, `conflict`).

### Failure policy

Failures are `OfficeExchangeError` values with a stable `code`. Reads that cannot open a file report `conversion-failed` with the path instead of leaking a Node error. A non-empty `unsupported` list on read means a whole-file regeneration would lose content; an in-place save carries it, and only a coupled construct intersecting the edit refuses. A ZIP64, encrypted, split, or otherwise unsupported OOXML container reports `conversion-failed` and writes nothing. A legacy conversion that exceeds `maxConversionMs` aborts the engine and reports `too-large`, leaving the source file unchanged.

-----

<a id="understand-the-implementation"></a>
## Understand the implementation

<details>
<summary>Implementation internals — click to expand</summary>

`src/service.ts` declares the abstract `OfficeExchange` service, the context augmentation, and the seam vocabulary re-exported by the package entry. `src/types.ts` holds the JSON-safe wire and seam vocabulary; `src/snapshots.ts` holds the Univer snapshot types so a Remote Client never loads the editor core. `src/remote.ts` implements the `officeDocument` Remote surface: Session scope resolution, workspace confinement, version threading, the JSON projection, the workspace-scoped conversion directory, the recovery-draft store, command dispatch and completion, and the remaining save guarantees. `src/commands.ts` owns the `officeCommandBroker` service and its pending-request registry. The Remote also requires the `sessions` service: it appends document facts only to a live Session. `src/detect.ts` classifies preserved constructs and coupled markers. `src/package/` holds the Host-only ZIP/OPC layer: the reader indexes local and central records, the writer copies untouched records and rebuilds the central directory, and `ooxml-package.ts` exposes the part map, dirty set, and no-op fast path. `src/source/` holds the source-span layer: the namespace-aware byte scanner, the unit digests, the sidecar validation, and the splice plan. `src/provider/index.ts` implements `LocalOfficeExchange` on top of the pure `provider/xlsx.ts`, `provider/docx.ts`, `src/detect.ts`, and `provider/legacy.ts` modules: it owns file reads, format dispatch, writer selection, typed failures, caller-directed legacy conversion directories, and the validated `Config` schema. The office editor and agent tools remain consumers; this package registers no tool, and the Remote's only Session writes are the `document/*` facts named above.

</details>

-----

<a id="further-exploration"></a>
## Further Exploration

- [Document group](../README.md) — the Office conversion packages in this group.
- [Office to PDF](../office-to-pdf/README.md) — the LibreOffice conversion provider used for legacy formats.

-----

<a id="model-experience"></a>
## Model Experience

### Exchange seam

#### What the model sees

Nothing directly; agent document tools consume `ctx.officeExchange` and own their model-visible schemas and results.

#### Token effect

Zero direct tokens; the seam returns snapshots and feature lists to its consumers and adds no prompt or result text.

#### KV Cache effect

None; the seam neither assembles nor sends a provider request.

## Known Limitations and Deferred Work

<a id="known-limitations-and-deferred-work"></a>

- `LocalOfficeExchange` reads and writes local paths only; workspace confinement belongs to the tool and session layers that supply the path.
- Legacy conversion requires the LibreOffice kit; an unavailable engine fails the conversion instead of falling back.
- PowerPoint editing is read-only until a slide unit and its conversion are added.
- An in-place save requires the source package. A snapshot without `data.officeExchange.source`, or a caller that supplies no source bytes, uses the legacy regeneration writer; that path rebuilds the package from the model and claims no verbatim entries. A `.docm` or `.xlsm` target refuses that path because regeneration would drop the VBA project and write a macro-free package.
- ZIP64 archives, encrypted entries, split archives, and compression methods other than stored or deflate are refused as `conversion-failed` instead of being rewritten. File-size, cell-count, entry-count, and anchor-payload bounds fail with no write.
- Auxiliary parts are append-only: a required `w:style` is appended to `word/styles.xml`, a required list definition to `word/numbering.xml`, and a required relationship or content-type declaration before its closing tag. `word/settings.xml` is never written. An edited cell style reuses a matching `<xf>` or appends only the needed `<numFmt>`, `<font>`, `<fill>`, `<border>`, and `<xf>` entries to `xl/styles.xml` and updates that collection's `count`; existing component bytes stay untouched. A model style component the writer cannot express, a source `<xf>` extension, a theme color reference, a diagonal border, an empty self-closing style sheet, or a package with no `xl/styles.xml` refuses with the component named. A changed cell hyperlink refuses because worksheet relationships are not regenerated.
- A Word table row or column resize, a Word unit reorder, or an unlocatable or ambiguous source span is refused with the construct named. Independent unmapped properties inside an edited unit are carried: a cell style index, table-cell and row properties, unmapped paragraph and run properties, cell width, and fit-to-text.
- A worksheet row or column shift refuses with the construct named when the sheet carries a merged range, conditional-formatting range, data-validation range, autoFilter range, or Excel-table address range; a declared table part whose relationship, target part, or `ref` cannot be resolved refuses the shift instead of leaving a stale table reference, and a plain sheet shifts by row or cell span operations.
- The saved file is assembled from source local records and published as one whole file through the guarded replace; it is not patched on disk in place.
- Section page borders, page numbering, document grid, line numbering, and vertical text direction are not mapped. Header and footer drawings and fields are not mapped either; header and footer images, like body images, are preserved opaquely rather than rendered, while their lists and hyperlinks use the same mapping as body content. Footnote and endnote bodies are preserved opaquely as whole parts, so their text is not editable. An in-place save keeps every `word/numbering.xml` definition the source carries, including one no `w:num` references; the regeneration writer drops an unreferenced abstract definition because it has no list instance to key. Hyperlink relationship targets are always written as external.
- Table cell fit-text, hide-mark, and text-direction properties, `w:tblCellSpacing`, row `w:gridBefore`/`w:gridAfter` offsets, floating table positions, and table-style conditional formatting are not mapped into the editor model. An in-place save keeps them in unedited content and re-emits the ones independent of an edited unit. Cell `w:insideH`/`w:insideV` borders are used only as fallbacks for an absent outer edge and are written back as explicit cell edges, cell border line-style variants collapse to the model's solid, dotted, or dashed set, and an automatic table width is written back as its resolved grid width.
- Theme parts and theme-derived styles are not mapped: a theme font or color reference without a direct value does not become editable formatting. An in-place save copies theme parts and existing style definitions unchanged, and appends only a style definition it needs; the regeneration writer reconstructs only Normal and the heading styles. A paragraph that names a style keeps its `pStyle` id, and its direct run and paragraph formatting is preserved.
- Workbook row and column outline levels have no Univer model field and are not editable; the regeneration writer drops them, and an in-place save copies the sheet around the edited cell spans. Sheet scroll offsets and print settings have the same disposition. Excel has no distinct "clip" wrap flag, so a cell whose model uses the clip strategy is written as no-wrap, and Excel's `fill` and vertical `distributed`/`justify` alignments map onto the nearest Univer alignment.
- Unique-value, duplicate-value, and average conditional formats have no ExcelJS writer, so the regeneration writer emits equivalent `COUNTIF` or `AVERAGE` expressions. Icon sets outside Excel's built-in set render no icons. ExcelJS does not retain a rule's text attribute, stop-if-true, or data-bar gradient, border, direction, and axis options, so those options are not editable.
- ExcelJS does not retain a hyperlink's screen tip when reading, so an edited link round-trips only its URL and display text; an in-place save of another cell copies the untouched link. Data validation types outside list, whole, decimal, date, time, text length, custom, and any (checkbox and list-multiple) reject on save instead of being dropped.
- Cell images are stored as base64 sheet drawings and written back with their cell anchors; worksheet backgrounds use the same storage. Absolute-anchor pictures and non-picture drawing elements are preserved as `drawing` content and remain uneditable. Media with extensions other than png, jpeg, or gif is carried through, but the regeneration writer writes `image/<extension>` as its content type, which some applications may not recognize.
- `convertLegacy` keeps converted editable files under `<workspaceRoot>/.dsh/office-tmp/`, which the `officeDocument` Remote can read under the same root check. A failed conversion removes the directory it just created; a successful one persists until the user or session deletes it, and nothing removes it automatically yet.
- Recovery drafts accumulate under `<workspaceRoot>/.dsh/office-tmp/drafts/` until `clearDraft` removes them; nothing prunes drafts for documents that are no longer open yet.
- The client reports a missing file through `markDocumentMissing`, but no client trigger detects a move yet: `rebindDocument` is exposed on the Remote surface without an automatic caller. The open set is process-local for deduplication, so a Host restart makes the next read append `document/open` again; the projection keeps the document's position and version.
- The Remote records facts only for a live Session. A read for a Session that is not loaded rejects with `gateway/bad-request` instead of reading without durability.

<a id="dev-note"></a>
### Dev Note

<details>
<summary>Working context for maintainers — click to expand</summary>

License check (2026-11-18): `@univerjs/core` and the four sheet feature plugins (`@univerjs/data-validation`, `@univerjs/drawing`, `@univerjs/sheets-conditional-formatting`, `@univerjs/sheets-data-validation`, `@univerjs/sheets-drawing`, `@univerjs/sheets-hyper-link`) are Apache-2.0 and `exceljs` is MIT. `@univerjs/core` stays a peer dependency because this package imports only its types and emits them in declarations; the sheet feature plugins are mounted by the editor Client. `exceljs` is a runtime dependency of the `.xlsx` provider.

</details>
