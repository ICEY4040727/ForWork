---
description: "Office Exchange capability seam: OOXML-to-unit-snapshot mapping, unsupported-feature detection, and legacy conversion shared by the office editor and agent document tools."
kind: "package-reference"
---

# @forwork/office-exchange

English | [中文](README.zh.md)

## Summary

Define the contract for moving Office files in and out of the editable unit model. The seam names conversion and loss-detection; providers own the ExcelJS, SheetJS, docx, and LibreOffice mechanics. The editor and agent tools consume `ctx.officeExchange` without importing a conversion library.

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
})

const { format, snapshot, unsupported } = await ctx.officeExchange.read(path, signal)
```

`read(path)` returns the detected format, an editable unit snapshot, and every feature that would not survive an edit-and-save. `write(snapshot, format)` encodes the snapshot back to file bytes without writing them. `detectUnsupported(path)` reports loss-prone features without reading content into the editor. `convertLegacy(path, signal?, outputDir?)` turns a `.doc/.xls/.ppt` file into its editable equivalent; an omitted `outputDir` writes into a fresh temporary directory outside any workspace.

Supported formats:

- `.xlsx` and `.docx` are editable: `read` and `write` map them to Univer workbook and document snapshots.
- `.doc`, `.xls`, and `.ppt` must be converted with `convertLegacy` before `read`; reading one directly raises `unsupported-format`.
- `.pptx` is preview-only: `detectUnsupported` inspects it, while `read` and `write` raise `unsupported-format`.

Every failure is an `OfficeExchangeError` with a stable `code`; bounds resolve once at mount and invalid values fail loudly.

`formatOf(path)` classifies a path by extension; `isEditableFormat` separates directly editable `.docx/.xlsx` from formats that require conversion; `isLegacyFormat` names the legacy set.

### Unsupported-feature detection

`read` and `detectUnsupported` report every construct an edit-and-save would drop; a non-empty list refuses the save instead of writing a lossy file.

- Documents: macros, tracked changes, comments, charts, drawings, shapes, headers and footers, embedded objects, inline images (`word/media/`), hyperlinks, lists and numbering (`word/numbering.xml` definitions or `w:numPr`), unsupported run formatting (font, size, color) and paragraph alignment, merged-cell spans (a `w:gridSpan` above one or any `w:vMerge`), and tables nested inside a cell.
- Workbooks: charts, drawings, comments, macros, embedded objects, conditional formatting, data validation, frozen panes, sheet tab colors, and cell hyperlink targets.

Bold, italic, underline, heading styles, number formats, fills, borders, merged ranges, row heights, and column widths round-trip, so detection does not report them.

### Remote surface

The plugin also mounts `ctx.officeDocument`, a session-scoped Remote surface over the same seam. Its generated Client namespace is `remote.officeDocument`; mount the package in the Host composition and add `@forwork/office-exchange/remote` to the Client assembly.

The Remote resolves the live Session from the header-derived scope and records `document/*` facts there; the `session-document` plugin folds them into the open set.

- `read(workspaceFileScope, path)` resolves the path from the Session's workspace root, confines it with `assertWithinRoots`, reads it through the filesystem seam under the provider's `maxFileBytes` bound, and returns `{ format, snapshot, unsupported, version }`. The first read of a document in this process appends `document/open` once; later reads update no fact.
- `save(workspaceFileScope, path, snapshot, format, expectedVersion)` re-detects the current source through the Exchange before writing and rejects as `office-document/unsupported-feature`, naming the first feature, when an edit-and-save would drop it. Otherwise it serializes the snapshot through the provider and publishes it with a filesystem `replaceIfVersion` guard. A file that changed since the read rejects as `office-document/conflict`; the write never falls back to an unconditional save. An accepted save appends `document/save` with `actor: 'user'` and the next content version; the in-memory user changes stay in the editor, so no operation steps or inverse are recorded.
- `closeDocument(workspaceFileScope, documentId, reason)` appends `document/close` and drops the process-local open entry. A tab closed after a Host restart still leaves the durable fact.
- `rebindDocument(workspaceFileScope, documentId, toPath)` appends `document/rebind` and keeps the identity; both the old and new paths must resolve inside the workspace root.
- `markDocumentMissing(workspaceFileScope, documentId)` appends `document/missing` once and keeps the document open; a later save, edit, or rebind marks it available again.
- `detectUnsupported(workspaceFileScope, path)` reports loss-prone features without returning content.
- `convertLegacy(workspaceFileScope, path)` returns the converted file path inside `<workspaceRoot>/.dsh/office-tmp/convert-<hash>-<random>/`, so the same Remote can `read` it under the unchanged root check. Each call gets its own directory.
- `writeDraft(workspaceFileScope, documentId, documentPath, contentVersion, snapshot)` persists one recovery draft and returns `{ savedAt }`. The draft lives at `<workspaceRoot>/.dsh/office-tmp/drafts/<sha256(documentId)>.json`; the digest keeps a caller-chosen id from naming a path outside the draft directory. It replaces an existing draft and never writes the document path itself.
- `readDraft(workspaceFileScope, documentId)` returns `{ documentId, path, contentVersion, snapshot, savedAt }`, or `null` when the document has no draft. A malformed draft rejects as `office-document/conversion-failed`.
- `clearDraft(workspaceFileScope, documentId)` removes the draft; clearing a document without a draft is a no-op.
- `completeCommand(workspaceFileScope, requestId, result)` resolves one pending correlated command with the live editor's answer. The provider also mounts `ctx.officeCommandBroker`, the service that owns pending requests, their timeout (`commandTimeoutMs`), and the last observed numeric content version per document. A result for an unknown, late, or other-Session request is ignored and logged, never applied.

`completeCommand` answers a `document/command` Session event appended by the Host tool layer; this package does not append that event itself. The result union is `read`, `applied`, `pending-edits`, `conflict`, `unsupported-feature`, or `document-not-open`; a read or applied answer carries the live snapshot and numeric content version, and a conflict carries `currentVersion`. The Host writes files only after an `applied` answer.

Drafts are version-bound: the record carries the `contentVersion` the snapshot was captured from, so a restore path can compare it with the file's current version before offering recovery. Both the draft directory and the resolved draft file pass `assertWithinRoots`, so a link that points outside the workspace root is refused rather than written through.

`OfficeDocumentSnapshot` is a JSON projection tagged with the snapshot kind: Univer's native interfaces admit unconstrained custom fields, which a JSON wire cannot carry. The Host validates the projection before serializing and maps every failure onto `office-document/*` Remote codes (`unsupported-format`, `unsupported-feature`, `out-of-scope`, `bound-exceeded`, `conversion-failed`, `conflict`).

### Failure policy

Failures are `OfficeExchangeError` values with a stable `code`. Reads that cannot open a file report `conversion-failed` with the path instead of leaking a Node error. A non-empty `unsupported` list on read means the write path must refuse rather than save a lossy file; the product names the offending feature and offers an alternative. A legacy conversion that exceeds `maxConversionMs` aborts the engine and reports `too-large`, leaving the source file unchanged.

-----

<a id="understand-the-implementation"></a>
## Understand the implementation

<details>
<summary>Implementation internals — click to expand</summary>

`src/service.ts` declares the abstract `OfficeExchange` service, the context augmentation, and the seam vocabulary re-exported by the package entry. `src/types.ts` holds the JSON-safe wire and seam vocabulary; `src/snapshots.ts` holds the Univer snapshot types so a Remote Client never loads the editor core. `src/remote.ts` implements the `officeDocument` Remote surface: Session scope resolution, workspace confinement, version threading, the JSON projection, the workspace-scoped conversion directory, the recovery-draft store, command dispatch and completion, and the remaining save guarantees. `src/commands.ts` owns the `officeCommandBroker` service and its pending-request registry. The Remote also requires the `sessions` service: it appends document facts only to a live Session. `src/provider/index.ts` implements `LocalOfficeExchange` on top of the pure `provider/xlsx.ts`, `provider/docx.ts`, `provider/detect.ts`, and `provider/legacy.ts` modules: it owns file reads, format dispatch, typed read failures, caller-directed legacy conversion directories, and the validated `Config` schema. The office editor and agent tools remain consumers; this package registers no tool, and the Remote's only Session writes are the `document/*` facts named above.

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
- `convertLegacy` keeps converted editable files under `<workspaceRoot>/.dsh/office-tmp/`, which the `officeDocument` Remote can read under the same root check. A failed conversion removes the directory it just created; a successful one persists until the user or session deletes it, and nothing removes it automatically yet.
- Recovery drafts accumulate under `<workspaceRoot>/.dsh/office-tmp/drafts/` until `clearDraft` removes them; nothing prunes drafts for documents that are no longer open yet.
- The client reports a missing file through `markDocumentMissing`, but no client trigger detects a move yet: `rebindDocument` is exposed on the Remote surface without an automatic caller. The open set is process-local for deduplication, so a Host restart makes the next read append `document/open` again; the projection keeps the document's position and version.
- The Remote records facts only for a live Session. A read for a Session that is not loaded rejects with `gateway/bad-request` instead of reading without durability.

<a id="dev-note"></a>
### Dev Note

<details>
<summary>Working context for maintainers 鈥?click to expand</summary>

License check (2026-11-14): `@univerjs/core` is Apache-2.0 and `exceljs` is MIT. `@univerjs/core` stays a peer dependency because this package imports only its types and emits them in declarations; `exceljs` is a runtime dependency of the `.xlsx` provider.

</details>
