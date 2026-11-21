# Office Exchange

English | [中文](office-exchange.zh.md)

The [document package family](../../packages/document/README.md) moves Office files between disk bytes and the editable unit snapshots consumed by the office editor and agent document tools. [office-exchange](../../packages/document/office-exchange) declares the `ctx.officeExchange` capability seam, supplies the `LocalOfficeExchange` provider, and hosts the session-scoped `officeDocument` Remote. The package registers no model-facing tool and appends no Session event. The [office editor Client](../../packages/client/ui-office-editor/README.md) reads and saves through `remote.officeDocument`, and the model-facing [office document tools](../../packages/document/tool-office/README.md) consume the seam and the Remote; [Office editing](office-editing.md) describes the subsystem end to end.

## Ownership

| Owner | Responsibility |
|---|---|
| [office-exchange](../../packages/document/office-exchange) | `ctx.officeExchange`: read, write, unsupported-feature detection, and legacy conversion; the `officeDocument` Remote |
| [api/remotes](../../packages/api/remotes/README.md) | Mounts the generated `remote.officeDocument` Client namespace |
| [base bundle](../../packages/bundle/base/README.md) | Mounts `office-exchange` with deployment bounds |

## Formats and unit snapshots

`OfficeFormat` is `docx`, `docm`, `xlsx`, `xlsm`, `pptx`, `doc`, `xls`, or `ppt`. The editable set is `EditableOfficeFormat` (`docx`, `docm`, `xlsx`, `xlsm`); the legacy set is `LegacyOfficeFormat` (`doc`, `xls`, `ppt`).

| Format | Operations |
|---|---|
| `.docx`, `.docm`, `.xlsx`, `.xlsm` | Editable: `read` maps file bytes to a `UnitSnapshot`; `write` splices the edited units back into the source package. An in-place save keeps a macro-enabled file's VBA project and macro-enabled main content type. |
| `.doc`, `.xls`, `.ppt` | Legacy: `convertLegacy` produces an editable file; `read` on the original rejects with `unsupported-format` |
| `.pptx` | Preview-only: `detectUnsupported` inspects it; `read` and `write` reject with `unsupported-format` |

`UnitSnapshot` is the editors' native serializable snapshot, a discriminated union of `{ kind: 'workbook', data: IWorkbookData }` and `{ kind: 'document', data: IDocumentData }`. `read` returns `ExchangeReadResult` with the detected `format`, the parsed `snapshot`, and every `UnsupportedFeature` a whole-file regeneration would drop. Each `UnsupportedFeature` carries a `kind` — `macro`, `tracked-changes`, `comment`, `chart`, `drawing`, `shape`, `formula-recalc`, `embedded-object`, or `other` — and an optional human-readable `detail`. A non-empty list describes what the source carries; it never refuses a save by itself.

`formatOf(path)` classifies a path by extension, `isEditableFormat` narrows to `docx`/`docm`/`xlsx`/`xlsm`, and `isLegacyFormat` narrows to `doc`/`xls`/`ppt`.

## In-place save

`ctx.officeExchange.write(snapshot, format, source?, signal?)` returns `OfficeWriteResult` (`bytes`, `disposition`) without writing a file. With a snapshot sidecar and a source package, the writer assembles the output from the source bytes: every entry the edit does not change is copied from its source ZIP local record, so compressed data, CRC, timestamps, extra fields, compression method, entry order, and data descriptors stay identical. A save that changes no unit returns the source bytes unchanged.

Inside a part the edit changes, only the edited unit's XML span is replaced: a paragraph, a table cell, or a worksheet `<c>` element. A body-level content control is anchored to its whole `w:sdt` element, so editing that unit refuses while a sibling edit leaves it untouched. Every byte outside the edited spans is copied from the source, and a worksheet row or column insert or delete is expressed as span operations rather than a whole-sheet rewrite; a shift on a sheet that carries a merged range, conditional-formatting range, data-validation range, autoFilter range, or Excel-table address range is refused with that construct named, because the range coordinates would not move with the shifted rows or columns. A required style, numbering definition, relationship, or content-type declaration is appended before its part's closing tag instead of regenerating the part; an edited cell style reuses a matching `<xf>` or appends the needed style components to `xl/styles.xml` and updates their counts. `word/settings.xml` is never written, and a changed worksheet string is written as an inline string so `xl/sharedStrings.xml` stays untouched.

`read` attaches the source sidecar at `data.officeExchange.source` inside the editable data object. It records `schemaVersion: 1`, the `format`, the source `packageDigest` and `byteLength`, the `anchorMode` (`per-unit` or `sheet-summary`), the `sourceParts`, and either `anchors` or `sheets`. Each anchor names the unit `kind`, its `part`, a model `key`, the `start` and `end` byte offsets in the uncompressed part, and a digest of the canonical unit projection. The sidecar survives `unit.save()`, the JSON wire, and a recovery draft, so a save that follows an editor round trip still finds its anchors.

`maxSourceAnchors` (100,000 by default) and `maxSourceAnchorBytes` (8 MiB by default) bound the sidecar. A workbook above the per-cell anchor count records `sheet-summary` entries and rebuilds per-cell anchors from the source at save. A source whose units cannot be anchored, or whose anchor payload stays above the budget, is read without a sidecar and therefore saves through the regeneration writer. When a save's source no longer matches the sidecar's `packageDigest` or `byteLength`, the writer re-anchors that source before planning, which makes a second save of a just-written file work without a re-read. A source that cannot be re-anchored is refused as `conversion-failed`; the filesystem version guard remains authoritative for conflicts.

A snapshot without `data.officeExchange.source`, or a caller that supplies no source package, uses the legacy regeneration writer. That writer rebuilds the package from the editable model and returns a disposition that claims no verbatim entries. A `.docm` or `.xlsm` target refuses that path because regeneration would write a macro-free package without the VBA project.

## Preservation and refusal

`detectPreserved(bytes, config?)` classifies a source package without deciding whether a save may proceed. `PreservedConstructKind` names what an in-place save carries verbatim outside an edited unit: `macro`, `chart`, `pivot-table`, `excel-table`, `defined-names`, `calc-chain`, `comment`, `embedded-object`, `drawing`, `shape`, `theme`, `font-table`, `web-settings`, `custom-xml`, `external-link`, `field`, `bookmark`, `content-control`, and `document-properties`. Presence alone never refuses a save.

`CoupledConstructKind` names the constructs whose meaning or anchors depend on the content an edit changes: `tracked-change`, `comment-range`, `content-control-wrap`, `field-owned-text`, `rich-text-shared-string`, `shared-formula`, `array-formula`, and `ambiguous-alternate-content`. The writer plans against the edited unit spans and returns a refusal naming the construct and its part or byte offset when one intersects them. A changed worksheet hyperlink also refuses because the worksheet relationship part is not regenerated. A worksheet row or column shift on a sheet carrying a merged range, conditional-formatting `sqref`, data-validation `sqref`, an autoFilter range, or an Excel-table address range refuses with the construct named, because those ranges are not shifted. A `.docm` or `.xlsm` save without an anchored source refuses as `unsupported-feature` instead of regenerating a macro-free package. A refusal writes nothing, records no save fact, and leaves the source bytes unchanged. This differs from a preserved class: the file contains it, but the user may still edit elsewhere.

Unmapped properties inside an edited span that do not depend on the changed content are merged into the spliced unit instead of refusing: a cell's `s` style index, `w:tcPr`/`w:trPr` children, unmapped `w:pPr`/`w:rPr` children, cell width, and fit-to-text. A construct is coupled only when the edit would change its meaning or invalidate its anchor.

## Seam operations and failures

`ctx.officeExchange.read(path, signal?)` returns an `ExchangeReadResult`. `ctx.officeExchange.write(snapshot, format, source?, signal?)` returns `OfficeWriteResult` without writing a file. `detectUnsupported(path, signal?)` reports loss-prone features without returning content. `detectPreserved(bytes, config?)` reports the preserved classes and coupled markers for bytes the caller already holds. `convertLegacy(path, signal?, outputDir?)` returns the path of the editable file; an omitted `outputDir` writes into a fresh temporary directory outside any workspace.

`OfficeExchangeConfig` bounds every provider operation: `maxFileBytes` (25 MiB by default), `maxCells` (2,000,000), `maxConversionMs` (60,000 ms), `commandTimeoutMs` (5,000 ms), `maxSourceAnchors` (100,000), and `maxSourceAnchorBytes` (8 MiB). The provider resolves the bounds once at mount and rejects a non-positive or non-integer value with `RangeError` instead of truncating work. A legacy conversion that exceeds `maxConversionMs` is aborted and reported as `too-large`, leaving the source untouched. Every operation failure is an `OfficeExchangeError` whose stable `code` is `unsupported-format`, `unsupported-feature`, `too-large`, `conversion-failed`, `not-implemented`, or `out-of-scope`.

## The officeDocument Remote

The Remote carries `OfficeDocumentSnapshot`, a JSON projection of a unit snapshot: `{ kind: 'workbook' | 'document', data: JsonValue }`. Univer's native interfaces admit unconstrained custom fields that JSON cannot carry, so the Host validates the projection before serializing and restores the native snapshot for a save; a non-JSON value rejects as `conversion-failed`.

| Method | Behavior |
|---|---|
| `read(workspaceFileScope, path)` | Resolves `path` from the Session workspace root, confines it to that root, reads it through the filesystem seam under `maxBytes`, and returns `OfficeDocumentReadResult` (`format`, `snapshot`, `preserved?`, `unsupported`, `version`); the open fact records the workspace-relative locator |
| `save(workspaceFileScope, path, snapshot, format, expectedVersion, attribution)` | Reads the source at `expectedVersion`, writes from those bytes, and publishes the result with a filesystem `replaceIfVersion` guard; returns `OfficeDocumentSaveResult` (`version`, optional `disposition`) and records `DocumentSaveAttribution` (`actor`, optional `turn`) plus the disposition on `document/save` |
| `detectUnsupported(workspaceFileScope, path)` | Reports loss-prone features without returning content |
| `convertLegacy(workspaceFileScope, path)` | Returns the path of the converted editable file in a fresh workspace-scoped directory |

`read` records the file's opaque content version and re-stats after parsing; a file that changed while it was read rejects instead of returning a snapshot bound to a stale version. `save` never falls back to an unconditional write: when the file changed after the read, the filesystem reports `FS_STALE_VERSION` and the Remote raises `office-document/conflict`. A `format` that does not match the snapshot kind, or a non-editable format, also rejects rather than writing a partial file. A write that intersects a coupled construct rejects as `office-document/unsupported-feature` with the construct named, and appends no save fact.

`SaveDisposition` travels on the save result and the `document/save` fact. It carries `schemaVersion: 1`, an exact `verbatimCount`, and one `{ count, names }` group for `spliced`, `added`, and `removed` entries. Counts are exact; names are a sorted ASCII prefix capped at 256 per group. The property is optional, so a save fact recorded before dispositions existed, or a legacy regeneration, reads as absent rather than invalid.

The Remote maps seam and filesystem failures onto stable Remote codes: `unsupported-format`, `unsupported-feature`, `out-of-scope`, `bound-exceeded`, `conversion-failed`, and `conflict`, each carrying the offending `path`. An aborted caller signal rejects as `gateway/cancelled`. `convertLegacy` creates `<workspaceRoot>/.dsh/office-tmp/convert-<hash>-<random>/` and converts into it, so the same Remote can read the result under the unchanged root check. Each call gets its own directory; a failed conversion removes that directory, and a successful one persists until the user or Session deletes it.

## Mount and composition

`LocalOfficeExchange` validates its `Config` and mounts `OfficeDocumentRemote` with `maxBytes` set to `maxFileBytes`. The [base bundle](../../packages/bundle/base/README.md) mounts the provider in an `office-exchange` row with `maxFileBytes: 26214400`, `maxCells: 2000000`, and `maxConversionMs: 60000`. The Client assembly mounts the generated descriptor from `@forwork/office-exchange/remote` through [api/remotes](../../packages/api/remotes/README.md), which exposes it as `remote.officeDocument`.

<!-- BEGIN GENERATED cordis-surface (gen-cordis-catalog.ts) — do not edit between markers -->

<a id="cordis-surface"></a>

## Cordis API

Generated from source by `scripts/gen-cordis-catalog.ts` (verified fresh by `pnpm run verify-cordis-catalog` in doc-sync; regenerate with `pnpm run gen-cordis-catalog`) — the language sides differ only in locale-specific paired document paths. Signature blocks use a `ts cordis-catalog` fence and keep the original source JSDoc; dispatch modes are defined in the [primer](../cordis-primer.md#dispatch-modes), and the framework-inherited `ctx` API lives in [cordis-api/inherited.md](../cordis-api/inherited.md).

<a id="ctxofficecommandbroker--officecommandbroker"></a>

### `ctx.officeCommandBroker` — `OfficeCommandBroker`

Owns pending document commands, the live target registrations, and the last observed content version per document. One instance serves every Session in its composition; a result is accepted only for the Session that recorded the request.

```ts cordis-catalog
/**
 * Append one `document/command` request and wait for the live editor's
 * answer.
 * A document with no registered live target is answered as
 * `document-not-open` without appending a command or starting the timeout.
 * @param request - session, document, operation, and expected version.
 * @param signal - caller cancellation; aborting withdraws the pending request.
 * @returns the editor's result, `document-not-open` when no target is registered, or
 * `timeout` when a registered target did not answer in time.
 * @throws the reason carried by an aborted `signal`, or an append failure.
 */
async request(request: DocumentCommandRequest, signal: AbortSignal): Promise<OfficeCommandOutcome>

/**
 * Resolve one pending request with the live editor's result.
 *
 * A result for an unknown request id, a request from another Session, or a
 * request that already timed out is ignored and reported as a diagnostic;
 * late answers never change durable state.
 * @param scope - Session identity that produced the result.
 * @param requestId - correlation id copied from the command event.
 * @param result - the live editor's answer.
 */
complete(scope: WorkspaceFileScope, requestId: string, result: OfficeCommandResult): void

/**
 * Register one live editor target that can answer commands for a document.
 *
 * A registration is identified by the caller's id so an attach after a detach
 * for the same document does not retire the newer target.
 * @param documentId - stable identity of the document.
 * @param registrationId - caller-minted identity of this attachment.
 */
attachTarget(documentId: DocumentId, registrationId: string): void

/**
 * Release one live editor target registration.
 * @param documentId - stable identity of the document.
 * @param registrationId - identity passed to {@link attachTarget}.
 */
detachTarget(documentId: DocumentId, registrationId: string): void

/**
 * Whether any live editor target is registered for one document.
 * @param documentId - stable identity of the document.
 * @returns `true` when at least one attachment can answer commands.
 */
hasTarget(documentId: DocumentId): boolean

/**
 * Read the last content version reported for one document.
 * @param documentId - stable identity of the document.
 * @returns the last reported version, or `0` before any command reported one.
 */
versionOf(documentId: DocumentId): number

/**
 * Record the content version one document currently holds.
 * @param documentId - stable identity of the document.
 * @param version - non-negative version reported by the live editor.
 */
remember(documentId: DocumentId, version: number): void
```

Types: [DocumentId](core.md)

Source: [`packages/document/office-exchange/src/commands.ts`](../../packages/document/office-exchange/src/commands.ts)

<a id="ctxofficedocument--officedocumentremote"></a>

### `ctx.officeDocument` — `OfficeDocumentRemote`

Host Remote service backing the generated `remote.officeDocument` namespace. Reads return a JSON snapshot bound to the file's observed version; saves replace only that version, so a concurrent external edit is reported as a conflict rather than overwritten.

The service is the producer of `document/*` facts for the editor: a first read appends `document/open` once per Session, a successful save appends `document/save`, and the lifecycle endpoints append their matching facts. The open set is owned by this service per Session; no module-level state is used.

```ts cordis-catalog
/**
 * Read an editable Office file into a JSON snapshot bound to its version.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param path - absolute or workspace-relative path; it must resolve inside the workspace root.
 * The recorded document identity uses the same workspace-relative locator the editor tab uses,
 * so two spellings of one file never open two documents.
 * @param signal - caller cancellation.
 * @returns the detected format, snapshot, preserved construct classes,
 * unsupported features, and observed version.
 * @throws RemoteError when the path is out of scope, the format is not editable, or the read fails.
 */
@Remote async read( workspaceFileScope: WorkspaceFileScope, path: string, signal: AbortSignal, ): Promise<OfficeDocumentReadResult>

/**
 * Serialize a snapshot and replace the file only at the caller's version.
 *
 * The file is read and stat-verified at `expectedVersion` before the writer
 * runs, so the bytes the writer receives are the bytes the caller read. A
 * preserved construct never refuses the save by itself: the writer returns a
 * typed refusal only when the edit intersects a construct whose meaning or
 * anchors depend on the changed content. A refusal writes nothing and records
 * no fact; the final `replaceIfVersion` write refuses a change that happened
 * during serialization.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param path - absolute or workspace-relative path; it must resolve inside the workspace root.
 * @param snapshot - snapshot returned by {@link read}, optionally modified.
 * @param format - target editable format; it must match the snapshot kind.
 * @param expectedVersion - version returned by {@link read}; a changed file is refused.
 * @param attribution - who caused the save and the agent turn, when one caused it.
 * @param signal - caller cancellation.
 * @returns the version the guarded write published and the writer's disposition.
 * @throws RemoteError with `office-document/unsupported-feature` when the edit
 * intersects a coupled construct, naming that construct in the message.
 * @throws RemoteError with `office-document/conflict` when the file changed after the read.
 */
@Remote async save( workspaceFileScope: WorkspaceFileScope, path: string, snapshot: OfficeDocumentSnapshot, format: OfficeFormat, expectedVersion: string, attribution: DocumentSaveAttribution, signal: AbortSignal, ): Promise<OfficeDocumentSaveResult>

/**
 * Record that an open document left the editor.
 *
 * The fact is appended even when this process no longer tracks the document,
 * so a tab closed after a Host restart still leaves a durable close. History
 * is never deleted; the open set is derived by folding the log.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param documentId - Session-scoped identity of the closing document.
 * @param reason - producer-supplied closure reason; the projection stores it unread.
 * @param signal - caller cancellation.
 * @throws RemoteError when the identity belongs to another Session or no live Session exists.
 */
@Remote closeDocument( workspaceFileScope: WorkspaceFileScope, documentId: string, reason: string, signal: AbortSignal, ): Promise<void>

/**
 * Record that an open document's file moved or was renamed.
 *
 * The identity is preserved: only the locator changes, and the new path must
 * resolve inside the Session workspace root.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param documentId - Session-scoped identity retained across the move.
 * @param toPath - new path; it must resolve inside the workspace root.
 * @param signal - caller cancellation.
 * @throws RemoteError when the identity belongs to another Session, either path is out of scope, or no live Session exists.
 */
@Remote async rebindDocument( workspaceFileScope: WorkspaceFileScope, documentId: string, toPath: string, signal: AbortSignal, ): Promise<void>

/**
 * Record that an open document's file no longer resolves.
 *
 * The document stays open with its identity and history; a later save, edit,
 * or rebind marks it available again.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param documentId - Session-scoped identity that stays open.
 * @param signal - caller cancellation.
 * @throws RemoteError when the identity belongs to another Session or no live Session exists.
 */
@Remote markDocumentMissing( workspaceFileScope: WorkspaceFileScope, documentId: string, signal: AbortSignal, ): Promise<void>

/**
 * Report the features a whole-file regeneration would drop without returning content.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param path - absolute or workspace-relative path; it must resolve inside the workspace root.
 * @param signal - caller cancellation.
 * @returns every detected unsupported feature; empty when regeneration loses no mapped content.
 * @throws RemoteError when the path is out of scope or its format cannot be inspected.
 */
@Remote async detectUnsupported( workspaceFileScope: WorkspaceFileScope, path: string, signal: AbortSignal, ): Promise<readonly UnsupportedFeature[]>

/**
 * Convert a legacy binary Office file to an editable one.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param path - absolute or workspace-relative legacy path; it must resolve inside the workspace root.
 * @param signal - caller cancellation.
 * @returns the path of the converted editable file inside the workspace root.
 * @throws RemoteError when the path is out of scope or conversion fails.
 */
@Remote async convertLegacy(workspaceFileScope: WorkspaceFileScope, path: string, signal: AbortSignal): Promise<string>

/**
 * Persist one recovery draft for an open document.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param documentId - caller-owned identity of the open document; it selects the draft file.
 * @param documentPath - document path the draft is captured for; it is recorded, never written.
 * @param contentVersion - file version the snapshot was captured from.
 * @param snapshot - wire snapshot of the editable content.
 * @param signal - caller cancellation.
 * @returns the time the draft was written.
 * @throws RemoteError when the draft store is out of scope or cannot be written.
 */
@Remote async writeDraft( workspaceFileScope: WorkspaceFileScope, documentId: string, documentPath: string, contentVersion: string, snapshot: OfficeDocumentSnapshot, signal: AbortSignal, ): Promise<OfficeDocumentDraftWriteResult>

/**
 * Read the recovery draft persisted for one open document.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param documentId - caller-owned identity of the open document; it selects the draft file.
 * @param signal - caller cancellation.
 * @returns the stored draft, or `null` when the document has no draft.
 * @throws RemoteError when the draft store is out of scope or its content is malformed.
 */
@Remote async readDraft( workspaceFileScope: WorkspaceFileScope, documentId: string, signal: AbortSignal, ): Promise<OfficeDocumentDraft | null>

/**
 * Remove the recovery draft persisted for one open document.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param documentId - caller-owned identity of the open document; it selects the draft file.
 * @param signal - caller cancellation.
 * @throws RemoteError when the draft store is out of scope or its entry cannot be removed.
 */
@Remote async clearDraft(workspaceFileScope: WorkspaceFileScope, documentId: string, signal: AbortSignal): Promise<void>

/**
 * Answer one pending document command with the live editor's result.
 *
 * A result for an unknown request id is ignored and logged by the broker;
 * the call itself succeeds so a late or duplicate answer never becomes a
 * transport error.
 * @param workspaceFileScope - Session identity that ran the command on the wire.
 * @param requestId - correlation id copied from the `document/command` event.
 * @param result - the live editor's status and, for read or applied commands, content.
 * @param signal - caller cancellation; an aborted call drops the answer.
 * @throws RemoteError with `gateway/bad-request` when a reported version is not a non-negative safe integer.
 */
@Remote completeCommand( workspaceFileScope: WorkspaceFileScope, requestId: string, result: OfficeCommandResult, signal: AbortSignal, ): Promise<void>

/**
 * Register one live editor target for a document.
 *
 * The renderer calls this when its command bridge attaches and releases the
 * matching registration when that bridge detaches. A command dispatched while
 * no registration exists is answered as `document-not-open` instead of
 * waiting for the timeout.
 * @param workspaceFileScope - Session identity that owns the document.
 * @param documentId - Session-scoped identity of the document.
 * @param registrationId - caller-minted identity of this attachment.
 * @param signal - caller cancellation; an aborted call registers nothing.
 * @throws RemoteError when the identity belongs to another Session or no live Session exists.
 */
@Remote registerCommandTarget( workspaceFileScope: WorkspaceFileScope, documentId: string, registrationId: string, signal: AbortSignal, ): Promise<void>

/**
 * Release one live editor target registration.
 * @param workspaceFileScope - Session identity that owns the document.
 * @param documentId - Session-scoped identity of the document.
 * @param registrationId - identity passed to {@link registerCommandTarget}.
 * @param signal - caller cancellation; an aborted call releases nothing.
 * @throws RemoteError when the identity belongs to another Session or no live Session exists.
 */
@Remote unregisterCommandTarget( workspaceFileScope: WorkspaceFileScope, documentId: string, registrationId: string, signal: AbortSignal, ): Promise<void>

/**
 * Dispatch one correlated command from a Host document tool.
 *
 * The broker appends `document/command` and waits for the live editor; this
 * method is the Host-only entry point behind the tool's command executor and
 * is not a Remote endpoint.
 * @param request - session, document, operation, and expected version.
 * @param signal - caller cancellation; aborting withdraws the pending request.
 * @returns the editor's result, `document-not-open` without a registered target, or
 * `timeout` when a registered target did not answer in time.
 * @throws RemoteError with `gateway/internal` when the broker is not mounted.
 */
async dispatchCommand(request: DocumentCommandRequest, signal: AbortSignal): Promise<OfficeCommandOutcome>
```

Source: [`packages/document/office-exchange/src/remote.ts`](../../packages/document/office-exchange/src/remote.ts)

<a id="ctxofficeexchange--officeexchange-abstract-seam"></a>

### `ctx.officeExchange` — `OfficeExchange` (abstract seam)

Service Definition for Office file interchange. A provider implements these operations for the deployment's conversion libraries; consumers depend on the seam only.

```ts cordis-catalog
/**
 * Read an Office file into an editable unit snapshot.
 * @param path - Absolute or workspace-relative file path.
 * @param signal - Optional cancellation.
 * @returns The detected format, snapshot, and unsupported features.
 */
abstract read(path: string, signal?: AbortSignal): Promise<ExchangeReadResult>

/**
 * Serialize a unit snapshot back to Office file bytes.
 *
 * `source` carries the bytes the caller read at its expected version so an
 * in-place writer can splice them. A snapshot without the source sidecar, or
 * a caller that supplies no source, uses the regeneration writer.
 * @param snapshot - Editable content to serialize.
 * @param format - Target Office format.
 * @param source - Source package bytes read at the caller's expected version;
 * omitted for a snapshot that has no source file.
 * @param signal - Optional cancellation.
 * @returns The encoded file bytes and what the save carried or changed.
 */
abstract write( snapshot: UnitSnapshot, format: OfficeFormat, source?: SourcePackage, signal?: AbortSignal, ): Promise<OfficeWriteResult>

/**
 * Report features a whole-file regeneration would drop.
 * @param path - File path to inspect.
 * @param signal - Optional cancellation.
 * @returns Detected unsupported features; empty when regeneration loses no mapped content.
 */
abstract detectUnsupported(path: string, signal?: AbortSignal): Promise<readonly UnsupportedFeature[]>

/**
 * Convert a legacy binary Office file to an editable format.
 * @param path - Legacy file path.
 * @param signal - Optional cancellation.
 * @param outputDir - Directory receiving the converted file; omitted uses a
 * fresh temporary directory outside any workspace.
 * @returns Path of the converted editable file.
 */
abstract convertLegacy(path: string, signal?: AbortSignal, outputDir?: string): Promise<string>
```

Source: [`packages/document/office-exchange/src/service.ts`](../../packages/document/office-exchange/src/service.ts)
<!-- END GENERATED cordis-surface -->
