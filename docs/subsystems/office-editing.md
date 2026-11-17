# Office editing

English | [中文](office-editing.zh.md)

Office editing turns a Word or Excel file opened in a Session into a live edit unit, keeps that unit's identity and versions across edits, and lets the agent read and edit the unit through tools rather than through the file. [office-exchange](../../packages/document/office-exchange/README.md) converts between files and unit snapshots and hosts the `officeDocument` Remote; [session-document](../../packages/session/session-document/README.md) owns document identity and the `document/*` Session facts; [tool-office](../../packages/document/tool-office/README.md) registers the agent tools; [ui-office-editor](../../packages/client/ui-office-editor/README.md) renders the right-Sidebar tab that holds the live unit. [Office Exchange](office-exchange.md) owns the seam vocabulary, format classification, Remote methods, and mount bounds; this page owns how the participants cooperate.

## Ownership

| Owner | Responsibility |
|---|---|
| [office-exchange](../../packages/document/office-exchange/README.md) | `ctx.officeExchange` conversion seam, `LocalOfficeExchange`, the `officeDocument` Remote, and the `officeCommandBroker` command registry |
| [session-document](../../packages/session/session-document/README.md) | `DocumentId` bindings, version-guarded change records, `document/*` fact writers, the `document` projection, and the restore and recovery rules |
| [tool-office](../../packages/document/tool-office/README.md) | `document_read`, `document_edit_cells`, `document_edit_paragraph`, and `document_convert` |
| [ui-office-editor](../../packages/client/ui-office-editor/README.md) | the `office` Sidebar tab, the Univer adapter, save and conflict handling, draft scheduling, and command answers |
| [ui-sidebar-right](../../packages/client/ui-sidebar-right/README.md) | tab-type registration, per-Session tab layout, and the unsaved-close veto |
| [ui-sidebar-documentpreview](../../packages/client/ui-sidebar-documentpreview/README.md) | the read-only preview, including `.pptx` through a Host-rendered PDF |

## Formats and the editable set

`.docx` and `.xlsx` are editable: the editor mounts them as Univer workbook or document units, and saves serialize them back through the Exchange provider. A legacy `.doc`, `.xls`, or `.ppt` source is converted first and the copy opens; the original is never written. `.pptx` never becomes an editable unit: `detectUnsupported` can inspect a presentation, `read` and `write` reject it with `unsupported-format`, and the read-only document preview renders it as a Host PDF.

A read reports every `UnsupportedFeature` an edit-and-save would drop. While that list is non-empty, the editor and the tools refuse to save instead of writing a lossy file. [Office Exchange](office-exchange.md) defines the format unions, unit snapshots, and feature kinds.

## Identity, versions, and guarded saves

A document's Session-scoped identity is `<sessionId>:<path>`, produced by `documentIdFor`. The identity survives `document/rebind` and `document/missing`; the path is a locator that may change, and a missing file leaves the document in the open set with its history.

Two version values answer different questions. The `officeDocument` read returns an opaque filesystem version, and `save` passes it back as `expectedVersion`: the filesystem `replaceIfVersion` guard publishes the write only while the file still holds that version, otherwise the Remote raises `office-document/conflict`. The numeric `contentVersion` starts at `0` when a document opens and advances by one for each accepted save; an agent command compares its `fromVersion` against it, and the `document` projection reports it.

The editor reloads a clean tab when the file changes externally and prompts before discarding edits on a dirty tab. A save has no unconditional fallback, and a document with unsupported features cannot be saved from either the editor or the tools.

## Data paths

| Flow | Sequence |
|---|---|
| open | `officeDocument.read` returns the format, snapshot, unsupported features, and file version to the editor adapter; the first read of a document in this process appends `document/open` once |
| user save | the dirty unit calls `officeDocument.save` with its read version and the `user` actor; the guarded write appends `document/save` |
| agent read | `document_read` dispatches a `read` command through `officeCommandBroker`; the live tab answers with its snapshot and numeric version, and the tool returns a bounded summary |
| agent edit | the tool reads the file at the identity's workspace-relative path to compute the inverse, dispatches the steps as one command, and after an `applied` answer performs one guarded save with the `agent` actor and open turn, which appends `document/save`, and appends `document/edit` with the steps and inverse, then reports the new numeric version |
| legacy open | `convertLegacy` writes an editable copy under `.dsh/office-tmp/`; the tab opens the copy with a notice, and the original stays unchanged |
| draft recovery | an idle dirty unit calls `writeDraft`; the next open calls `readDraft` and applies the recovery decision |
| restore | the cached `document` projection supplies `restoreSet`, and the tabs reopen in projection order; a missing document is reported in a notice |

## Session facts and the document projection

Seven `document/*` events carry the durable record: `open`, `save`, `edit`, `rebind`, `missing`, `close`, and `command`. They are append-only. `document/command` records a request, not a change; every accepted write appends `document/save` with its actor and any causing turn, and an accepted agent change also appends `document/edit` with its exact inverse. The open identity is the workspace-relative locator, so an absolute and a relative spelling of one file are one document. `session-document` defines the payloads and the fact writers.

The `document` projection folds those facts into `{ order, documents }`: the display order and one entry per open document with its identity, current path, format, numeric content version, `open` or `missing` status, and unsupported features. `document/open` inserts or updates an entry, save and edit advance its version, rebind replaces the path, missing marks it unavailable, and close removes it; an unrelated event returns the same state reference. The plugin also renders the open set as a prompt section and exposes `restoreSet` and `recoveryDecision` to the Client. [Session Projection](session-projection.md) owns the projection seam, and the [persistence catalog](../persistence-catalog.md) carries the generated event types.

## The Sidebar editor

The `office` tab type claims Session-scoped `dsh-resource://file/session/<sessionId>/<path>` addresses whose basename ends in `.docx`, `.xlsx`, `.doc`, `.xls`, or `.ppt`; `.pptx` is deliberately absent. The tab keeps its body mounted across tab switches, so a live edit survives navigation. The body reads through `officeDocument.read` and mounts the Univer Sheets or Docs preset behind an adapter.

Dirty state drives three behaviors: an edit schedules a recovery draft after 2 s without another edit; a Sidebar close is vetoed, and the panel then offers save-and-close, discard-and-close, or cancel; an external change prompts before discarding. `DocumentPanel` lists the Session's open Office tabs and focuses one when clicked. A legacy open converts the source and reopens the copy in the tab's place with a dismissible notice naming the original.

Restore runs once per Session per window, triggered either by the first Office body that mounts or by the Session becoming the main-view Session. It reads the cached `document` projection and opens each document in projection order through the Sidebar's per-Session open path; a document the projection reports missing appears in a dismissible notice instead of a dead tab.

## The agent command bridge

The read and edit tools never change the live unit directly. A read dispatches a `read` command; an edit first reads the file through `officeDocument.read` to compute the inverse, then dispatches the steps. The provider's `officeCommandBroker` appends one `document/command` and waits for the live tab, which answers exactly once with `read`, `applied`, `pending-edits`, `conflict`, `unsupported-feature`, or `document-not-open`. A request that times out writes nothing. The broker accepts an answer only for the Session that recorded the request, and it ignores and logs an unknown, late, or foreign answer.

An accepted agent edit saves once. The tool verifies that the editor's numeric version is `fromVersion + 1`, saves the returned snapshot with the file version observed before the command, records the save as `actor: 'agent'` with the open turn, and appends `document/edit` with the steps and the inverse. A refusal leaves the file and the log unchanged. [tool-office](../../packages/document/tool-office/README.md) lists the model-visible result codes.

## Drafts and recovery

While a unit is dirty, each edit replaces one pending draft candidate, and the scheduler writes it after 2 s without another edit. `writeDraft` stores the wire snapshot and the file version it was captured from at `<workspaceRoot>/.dsh/office-tmp/drafts/<sha256(documentId)>.json`; it never writes the document path. Reopening the document reads the draft once: a draft from the current file version is offered directly, and an older one is labelled stale and applied only when the user chooses it. A successful save and a discard both remove the draft, and nothing prunes drafts for documents that stay closed.

## Legacy conversion

`convertLegacy` runs inside the Session workspace. It creates `<workspaceRoot>/.dsh/office-tmp/convert-<hash>-<random>/` and writes the editable copy there, so the same Remote can read the result under the unchanged root check. Each call gets its own directory, and a failed conversion removes it. The original file is never modified. The editor opens an editable copy in the tab with a notice; a converted `.ppt` becomes a `.pptx`, which stays in the read-only preview.

## Composition

The [base bundle](../../packages/bundle/base/README.md) mounts `office-exchange`, `session-document`, and `tool-office`. The [Web application bundle](../../packages/bundle/web-app/README.md) mounts `ui-office-editor`, and [api/remotes](../../packages/api/remotes/README.md) exposes the generated `remote.officeDocument` namespace. `tool-office` requires the provider's command broker: a composition that mounts the Remote without the broker still reads files but refuses every command as `conversion-failed`. The generated service reference for `ctx.officeExchange`, `ctx.officeDocument`, and `ctx.officeCommandBroker` lives on [Office Exchange](office-exchange.md#cordis-surface).
