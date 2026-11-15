---
description: "Embedded Office document editor tab for the right Sidebar, with Client and Host plugin faces."
kind: "package-reference"
---

# @deepseek-ai/dsh-client-ui-office-editor

English | [中文](README.zh.md)

## Summary

Client plugin package for an embedded Office document editor in the right Sidebar. It registers the `office` tab type for Session-scoped `.docx` and `.xlsx` addresses, reads and saves through the `officeDocument` Remote, and mounts the Univer Sheets or Docs core preset behind an adapter seam. Edits are captured as idle-debounced recovery drafts, and reopening the document offers a stored draft instead of applying it silently. Opening a legacy `.doc`, `.xls`, or `.ppt` source converts it and opens the copy in a new tab; an editable copy opens here with a notice naming the original, and a converted presentation stays in the read-only preview.

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

The Web bundle mounts the package. Opening a `.docx` or `.xlsx` file in the right Sidebar claims the address with the `office` type and opens an editable sheet or paragraph surface. Spreadsheets and word-processing documents use separate Univer presets behind the adapter. Saving writes through the `officeDocument` Remote at the version the read observed; a stale version surfaces as a conflict. A file that reports an unsupported feature cannot be saved, and the panel names the features that would be lost.

An agent command for the open document reaches this tab through the Session's `document/command` feed. The tab answers with the live snapshot and its numeric content version, applies edit steps through the mounted adapter, refuses a stale edit as a conflict, refuses while user edits are unsaved, and ignores commands for another document.

A legacy `.doc`, `.xls`, or `.ppt` source is converted through the same Remote and reopened as a copy in that tab's place; the original is never written. The converted tab shows a dismissible banner naming the source. The body also lists the Session's open Office documents and focuses one when clicked.

An external file change reloads a clean tab and prompts before discarding unsaved edits. Closing a dirty tab is vetoed by the Sidebar's cleanup-failure path, and the panel then offers save-and-close, discard-and-close, or cancel.

While the unit is dirty, an edit schedules a recovery draft after 2 s without another edit. Reopening the document reads that draft once and offers recovery; a draft captured from the current file version is offered directly, while one from an older version is labelled stale and is still applied only when the user chooses it. A successful save and a discard both remove the stored draft.

The first Office body that mounts in a Session reads the `document` projection and reopens the documents the Session still holds, in projection order. One pass runs per Session per window; a document that already has an Office tab is not opened again, and a document the projection reports missing is listed in a dismissible notice instead of a dead tab. Each restored document then follows the ordinary open path above, so its stored draft is offered by the same recovery surface.

-----

<a id="understand-the-implementation"></a>
## Understand the implementation

<details>
<summary>Implementation internals — click to expand</summary>

`src/index.ts` is the Host companion and defines an empty apply function for the Client loader. `src/client/index.ts` is the browser plugin: it registers `officeDefinition()` into `ctx.sidebarRightTabs`, the dictionaries, the body under `sidebar.right.pane.tab`, the chip title under `sidebar.right.pane.tab.title`, and the dirty-close veto through `ctx.sidebarRight.registerCloseHandler`. The definition keeps the body mounted across tab switches so a live edit survives.

`definition.ts` declares the `office` kind, its `.docx`/`.xlsx` patterns plus the legacy `.doc`/`.xls`/`.ppt` sources it converts, and the Session-file veto. `OfficePanel.tsx` reads the file metadata through `useResource`, calls `useOfficeDocument` against the injected `officeDocument` Remote, and mounts the adapter only after a successful read. `useOfficeDocument.ts` owns read/save state, the expected-version guard, failure classification, the external-change rule, the recovery offer, and the legacy conversion that reopens the copy through the tab's own `openResource`. `command-bridge.ts` subscribes one tab to the Session's appended `document/command` facts, resolves each request exactly once against the live document identity, numeric content version, and dirty state, applies edit steps through the adapter handle, and settles the Host call through `officeDocument.completeCommand`. `drafts.ts` owns the idle-debounced scheduler: each panel creates one instance, an edit replaces its pending candidate, `flush` writes immediately, `cancel` waits for an in-flight write, and every refusal or rejection reaches the hook's draft failure state. `changes.ts` owns the dirty and close-warning snapshot store the panel and close handler share; `notices.ts` owns the converted-copy records keyed by the converted tab's content id; `restore.ts` owns the per-Session restore record and the hook that opens the missing documents in projection order. `DocumentPanel.tsx` lists the Session's open `office` tabs from `sidebarRight.openTabs`. `target.ts` maps addresses to Remote paths, names the shared document identity, and builds the converted copy's address with the workspace-path helper. `face.ts` binds the Remote, the adapter, the command bridge, those stores, and the open-tab inventory for the slot registration.

`adapter.ts` owns the runtime boundary: `OfficeAdapter`, `OfficeEditorHandle`, `createAdapter(runtimeFactory?)`, and a controllable test adapter. The default factory dynamically imports `editor/sheets.ts` or `editor/docs.ts`; those modules construct `createUniver` with `UniverSheetsCorePreset` or `UniverDocsCorePreset`, their en-US locale default exports and CSS, create the unit from the wire snapshot, and report mutations from the preset's command signal. No module imports Univer at top level, so jsdom tests never load it. Undo and redo are the preset header's standard controls; the handle forwards the same native history.

`tsconfig.host.json` and `tsconfig.client.json` compile the two plugin faces separately, and `tsdown.config.ts` emits the browser bundle plus the lazily requested sheets and docs chunks.

</details>

-----

<a id="further-exploration"></a>
## Further Exploration

- [ui-sidebar-right](../ui-sidebar-right/README.md) — the right Sidebar and its tab-type registry.
- [ui-sidebar-documentpreview](../ui-sidebar-documentpreview/README.md) — the read-only document preview beside the editor.
- [Client group](../README.md) — the browser packages in this group.

-----

<a id="model-experience"></a>
## Model Experience

### Editor surface

#### What the model sees

Nothing. The browser plugin's `apply` registers no tool schema, prompt section, or tool result text.

#### Token effect

Zero direct tokens; the package contributes nothing to a model request.

#### KV Cache effect

None; the package neither assembles nor sends a provider request.

## Known Limitations and Deferred Work

<a id="known-limitations-and-deferred-work"></a>

- A real Univer mount needs a browser (canvas, ResizeObserver, preset styles); the unit lane replaces the runtime factory and the assembled browser smoke owns the real mount.
- The close warning is an in-panel banner shown after the Sidebar vetoes removal, not a native modal. The body stays mounted through hiding, so the pending prompt appears when its tab is activated.
- The editor UI ships with the preset's en-US locale only.
- A converted `.ppt` copy is a `.pptx`, which stays read-only in the document preview, so presentations never open in this editor.
- Converted-copy notices come from this window's plugin state, and the open-document list comes from the Sidebar's open-tab inventory.
- Restore runs from the first Office body mounted in a Session, because the Sidebar opens resources only through an existing tab's `openResource` action (or its own internal per-Session path). The Sidebar persists each Session's tab layout, so a restart normally mounts those bodies and restore reconciles the durable set; a Session with no Office tab open at all has no body to trigger it.
- The one-pass restore record is keyed by Session and lives for this window only; a projection that arrives after the pass does not reopen anything.
- Recovery drafts are stored by the `officeDocument` Remote under `<workspaceRoot>/.dsh/office-tmp/drafts/` and are read only when a document opens. The offer is per tab open, so an external-change reload does not check for a draft again.
- The live content version is tab-local: it starts at 0 on mount and advances for each accepted save. A reload triggered by the Host's own save can therefore make the next command report a conflict, after which the agent's read resynchronizes it.
- Univer core and presets are Apache-2.0; the package pins `@univerjs/core`, `@univerjs/presets`, `@univerjs/preset-sheets-core`, and `@univerjs/preset-docs-core` at 1.4.0 for the editor surface.

<a id="dev-note"></a>
### Dev Note

<details>
<summary>Working context for maintainers — click to expand</summary>

None.

</details>
