---
description: "Model-facing document tools for reading and editing Office files through the live editor command bridge."
kind: "package-reference"
---

# @forwork/tool-office

English | [中文](README.zh.md)

## Summary

Host plugin package for the agent's Office document tools. It registers `document_read`, `document_edit_cells`, `document_edit_paragraph`, and `document_convert`. Reads and edits go through the correlated `document/command` bridge to the live editor; the Host performs the single version-guarded save for every accepted edit and records `document/edit` with its inverse.

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

Mount the plugin in a Host composition that also provides `ctx.tools` and `ctx.officeDocument`:

```ts
import * as ToolOffice from '@forwork/tool-office'

declare const ctx: import('@forwork/cordis').Context

await ctx.plugin(ToolOffice, { summaryCells: 200, maxEditSteps: 2_000 })
```

`ctx.officeDocument` must be mounted with its provider, because the tools dispatch commands through the provider's `officeCommandBroker`. A composition that mounts the Remote without the broker still reads files, but every command refuses as `conversion-failed` instead of editing a detached copy.

The model-visible tools are:

- `document_read({ documentId })` — reads the live unit's structure (sheets, cells, formulas, paragraphs, tables) and never saves.
- `document_edit_cells({ documentId, fromVersion, operations })` — replaces cell values or formulas, then saves once.
- `document_edit_paragraph({ documentId, fromVersion, operations })` — replaces paragraphs, then saves once.
- `document_convert({ path })` — converts a legacy `.doc` or `.xls` file to an editable copy inside the workspace.

A document id is Session-scoped: `<sessionId>:<path>`, the identity the Office editor already uses for drafts. The path is resolved against the Session's workspace root and confined with `assertWithinRoots` before a command is sent.

`fromVersion` is the document's numeric content version: `0` when the document opens, plus one for every accepted save. `document_read` returns the live value; an edit whose caller version differs is refused as `conflict` with `currentVersion`.

Every refusal is a structured result with a stable `code`: `conflict`, `unsupported-feature`, `out-of-scope`, `document-missing`, `document-not-open`, `pending-edits`, `bound-exceeded`, `timeout`, `invalid-request`, or `conversion-failed`.

-----

<a id="understand-the-implementation"></a>
## Understand the implementation

<details>
<summary>Implementation internals — click to expand</summary>

`src/index.ts` registers the four tool definitions and builds the Host runtime from `ctx.officeDocument`; the deployment caps come from the validated plugin `Config`. `src/read.ts` dispatches a `{ kind: 'read' }` command and summarizes the returned snapshot through `src/summary.ts`. `src/edit.ts` owns the shared edit path: read the file at the identity's workspace-relative path so the editor tab and the open fact keep one document id, compute the inverse from the pre-edit snapshot, dispatch `{ kind: 'edit', steps }`, then save the returned snapshot with the filesystem version observed before the command and the agent's open turn, and append one `document/edit` fact with `actor: 'agent'`. `src/edit-cells.ts` and `src/edit-paragraph.ts` validate their operation lists and build typed steps; `src/convert.ts` is Host-side. `src/errors.ts` maps Remote and Exchange failures onto the model-facing codes.

The command event carries no snapshot. A refused, unanswered, or timed-out command writes nothing and appends no edit fact, so only an accepted save can change durable document content.

</details>

-----

<a id="further-exploration"></a>
## Further Exploration

- [Office Exchange](../office-exchange/README.md) — the conversion capability, Remote surface, and command broker these tools consume.
- [Document group](../README.md) — the Office conversion packages in this group.
- [Tool catalog](../../../docs/tool-catalog.md) — generated model-facing tool reference.

-----

<a id="model-experience"></a>
## Model Experience

### Document tools

#### What the model sees

Four tool schemas. `document_read` returns a bounded structure summary with sheet and cell counts, a row-major cell preview, formulas, paragraphs, and table dimensions. `document_edit_cells` and `document_edit_paragraph` return the new `contentVersion` and the operations the editor applied. `document_convert` returns the converted path and editable format. Refusals return `status: "error"`, a stable `code`, and an actionable message.

#### Token effect

The read summary is the only variable-size model input this package produces. `summaryCells` (default 200) bounds the cell preview while counts describe the complete unit, so a large workbook never floods a request; a truncated summary is marked `truncated`.

#### KV Cache effect

None; the package neither assembles nor sends a provider request.

## Known Limitations and Deferred Work

<a id="known-limitations-and-deferred-work"></a>

- The stable document id currently equals `<sessionId>:<path>`. The `document` projection (US3) owns the stable-id mapping and the model-facing open-document set; until it lands, a model must obtain the id from the editor or a prior read.
- The tools expose cell and paragraph edits. The `document/edit` step union also carries row, column, formula, and style steps: `document_edit_cells` emits formula steps, while no tool surface emits row/column insert or delete or style operations yet.
- `document_convert` returns a converted copy under the workspace; nothing prunes `.dsh/office-tmp/` conversion directories automatically.
- Editing requires a live editor instance. Headless sessions refuse with `document-not-open` rather than editing and saving a detached file.

<a id="dev-note"></a>
### Dev Note

<details>
<summary>Working context for maintainers — click to expand</summary>

Tests mount the real Host composition (`LocalFileSystem`, `LocalOfficeExchange`, `ToolRuntime`, `AgentLoop`) and spy only `officeDocument.dispatchCommand`, so the save path, event append, and file confinement run for real. `tests/harness.ts` owns that composition.

</details>
