---
description: "Embedded Office document editor tab for the right Sidebar, with Client and Host plugin faces."
kind: "package-reference"
---

# @deepseek-ai/dsh-client-ui-office-editor

## Summary

Client plugin package for an embedded Office document editor in the right Sidebar. It reserves the package for a tab type that opens `.docx` and `.xlsx` files, an editor surface, and open/save paths through `ctx.officeExchange`. The plugin currently registers no tab, panel, or editor, and no source module imports Univer yet.

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

The package is not wired into a profile. Mounting it registers nothing and adds no visible control.

The planned editor opens a workspace document in a right-Sidebar tab, edits cells or paragraphs in place, and saves through `ctx.officeExchange`. Spreadsheets and word-processing documents use separate editor presets, and `.pptx` stays in the read-only document preview.

-----

<a id="understand-the-implementation"></a>
## Understand the implementation

<details>
<summary>Implementation internals — click to expand</summary>

`src/index.ts` is the Host companion and defines an empty apply function for the Client loader. `src/client/index.ts` is the browser plugin and currently registers no slot contribution, tab type, service, or store.

`tsconfig.host.json` and `tsconfig.client.json` compile the two plugin faces separately, and `tsdown.config.ts` builds the browser bundle from the emitted Host declarations. Editor code, Univer presets, and tab registration belong to the browser face when they land.

</details>

-----

<a id="further-exploration"></a>
## Further Exploration

- [ui-sidebar-right](../ui-sidebar-right/README.md) — the right Sidebar and its tab-type registry.
- [ui-sidebar-documentpreview](../ui-sidebar-documentpreview/README.md) — the read-only document preview beside the editor.
- [Office Exchange](../../document/office-exchange/README.md) — the conversion capability behind open and save.
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

- The editor is not implemented; the plugin registers no tab type, document panel, or editor preset.
- Univer core and presets are Apache-2.0; the package pins `@univerjs/core`, `@univerjs/presets`, `@univerjs/preset-sheets-core`, and `@univerjs/preset-docs-core` at 1.4.0 for the editor surface.
- The Client injects no services and has no bundle row, so a profile cannot mount it as a working editor yet.

<a id="dev-note"></a>
### Dev Note

<details>
<summary>Working context for maintainers — click to expand</summary>

None.

</details>
