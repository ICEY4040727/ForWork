---
description: "Model-facing document tools for reading and editing Office files through the shared Office Exchange capability."
kind: "package-reference"
---

# @deepseek-ai/dsh-tool-office

## Summary

Host plugin package for the agent's Office document tools. It reserves the package for the structured read, cell-edit, paragraph-edit, and conversion operations that consume `ctx.officeExchange` and report file-level failures to the model. The plugin currently registers no tool, so the package contributes nothing to a model request.

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

Mount the plugin in a Host composition that also provides `ctx.officeExchange`. The current entry point registers nothing, so mounting it has no model-visible effect and exposes no tool name, schema, or result text.

The planned tools read document structure, edit cells and paragraphs by stable identity, and convert legacy formats. Each operation will carry its own model-visible schema and structured failure codes.

-----

<a id="understand-the-implementation"></a>
## Understand the implementation

<details>
<summary>Implementation internals — click to expand</summary>

`src/index.ts` exposes one plugin function and augments no Cordis context member. The package provides no service and creates no session events.

Document tool implementations will register through the Host tool registry and route conversion, unsupported-feature detection, and write-back through `ctx.officeExchange`. Filesystem writes remain bounded by the consuming operation's authority rules.

</details>

-----

<a id="further-exploration"></a>
## Further Exploration

- [Office Exchange](../office-exchange/README.md) — the conversion capability these tools will consume.
- [Document group](../README.md) — the Office conversion packages in this group.
- [Tool catalog](../../../docs/tool-catalog.md) — generated model-facing tool reference.

-----

<a id="model-experience"></a>
## Model Experience

### Document tools

#### What the model sees

Nothing yet. The plugin's `apply` registers no tool schema, prompt section, or tool result text.

#### Token effect

Zero direct tokens; no model request contains output from this package.

#### KV Cache effect

None; the package neither assembles nor sends a provider request.

## Known Limitations and Deferred Work

<a id="known-limitations-and-deferred-work"></a>

- No tools are registered; the plugin entry point is the complete current behavior.
- Model-visible schemas, result text, and structured conflict or unsupported-feature failures arrive with the tool implementation.
- Bundle and product-composition registration, plus the `ctx.officeExchange` provider mount, remain with the composing packages.

<a id="dev-note"></a>
### Dev Note

<details>
<summary>Working context for maintainers — click to expand</summary>

None.

</details>
