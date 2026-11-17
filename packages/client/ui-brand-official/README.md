---
description: "Official ForWork sidebar branding for the official client build."
kind: "package-reference"
---

# @forwork/client-ui-brand-official

English | [中文](README.zh.md)

## Summary

This package fills the official sidebar brand slots for an `official` client build. The shipped Web and Desktop compositions load it by default; other build profiles keep their existing branding. The blank-home mark and its hover leaf come from the `dsh-client-ui-conversation` fallback, not from this package. The plugin adds no model requests and no temporary background resources.

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

The shipped Web and Desktop compositions already mount this plugin. An `official` client build registers the ForWork sidebar artwork; no extra installation or enablement is needed.

### Choosing the profile

`DSH_CLIENT_BUILD_PROFILE` selects which brand renders. An `official` build shows the official sidebar mark and name; any other value leaves the shell fallbacks — the door mark and the local-build label — in place. The conversation hero shows the door mark from `dsh-client-ui-conversation` regardless of profile, because that fallback is already the official mark. The plugin still loads and validates in both cases; only the sidebar artwork registration is profile-gated.

### Home frame layer

The package does not occupy `conversation.hero.frame`. The slot remains declared by `dsh-client-ui-conversation` and is available to another bundle; the blank home keeps the static door mark and its hover leaf. Removing the previous click-triggered canvas decoration also removed the upstream whale atlas and its asset route.

### Replacing the brand

A deployment with its own identity leaves this package out and composes another package that occupies the sidebar slots and the hero-mark slot, which this package leaves on its fallback. Occupying a slot is the only composition route; the artwork identity is not configurable.

-----

<a id="understand-the-implementation"></a>
## Understand the implementation

<details>
<summary>Implementation internals — click to expand</summary>

The two sidebar occupants install as one declaration-aware registration set: nested `ctx.slots.inject()` calls wait on the sidebar declaration, so the set works whether this row activates before or after the declarer, withdraws both occupants when the declaration collapses, and leaves no partial brand mix during HMR. The browser half is [`src/client/index.ts`](src/client/index.ts); the node half is an empty Loader seat. The browser title is a build-environment concern (`DSH_CLIENT_TITLE`), outside the slot system.

</details>

-----

<a id="further-exploration"></a>
## Further Exploration

Read these pages when the brand surface is not enough. They move from the slots this package occupies to the shell that renders them.

- [ui-sidebar](../ui-sidebar/README.md) — declares `sidebar.brand.mark` and `sidebar.brand.name` and renders their fallbacks.
- [ui-conversation](../ui-conversation/README.md) — declares `conversation.hero.brand.mark` and the available `conversation.hero.frame` layer.
- [Web client architecture](../../../docs/subsystems/web-client.md) — how browser plugin rows load and register slots.

-----

<a id="model-experience"></a>
## Model Experience

None, as the package contributes browser presentation only; nothing here reaches a model request.

#### KV Cache effect

None; this package neither assembles nor sends a provider request.

## Known Limitations and Deferred Work

<a id="known-limitations-and-deferred-work"></a>

These limits define how brand presentation is supplied. They are current package constraints, not a brand-design comparison or a task backlog.

- **One occupant set** — alternative presentation belongs in another Cordis package occupying the same slots.
- **No home decoration** — the package leaves `conversation.hero.frame` empty; the blank home shows the static door mark fallback.
- **The browser title is independent** — `DSH_CLIENT_TITLE` selects title text at build time rather than through a UI slot.

<a id="dev-note"></a>
### Dev Note

<details>
<summary>Working context for maintainers — click to expand</summary>

None.

</details>
