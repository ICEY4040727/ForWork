---
description: "Host-only document session primitives: stable identity bindings across rename and missing paths, and version-guarded attributable change records with their inverses."
kind: "package-library"
---

# @deepseek-ai/dsh-session-document

## Summary

`dsh-session-document` keeps a document's identity stable while its path moves and makes every accepted edit attributable and reversible. Identity bindings pair a Session-local `DocumentId` with the current locator and its availability; change records pair a structured edit with the actor, the version transition, and the exact inverse. Both are pure values: this library performs no filesystem access, owns no storage, registers no Cordis service, and emits no Session event. The `document` projection, restore, and the document tools feed its results into `document/*` facts.

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

### When to use it

Use this library wherever an open document's path may change or an edit must be undone at the document level: the `document` projection, restore after restart, the office editor, and the document tools. It is not a Cordis plugin and has no profile mount row. Callers own the filesystem observations and the Session log; the library owns only the pure transitions.

### Bind an identity

```text
const binding = createDocumentIdentity(documentId, 'C:/work/report.docx')
const afterMove = rebindDocument(binding, 'C:/work/renamed.docx', at, true)
const afterLoss = markDocumentMissing(afterMove, at)
const found = resolveDocumentBinding(bindings, documentId)
```

`documentId` survives every rebind and every missing observation. `rebindDocument` takes an explicit `exists` flag: the caller reports whether the new path resolved, so a move onto an unavailable path keeps `missing: true` instead of clearing it. A missing binding stays in the registry with its history; it never closes the document. Paths are opaque strings here, so any locator the filesystem seam produces works unchanged.

### Apply and undo a change

```text
const { record, version } = applyChange({
  documentId,
  actor: 'agent',
  currentVersion: liveVersion,
  fromVersion: claimedVersion,
  operation,
  inverse,
  at,
})
const undo = undoChange(record)
```

`applyChange` is the version guard. It owns no storage, so the caller passes the live document version as `currentVersion`; the change is accepted only when `fromVersion` equals it. An accepted change yields a new record and `toVersion = fromVersion + 1`. The record carries the structured steps and their inverse, never a unit snapshot, so document undo is reconstructable from the Session log. `undoChange` returns the inverse steps and the version to restore; `T056` consumes it.

`changeId` defaults to the deterministic `documentId:toVersion`; passed explicitly, it is kept. Every refusal is an `OfficeDocumentError` with a stable code:

- `stale-version` — `fromVersion` is not the live version; the caller must reload or merge.
- `invalid-operation` — the operation or its inverse is empty, so the change could not be applied or reversed.
- `invalid-transition` — a version cannot describe a real document or cannot be advanced safely.

-----

<a id="understand-the-implementation"></a>
## Understand the implementation

<details>
<summary>Implementation internals — click to expand</summary>

Both modules are collections of pure functions over immutable values; every function returns a new object and never mutates its inputs.

### Source map

| File | Role |
|---|---|
| [`src/document-identity.ts`](src/document-identity.ts) | `DocumentIdentityBinding` and its create, rebind, missing, and resolve transitions |
| [`src/document-changes.ts`](src/document-changes.ts) | `ChangeRecord`, `ChangeActor`, `OfficeDocumentError`, the version guard, and the undo primitive |
| [`src/index.ts`](src/index.ts) | Package entry re-exporting both modules |

</details>

-----

<a id="further-exploration"></a>
## Further Exploration

- [Session package map](../README.md) — the persistence, projection, and title packages this library supports.
- [Session projections subsystem](../../../docs/subsystems/session-projection.md) — how the `document` projection folds the emitted `document/*` facts.
- [Document events contract](../../../specs/001-office-agent-framework/contracts/document-events.contract.md) — the event vocabulary these values feed.

-----

<a id="model-experience"></a>
## Model Experience

None directly; the document tools own their model-visible schemas and results, and this library adds no prompt or tool text.

#### KV Cache effect

None; the library assembles no provider request.

## Known Limitations and Deferred Work

<a id="known-limitations-and-deferred-work"></a>

- **Filesystem observations are caller-supplied.** `rebindDocument` trusts the `exists` flag and `markDocumentMissing` trusts its caller; a registry that probes paths belongs with the consumer, not here.
- **The version guard does not serialize writers.** Callers must pass the live `currentVersion` from one authoritative document owner; the library detects a mismatch but cannot prevent two concurrent calls from reading the same version.
- **Record validity is checked structurally only.** The library requires a non-empty operation and inverse but cannot prove the inverse restores the prior content; the editing provider that produces both is responsible for exact reversibility.

<a id="dev-note"></a>
### Dev Note

<details>
<summary>Working context for maintainers — click to expand</summary>

None.

</details>
