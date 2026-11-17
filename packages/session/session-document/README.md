---
description: "Host document session primitives: stable identity bindings, version-guarded change records, durable document facts, and the open-document projection and prompt section."
kind: "package-reference"
---

# @forwork/session-document

English | [中文](README.zh.md)

## Summary

`dsh-session-document` keeps a document's identity stable across path moves, makes every accepted edit attributable and reversible, and folds durable `document/*` facts into the open-document set. Identity bindings pair a Session-local `DocumentId` with its locator and availability; change records pair an edit with the actor, version transition, and inverse; the `document` projection and prompt section read those facts back. Restore expands the projection into the set a client reopens, and the recovery rule matches a stored draft to its captured file. The base-profile `session-document` Host plugin accesses no filesystem and owns no storage: the Session log is the only durable source.

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

Use this package wherever an open document's path may change, an edit must be undone at the document level, or the agent needs the current open set: the office editor Remote, restore after restart, and the document tools. Base profiles mount the `session-document` plugin, which registers the `document` projection and the open-document prompt section; the fact helpers are usable wherever a live `Session` is at hand. Callers own the filesystem observations; the Session log owns durability.

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

### Record document facts

```text
recordDocumentOpen(session, { documentId, path, format, contentVersion })
recordDocumentSave(session, { documentId, actor, contentVersion, turn? })
recordDocumentRebind(session, { documentId, fromPath, toPath })
recordDocumentMissing(session, { documentId, path })
recordDocumentClose(session, { documentId, reason })
```

Each helper appends one ordinary Session event and returns it. Facts are append-only: a later observation appends a new fact instead of replacing an earlier one. A save whose changes stayed in the editor is `document/save`; it records the author (`user` or `agent`), the published version, and the causing turn when one exists, carries no operation steps, and never fabricates an inverse. An agent edit's steps and inverse stay in `document/edit`.

### Read the open set

The `document` projection holds `{ order, documents }`: the display order and one entry per open document with its identity, current path, format, content version, `open` or `missing` status, and unsupported features. `document/open` inserts or updates an entry, `document/save` and `document/edit` advance its version, `document/rebind` replaces the path, `document/missing` marks it unavailable, and `document/close` removes it. Unknown events keep the same state reference, so no downstream work runs. The wire view is the same value with a schema for client restore.

The plugin's prompt section renders the open set for the assembling Agent's Session, or nothing when there is no Agent or no open document. Each line leads with the entry's exact `documentId`, the identity a document tool call passes back unchanged. Both the projection and the section are disposed with the mounting fiber.

### Restore the open set

```text
const entries = restoreSet(projection)
const decision = recoveryDecision(draft, readResult.version)
```

`restoreSet` returns one entry per open document in the projection's own order, carrying the identity, current path, format, and status. A `missing` document stays in the set: the Session keeps it open with its history, so the caller reports its path instead of opening a dead tab. A closed document is absent from the projection already.

`recoveryDecision` is the one draft rule: `none` without a draft, `offer` when the draft's captured version equals the current file version, and `stale` otherwise. A stale draft is offered only behind an explicit choice; nothing here applies it.

The package also exposes `./types` (the projection's pure types and its `SessionProjectionMap` key) and `./client`, a type-only re-export of `./types`, so a client package binds `useProjection('document')` without importing a Host value module.

-----

<a id="understand-the-implementation"></a>
## Understand the implementation

<details>
<summary>Implementation internals — click to expand</summary>

Both modules are collections of pure functions over immutable values; every function returns a new object and never mutates its inputs.

### Source map

| File | Role |
|---|---|
| [`src/document-identity.ts`](src/document-identity.ts) | `DocumentIdentityBinding`, its create, rebind, missing, and resolve transitions, and the shared `documentIdFor` identity |
| [`src/document-changes.ts`](src/document-changes.ts) | `ChangeRecord`, `ChangeActor`, `OfficeDocumentError`, the version guard, and the undo primitive |
| [`src/document-facts.ts`](src/document-facts.ts) | Append-only writers for the five `document/*` facts |
| [`src/document-projection.ts`](src/document-projection.ts) | The `document` projection fold, state and wire schemas |
| [`src/types.ts`](src/types.ts) | The pure entry/state/view types and the `document` projection-key declaration |
| [`src/client.ts`](src/client.ts) | Type-only re-export of `./types` for client programs |
| [`src/document-restore.ts`](src/document-restore.ts) | `restoreSet`: the ordered, missing-flagged set a client reopens |
| [`src/document-recovery.ts`](src/document-recovery.ts) | `recoveryDecision`: offer, stale, or ignore one stored draft |
| [`src/plugin.ts`](src/plugin.ts) | The `session-document` plugin: projection registration and the open-document prompt section |
| [`src/index.ts`](src/index.ts) | Package entry re-exporting the modules, the prompt renderer, and the plugin's `name`/`inject`/`apply` |

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

### Document transitions

#### What the model sees

When at least one document is open, the prompt section renders a short English list: the open-document count implied by its position order, each document's Session-scoped `<sessionId>:<path>` id, its file name, its format, and a `missing` marker. The id is the exact identity the document tools accept, so the model can address an open document without a separate read. Without an Agent or without an open document the section renders nothing; the document tools own their own model-visible schemas and results.

##### Open-document prompt section

```markdown
Open documents in this session:
1. document-plugin:work/budget.xlsx (budget.xlsx, xlsx)
2. document-plugin:work/report.docx (report.docx, docx, missing)
```

#### Token effect

One line per open document, plus a one-line heading, only while documents are open. Each line repeats the full path in its id, so a line grows with the path length; a file name longer than 40 characters is elided to its first 37 characters and `...`, while the id stays exact. An empty open set adds zero tokens.

#### KV Cache effect

The list changes only when a `document/*` fact changes the open set or an entry's path, format, or status; it keeps a stable position for each document that stays open.

## Known Limitations and Deferred Work

<a id="known-limitations-and-deferred-work"></a>

- **Filesystem observations are caller-supplied.** `rebindDocument` trusts the `exists` flag and `markDocumentMissing` trusts its caller; a registry that probes paths belongs with the consumer, not here. The Office editor reports a missing file from its resource failure, but no client trigger reports a move yet: `rebindDocument` exists on the Remote surface without an automatic caller.
- **Restore and recovery stop at the decision.** `restoreSet` expands the projection and `recoveryDecision` classifies a draft; opening tabs, choosing one pass per Session, reporting missing paths, and clearing a draft after save all belong to the client that owns the tabs.
- **The projection carries no unsupported-feature facts.** `document/open` currently has no field for them, so the entry starts empty; the Exchange read result stays the authority for a save refusal until a fact carries the list.
- **The version guard does not serialize writers.** Callers must pass the live `currentVersion` from one authoritative document owner; the library detects a mismatch but cannot prevent two concurrent calls from reading the same version.
- **Record validity is checked structurally only.** The library requires a non-empty operation and inverse but cannot prove the inverse restores the prior content; the editing provider that produces both is responsible for exact reversibility.

<a id="dev-note"></a>
### Dev Note

<details>
<summary>Working context for maintainers — click to expand</summary>

None.

</details>
