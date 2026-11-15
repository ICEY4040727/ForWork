# Contract: `document/*` Session Events

**Owner**: `packages/core/session` (event vocabulary), `packages/session` (projection)
**Rule**: model-visible ⟺ logged (Constitution III)

## Events

| Event | Payload (required unless marked) | When |
|---|---|---|
| `document/open` | `documentId`, `path`, `format`, `contentVersion` | A document becomes open |
| `document/edit` | `documentId`, `actor`, `fromVersion`, `toVersion`, `operation`, `inverse` | An agent change is accepted |
| `document/save` | `documentId`, `actor` (`user` or `agent`), `contentVersion` | A save is accepted without exposed operation steps |
| `document/rebind` | `documentId`, `fromPath`, `toPath` | The file moved or was renamed |
| `document/missing` | `documentId`, `path` | The file no longer resolves |
| `document/close` | `documentId`, `reason` | The document is closed |

Optional fields: `turn` (owning turn), `at` (log-assigned timestamp).

## Invariants

- `documentId` is stable for the life of the document and unique within the session.
- `document/edit` is written only after its write passes the version guard; its `toVersion = fromVersion + 1`.
- `document/edit` carries the structured `operation` and its `inverse`; raw unit snapshots are never persisted in the log.
- `document/save` records a user save whose in-memory changes stay inside the editor: it carries the author and published version, never operation steps, and never a fabricated inverse.
- `document/rebind` and `document/missing` preserve `documentId` and all prior history.
- No event deletes or rewrites earlier history; the log is append-only.
- An event that reaches a model request MUST be reconstructable from the log.

## Projection

The `document` projection folds these events into the open-document set (id, path, format, version, status, unsupported features). It is the single source for the sidebar panel and the agent's open-document context. `document/open` inserts or updates an entry, `document/save` and `document/edit` advance its version, `document/rebind` replaces its path, `document/missing` marks it unavailable, and `document/close` removes it.

## Compatibility

Adding a member to the closed event set requires a `SESSION_FORMAT_VERSION` decision per the repository policy; same-version additions must remain safely readable by older readers.
