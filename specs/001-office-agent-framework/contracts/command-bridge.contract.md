# Contract: Document Command Bridge

**Owner**: Host `packages/document/tool-office` (requester) + Client `packages/client/ui-office-editor` (executor)
**Purpose**: Let Host-owned agent tools apply one operation to the single live editor unit and receive a correlated result (research R15, FR-020).

## Transport

1. The Host tool validates the session, permitted roots, and the document's expected version.
2. The Host appends a `document/command` session event:
   ```ts
   { requestId: string; documentId: string; operation: DocumentEditOperation; fromVersion: number }
   ```
   The event is additive and `ignorable`; it never carries a unit snapshot.
3. The client subscribes to its session's `document/command` events. While attached, it registers its live target through `registerCommandTarget(scope, documentId, registrationId)` and releases that registration through `unregisterCommandTarget` when it detaches. For a matching open document and a version that matches the live unit, it applies the operation through the adapter, then calls the `officeDocument` Remote:
   ```ts
   completeCommand(scope, requestId, result)
   ```
4. The Host resolves the pending request with the result. A document with no registered target fails `document-not-open` immediately. A request to a registered target that receives no result within the configured timeout (default 5000 ms) fails `timeout` and writes nothing.

## Result

```ts
type CommandResult =
  | { status: 'applied'; snapshot: OfficeDocumentSnapshot }
  | { status: 'pending-edits' }
  | { status: 'conflict'; currentVersion: number }
  | { status: 'unsupported-feature'; features: readonly UnsupportedFeature[] }
  | { status: 'document-not-open' }
```

- `applied` carries the post-command snapshot; the Host serializes it, persists with one version-guarded save (the only version bump), then appends `document/edit` with `actor = agent` and the inverse.
- `pending-edits` means the live unit has unsaved user edits; the Host writes nothing and returns the tool error.
- `conflict` means the live version differs from `fromVersion`; the Host writes nothing.
- `unsupported-feature` means applying or saving would drop content; the Host writes nothing.
- `document-not-open` means no live editor target was registered, or the registered target's unit was gone; the Host writes nothing.

## Read

`document_read` uses the same bridge with a `read` operation; it returns the structure summary and never saves.

## Guarantees

- One accepted request produces at most one save and one `document/edit` fact.
- The client never writes files; only the Host persists.
- A late or duplicate result for an unknown `requestId` is ignored and logged as a diagnostic.
- The bridge is local IPC on desktop and carries no document bytes off-device.
