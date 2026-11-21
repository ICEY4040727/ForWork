# Contract: Document Save Disposition

**Feature**: 002-in-place-ooxml-save

**Event**: `document/save`

## Payload

```ts
export interface DocumentSaveEventData {
  readonly documentId: DocumentId
  readonly actor: DocumentSaveActor
  readonly contentVersion: number
  readonly turn?: number
  /** Optional preservation summary; absent for saves that carry no span plan. */
  readonly disposition?: DocumentSaveDisposition
}

export interface DocumentSaveDisposition {
  readonly schemaVersion: 1
  readonly verbatimCount: number
  readonly spliced: { readonly count: number; readonly names: readonly string[] }
  readonly added: { readonly count: number; readonly names: readonly string[] }
  readonly removed: { readonly count: number; readonly names: readonly string[] }
}
```

## Rules

1. **Optional, additive.** The property is optional. Older readers must keep the event, project `contentVersion`, and ignore the new property. `SESSION_FORMAT_VERSION` stays 4; the envelope `ignorable` flag is not used and does not apply to nested payload changes.
2. **Exact counts.** `count` is the exact number of entries in that status, even when `names` is truncated.
3. **Deterministic names.** `names` are sorted ASCII and capped at 256. Truncation never fails the save.
4. **Truthful source.** The disposition is produced by the writer that assembled the package. No caller may synthesize it.
5. **Only accepted saves.** A refused save appends no `document/save` and therefore no disposition.
6. **Legacy saves.** A save through the legacy regeneration writer carries `disposition` with `schemaVersion: 1`, `verbatimCount: 0`, and counts reported as spliced/added/removed according to that writer's own accounting; it must not claim span identity.
7. **No projection dependency.** The projection ignores `disposition` unless a later change explicitly exposes it. If it is exposed, the projection entry, zod schema, `sameEntry`, and `stateVersion` are updated together and a separate compatibility decision is recorded.

## Verification

- `packages/document/office-exchange/tests/remote.spec.ts` asserts the exact save event data; it must include the disposition for the new path.
- `packages/session/session-document/tests/**` covers folding an event with and without the property.
- The persistence acknowledgement for `event:document/save` records the optional-property decision, and `pnpm run doc-sync` regenerates the schema, catalog, and known-event-type inventory.
- SDK projections: no checked-in scenario currently emits `document/save`. The plan adds one TypeScript and one Python scenario that perform an office save, or records explicitly that no scenario exists and no expected output changes. It must not claim an SDK diff that is not present.

## Compatibility statement

Adding an optional event-body property is a same-version change under `docs/persistence-changes/README.md`: an older reader preserves the save's required meaning (document id, actor, content version) and safely ignores the optional subtree. Making any field required, changing a type, or removing a field is a different decision and is out of scope for this contract.
