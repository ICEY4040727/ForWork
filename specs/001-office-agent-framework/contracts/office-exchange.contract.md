# Contract: `ctx.officeExchange`

> **Superseded in part by [002](../../002-in-place-ooxml-save/contracts/office-exchange-write.contract.md).** The `write` signature, its result, and the file-level refusal condition below describe the original 001 surface. The in-place write contract owns them now; the other operations on this page remain current.

**Service key**: `officeExchange`
**Owner**: `packages/document/office-exchange` (Service Definition; providers implement)

## Operations

```ts
read(path: string, signal?: AbortSignal): Promise<ExchangeReadResult>
write(snapshot: UnitSnapshot, format: OfficeFormat, source?: SourcePackage, signal?: AbortSignal): Promise<OfficeWriteResult>
detectUnsupported(path: string, signal?: AbortSignal): Promise<readonly UnsupportedFeature[]>
convertLegacy(path: string, signal?: AbortSignal): Promise<string>
```

`UnitSnapshot.data` is the editor core's native snapshot: `IWorkbookData` for workbooks and `IDocumentData` for documents (Option B). The provider owns the OOXML mapping to and from those native types; no project-owned neutral model is introduced. `write` returns `{ bytes, disposition }`; `source` carries the bytes read at the caller's expected version and selects the in-place writer when the snapshot also carries `data.officeExchange.source`.

## Preconditions

- `path` resolves inside the session's permitted roots; an unresolved or out-of-root path is rejected before any read.
- `write` is called only with a snapshot bound to the document's current content version.
- Conversion bounds (bytes, cells, time) are configured by the provider; exceeding a bound fails rather than truncates.

## Postconditions

- `read` returns the detected `format`, a `UnitSnapshot`, and every detected `unsupported` feature. It never mutates the source file.
- `write` never drops content it does not explicitly change. **Superseded by [002](../../002-in-place-ooxml-save/contracts/office-exchange-write.contract.md):** a non-empty `unsupported` list no longer refuses a write by itself; refusal is limited to a coupled construct intersecting the edit or an entry the writer cannot preserve.
- A successful write is the only point that advances the document's persisted version; in-memory commands do not.
- `write` output is a complete, valid file of the requested format.
- `convertLegacy` leaves the source untouched and returns the path of a converted editable file.

## Errors

`OfficeExchangeError` with a stable `code`:

| Code | Meaning |
|---|---|
| `unsupported-format` | Extension or content is not a supported office format |
| `unsupported-feature` | The file contains a feature the editable model cannot preserve |
| `too-large` | A configured bound was exceeded |
| `conversion-failed` | Parsing, conversion, or serialization failed |
| `not-implemented` | The provider does not implement an operation |
| `out-of-scope` | The target resolves outside the session's permitted write roots |

## Guarantees

- Failure names the offending feature or bound; no partial or lossy file is returned as success.
- The returned bytes are deterministic for the same snapshot and format.
