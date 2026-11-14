# Contract: Agent Document Tools

**Owner**: `packages/document/tool-office` (Consumer of `ctx.officeExchange`)
**Transport**: registered on the harness tool registry; model-visible schema and results

## Tools

| Tool | Input | Output | Refuses when |
|---|---|---|---|
| `document_read` | `documentId` | structure summary: sheets/paragraphs/tables/cells/formulas | document closed or missing |
| `document_edit_cells` | `documentId`, `fromVersion`, operations | new `contentVersion`, applied summary | stale version, out-of-root, unsupported feature |
| `document_edit_paragraph` | `documentId`, `fromVersion`, operations | new `contentVersion`, applied summary | stale version, unsupported formatting |
| `document_convert` | `path` | converted `path` | unsupported source, bound exceeded |

## Preconditions

- `documentId` refers to a document open in a live editor instance in the current session; if no live instance exists, the call fails with `document-not-open` rather than editing a detached copy (FR-020).
- `fromVersion` equals the document's current version; otherwise the call fails with a conflict.
- All target paths resolve inside the session's permitted roots.
- Operations are dispatched to the live instance through the renderer-side command bridge and take the same version-guard and Change Record path as user edits.

## Postconditions

- A successful edit produces one Change Record with `actor = agent` and an `inverse`, advances the version once, and is visible in the open editor without a manual reload.
- A refusal produces no partial write and no version change.
- Tool results are structured and reconstructable from the session log.

## Errors

| Result | Meaning |
|---|---|
| `conflict` | `fromVersion` is stale; caller must re-read and retry |
| `unsupported-feature` | Operation would drop content the engine cannot preserve |
| `out-of-scope` | Target is outside permitted roots |
| `document-missing` | The document is closed or its file is missing |
| `document-not-open` | No live editor instance exists for the document |
| `bound-exceeded` | The operation exceeds a configured cap |

## Guarantees

- The agent cannot bypass the filesystem seam, version guard, or sandbox.
- Repeated identical edits at the same `fromVersion` never apply twice.
