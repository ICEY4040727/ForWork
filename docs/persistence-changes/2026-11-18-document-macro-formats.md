---
description: "Records a persistence type transition and its compatibility acknowledgement."
kind: persistence-change
---

# 2026-11-18-document-macro-formats

English | [中文](2026-11-18-document-macro-formats.zh.md)

## Summary

Widens the DocumentFormat union persisted by document/open with the macro-enabled docm and xlsm literals.

## Table of Contents

- [Declaration](#declaration)
- [Compatibility](#compatibility)
- [Verification](#verification)
- [Dev Note](#dev-note)

<a id="declaration"></a>
## Declaration

```yaml persistence-change
schemaVersion: 1
id: 2026-11-18-document-macro-formats
baseline: false
changes:
  - root: "event:document/open"
    previous: "2026-11-14-document-events"
    after: "350212adc337e105fc7557dcf5b9115a5adb88c8e3609785d2712f77ab92a343"
    decision: same-version
```

<a id="compatibility"></a>
## Compatibility

Every existing open fact keeps its previous literal, and new readers fold both established and new literals without changing their meaning. Older readers admit the known document/open event and preserve a docm or xlsm payload, but their document projection schema rejects the new literals. documentProjectionDefinition.stateVersion advances from 1 to 2, so a projection checkpoint row written under the other schema is discarded and never migrated; rebuilding the projection or serving its view on the older schema fails on the new literal instead of reinterpreting it, before any mutation of document state. No required field, event type, envelope shape, or SESSION_FORMAT_VERSION changes.

<a id="verification"></a>
## Verification

pnpm exec vitest run packages/session/session-document/tests packages/document/office-exchange/tests/provider-service.spec.ts: 8 files, 66 tests passed. The session-document projection tests cover both new literals and stateVersion 2; pnpm exec tsc -b packages/core/session packages/session/session-document passed.

<a id="dev-note"></a>
## Dev Note

None.
