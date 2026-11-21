---
description: "Records a persistence type transition and its compatibility acknowledgement."
kind: persistence-change
---

# 2026-11-18-document-save-disposition

English | [中文](2026-11-18-document-save-disposition.zh.md)

## Summary

Adds an optional disposition preservation summary to the document/save Session event.

## Table of Contents

- [Declaration](#declaration)
- [Compatibility](#compatibility)
- [Verification](#verification)
- [Dev Note](#dev-note)

<a id="declaration"></a>
## Declaration

```yaml persistence-change
schemaVersion: 1
id: 2026-11-18-document-save-disposition
baseline: false
changes:
  - root: "event:document/save"
    previous: "2026-11-15-document-save"
    after: "9f2ac939d062f9304b0ebfba479f90a3eaa933821f0e2b615078928562122284"
    decision: same-version
```

<a id="compatibility"></a>
## Compatibility

Existing records remain valid. The property is optional, so older readers keep the document id, actor, and contentVersion and safely ignore the whole subtree instead of dropping the save fact; new readers treat absence as a save recorded before dispositions existed. No caller may synthesize a disposition, and the document projection ignores the property, so no projection entry, zod schema, sameEntry, or stateVersion change accompanies it. SESSION_FORMAT_VERSION stays at 4.

<a id="verification"></a>
## Verification

pnpm exec vitest run packages/session/session-document/tests packages/document/office-exchange/tests/provider-service.spec.ts packages/document/office-exchange/tests/config.spec.ts: 9 files, 72 tests passed.

<a id="dev-note"></a>
## Dev Note

None.
