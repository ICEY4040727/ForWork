---
description: "Records a persistence type transition and its compatibility acknowledgement."
kind: persistence-change
---

# 2026-11-15-document-save

English | [中文](2026-11-15-document-save.zh.md)

## Summary

Adds the document/save Session event for a user save whose in-memory changes stay inside the editor.

## Table of Contents

- [Declaration](#declaration)
- [Compatibility](#compatibility)
- [Verification](#verification)
- [Dev Note](#dev-note)

<a id="declaration"></a>
## Declaration

```yaml persistence-change
schemaVersion: 1
id: 2026-11-15-document-save
baseline: false
changes:
  - root: "event:document/save"
    previous: null
    after: "123e77bb36ab73e728f2d4c52e5cdb86343d11b563294d892f7bc0a19e47fc7d"
    decision: same-version
```

<a id="compatibility"></a>
## Compatibility

Both reader directions remain safe at format version 4. Older readers do not know document/save and refuse the Session at admission against KNOWN_SESSION_EVENT_TYPES instead of dropping a required fact; they never reinterpret a save as document/edit. The payload carries the document id, the authoring actor, and the published content version, so a new reader folds the version without inventing operation steps or an inverse. Every existing document event keeps its payload and meaning.

<a id="verification"></a>
## Verification

tsc -b packages/core/session packages/core/system-prompt packages/session/session-projection packages/session/session-document packages/document/office-exchange packages/document/tool-office packages/bundle/base passed; vitest run session-document: 30 tests passed; office-exchange remote/drafts/commands: 46 tests passed; office editor client: 165 tests passed.

<a id="dev-note"></a>
## Dev Note

None.
