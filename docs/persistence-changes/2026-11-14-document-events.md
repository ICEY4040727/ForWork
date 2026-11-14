---
description: "Records a persistence type transition and its compatibility acknowledgement."
kind: persistence-change
---

# 2026-11-14-document-events

English | [中文](2026-11-14-document-events.zh.md)

## Summary

Adds the five durable document/* Session events: open, edit, rebind, missing, and close.

## Table of Contents

- [Declaration](#declaration)
- [Compatibility](#compatibility)
- [Verification](#verification)
- [Dev Note](#dev-note)

<a id="declaration"></a>
## Declaration

```yaml persistence-change
schemaVersion: 1
id: 2026-11-14-document-events
baseline: false
changes:
  - root: "event:document/close"
    previous: null
    after: "ad2b80a7057a6a4d32e5bf7de0697a7c6c6989715990ea7946b2914e54d57322"
    decision: same-version
  - root: "event:document/edit"
    previous: null
    after: "1c7187ee16caf169de2c7d801139685865b15afbf386a739da5e4ded133a605f"
    decision: same-version
  - root: "event:document/missing"
    previous: null
    after: "db4455e69146554b94bb759e4eb958972b8e4127775513d1a013c1921dedcb34"
    decision: same-version
  - root: "event:document/open"
    previous: null
    after: "31b3e8f76f9c0408809813625d789bc832568ff3fb7bf65dc2f9f4d543828108"
    decision: same-version
  - root: "event:document/rebind"
    previous: null
    after: "99d72b97ac6e7e2fdef89bdfa688f4c00b7650762e3d9dff42c7274d0d1566c3"
    decision: same-version
```

<a id="compatibility"></a>
## Compatibility

Adding an ordinary event type is a same-version change in both directions. Older readers do not know these types and refuse the Session at admission against KNOWN_SESSION_EVENT_TYPES instead of silently dropping required facts. New readers preserve every document fact verbatim. The events carry no ignorable marker because the document projection cannot rebuild the open-document set without them. SESSION_FORMAT_VERSION stays at 4.

<a id="verification"></a>
## Verification

corepack pnpm exec tsc -b packages/core/session: passed. corepack pnpm exec vitest run packages/core/session/tests/document-events.spec.ts: 6 tests passed. corepack pnpm run gen-persistence-catalog --check: passed.

<a id="dev-note"></a>
## Dev Note

None.
