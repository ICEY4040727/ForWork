---
description: "Records a persistence type transition and its compatibility acknowledgement."
kind: persistence-change
---

# 2026-11-15-document-command-and-steps

English | [中文](2026-11-15-document-command-and-steps.zh.md)

## Summary

Adds the document/command Session event and extends the document/edit step union with row, column, formula, and style alternatives.

## Table of Contents

- [Declaration](#declaration)
- [Compatibility](#compatibility)
- [Verification](#verification)
- [Dev Note](#dev-note)

<a id="declaration"></a>
## Declaration

```yaml persistence-change
schemaVersion: 1
id: 2026-11-15-document-command-and-steps
baseline: false
changes:
  - root: "event:document/command"
    previous: null
    after: "2f9b9def9e02ddec07c7faea633a7ea8f8b572da2f6be849eeb26c47dd7b6995"
    decision: same-version
  - root: "event:document/edit"
    previous: "2026-11-14-document-events"
    after: "c1a6d74613aeff3b01501987b97ffb725ab1acf0453aef289c6280b4228563f2"
    decision: same-version
```

<a id="compatibility"></a>
## Compatibility

Both reader directions remain safe at format version 4. Older readers do not know document/command and refuse the Session at admission against KNOWN_SESSION_EVENT_TYPES instead of skipping a required correlated request; the request carries no snapshot and its accepted outcome remains the separate document/edit fact. Older readers that receive a row, column, formula, or style step in document/edit likewise refuse the Session at admission because the step union is required-on-read; an unknown step is never reinterpreted as a cell or paragraph. The four existing step shapes and their field meanings are unchanged, and new readers retain every older operation.

<a id="verification"></a>
## Verification

tsc -b packages/core/session packages/session/session-document packages/document/office-exchange packages/document/tool-office passed; vitest run core/session + session-document: 610 tests passed; office-exchange: 134 tests passed; tool-office: 80 tests passed with 100% per-file statement/branch/function/line coverage for packages/document/tool-office/src.

<a id="dev-note"></a>
## Dev Note

None.
