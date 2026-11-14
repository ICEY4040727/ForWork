---
description: "记录持久化类型更改及其兼容性确认。"
kind: persistence-change
---

# 2026-11-14-document-events

[English](2026-11-14-document-events.md) | 中文

## 概述

新增五个持久化的 document/* Session 事件：open、edit、rebind、missing 与 close。

## 目录

- [声明](#declaration)
- [兼容性](#compatibility)
- [验证](#verification)
- [开发备注](#dev-note)

<a id="declaration"></a>
## 声明

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
## 兼容性

新增普通事件类型在两个方向上都是同版本变更。旧版读取器不认识这些类型，会在会话准入阶段依据 KNOWN_SESSION_EVENT_TYPES 拒绝该 Session，而不是静默丢弃必需事实。新版读取器逐字保留全部文档事实。这些事件不携带 ignorable 标记，因为 document 投影缺少它们就无法重建已打开文档集合。SESSION_FORMAT_VERSION 保持为 4。

<a id="verification"></a>
## 验证

corepack pnpm exec tsc -b packages/core/session：通过。corepack pnpm exec vitest run packages/core/session/tests/document-events.spec.ts：6 个测试通过。corepack pnpm run gen-persistence-catalog --check：通过。

<a id="dev-note"></a>
## 开发备注

无。
