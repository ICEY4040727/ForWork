---
description: "记录持久化类型更改及其兼容性确认。"
kind: persistence-change
---

# 2026-11-18-document-save-disposition

[English](2026-11-18-document-save-disposition.md) | 中文

## 概述

为 document/save 会话事件增加可选的 disposition 保存摘要。

## 目录

- [声明](#declaration)
- [兼容性](#compatibility)
- [验证](#verification)
- [开发备注](#dev-note)

<a id="declaration"></a>
## 声明

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
## 兼容性

既有记录仍然有效。该属性可选，旧读取方保留文档 id、actor 与 contentVersion，并安全忽略整个子树，而不是丢弃保存记录；新读取方将缺失视为记录 disposition 之前的保存。任何调用方都不得自行合成 disposition，document 投影也忽略该属性，因此不伴随投影条目、zod schema、sameEntry 或 stateVersion 变更。SESSION_FORMAT_VERSION 保持为 4。

<a id="verification"></a>
## 验证

pnpm exec vitest run packages/session/session-document/tests packages/document/office-exchange/tests/provider-service.spec.ts packages/document/office-exchange/tests/config.spec.ts：9 个文件、72 个测试通过。

<a id="dev-note"></a>
## 开发备注

无。
