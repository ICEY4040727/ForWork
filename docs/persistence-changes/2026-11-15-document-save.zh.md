---
description: "记录持久化类型更改及其兼容性确认。"
kind: persistence-change
---

# 2026-11-15-document-save

[English](2026-11-15-document-save.md) | 中文

## 概述

新增 document/save 会话事件，用于记录内容改动仍留在编辑器内的用户保存。

## 目录

- [声明](#declaration)
- [兼容性](#compatibility)
- [验证](#verification)
- [开发备注](#dev-note)

<a id="declaration"></a>
## 声明

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
## 兼容性

在格式版本 4 下两个读取方向都保持安全。旧读取方不认识 document/save，会依据 KNOWN_SESSION_EVENT_TYPES 在会话准入时拒绝该会话，而不是丢弃必需的记录；也不会把保存误读为 document/edit。载荷携带文档 id、作者与已发布的 contentVersion，新读取方据此折叠版本，无需编造操作步骤或逆操作。既有文档事件的载荷与语义不变。

<a id="verification"></a>
## 验证

tsc -b packages/core/session packages/core/system-prompt packages/session/session-projection packages/session/session-document packages/document/office-exchange packages/document/tool-office packages/bundle/base 通过；vitest run session-document：30 个测试通过；office-exchange remote/drafts/commands：46 个测试通过；office editor 客户端：165 个测试通过。

<a id="dev-note"></a>
## 开发备注

无。
