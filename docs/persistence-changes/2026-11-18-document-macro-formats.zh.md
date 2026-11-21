---
description: "记录持久化类型更改及其兼容性确认。"
kind: persistence-change
---

# 2026-11-18-document-macro-formats

[English](2026-11-18-document-macro-formats.md) | 中文

## 概述

在 document/open 持久化的 DocumentFormat 联合类型中增加宏格式 docm 与 xlsm。

## 目录

- [声明](#declaration)
- [兼容性](#compatibility)
- [验证](#verification)
- [开发备注](#dev-note)

<a id="declaration"></a>
## 声明

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
## 兼容性

既有打开记录保留原有字面量，新读取方按相同语义折叠既有与新增字面量。旧读取方接受已知的 document/open 事件并保留 docm 或 xlsm 载荷，但其 document 投影 schema 不接受新字面量。documentProjectionDefinition.stateVersion 由 1 升至 2，因此按另一 schema 写入的投影检查点行会被丢弃且永不迁移；在旧 schema 上重建投影或提供其视图会因新字面量而失败，而不是重新解释该值，且发生在文档状态发生任何变更之前。必需字段、事件类型、信封结构与 SESSION_FORMAT_VERSION 均未改变。

<a id="verification"></a>
## 验证

pnpm exec vitest run packages/session/session-document/tests packages/document/office-exchange/tests/provider-service.spec.ts：8 个文件、66 个测试通过。session-document 投影测试覆盖两个新字面量与 stateVersion 2；pnpm exec tsc -b packages/core/session packages/session/session-document 通过。

<a id="dev-note"></a>
## 开发备注

无。
