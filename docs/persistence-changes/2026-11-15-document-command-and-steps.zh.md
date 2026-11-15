---
description: "记录持久化类型更改及其兼容性确认。"
kind: persistence-change
---

# 2026-11-15-document-command-and-steps

[English](2026-11-15-document-command-and-steps.md) | 中文

## 概述

新增 document/command Session 事件，并为 document/edit 的编辑步骤联合类型增加行、列、公式和样式分支。

## 目录

- [声明](#declaration)
- [兼容性](#compatibility)
- [验证](#verification)
- [开发备注](#dev-note)

<a id="declaration"></a>
## 声明

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
## 兼容性

在格式版本 4 下，两个读取方向都仍然安全。旧版读取器不认识 document/command，会依据 KNOWN_SESSION_EVENT_TYPES 在会话准入时拒绝该会话，而不是跳过必需的关联请求；该请求不携带快照，其被接受的结果仍写入独立的 document/edit 事实。旧版读取器若在 document/edit 中遇到行、列、公式或样式步骤，同样会在会话准入时拒绝该会话，因为步骤联合类型在读取时是必需的；未知步骤绝不会被重新解释为单元格或段落。原有四种步骤结构及其字段含义不变，新版读取器保留所有旧操作。

<a id="verification"></a>
## 验证

tsc -b packages/core/session packages/session/session-document packages/document/office-exchange packages/document/tool-office 通过；vitest run core/session 与 session-document：610 个测试通过；office-exchange：134 个测试通过；tool-office：80 个测试通过，packages/document/tool-office/src 的每文件语句、分支、函数和行覆盖率均为 100%。

<a id="dev-note"></a>
## 开发备注

无。
