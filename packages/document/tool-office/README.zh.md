---
description: "通过实时编辑器命令桥读取和编辑 Office 文件的模型面向文档工具。"
kind: "package-reference"
---

# @forwork/tool-office

[English](README.md) | 中文

## 概述

面向 agent Office 文档工具的 Host 插件包。它注册 `document_read`、`document_edit_cells`、`document_edit_paragraph` 与 `document_convert`。读取与编辑通过关联的 `document/command` 桥进入实时编辑器；Host 为每次接受的编辑执行唯一一次版本守卫保存，并记录带逆操作的 `document/edit`。

## 目录

- [使用这个包](#use-this-package)
- [理解实现](#understand-the-implementation)
- [延伸阅读](#further-exploration)
- [模型体验](#model-experience)
- [已知限制与延后工作](#known-limitations-and-deferred-work)
- [开发备注](#dev-note)

-----

<a id="use-this-package"></a>
## 使用这个包

在同时提供 `ctx.tools` 与 `ctx.officeDocument` 的 Host 组合中挂载该插件：

```ts
import * as ToolOffice from '@forwork/tool-office'

declare const ctx: import('@forwork/cordis').Context

await ctx.plugin(ToolOffice, { summaryCells: 200, maxEditSteps: 2_000 })
```

`ctx.officeDocument` 必须与其提供方一起挂载，因为这些工具通过提供方的 `officeCommandBroker` 分发命令。只挂载 Remote 而不挂载 broker 的组合仍可读取文件，但每条命令都以 `conversion-failed` 拒绝，而不是编辑一份脱离的副本。

模型可见的工具是：

- `document_read({ documentId })` —— 读取实时单元的结构（工作表、单元格、公式、段落、表格），从不保存。
- `document_edit_cells({ documentId, fromVersion, operations })` —— 替换单元格值或公式，然后保存一次。
- `document_edit_paragraph({ documentId, fromVersion, operations })` —— 替换段落，然后保存一次。
- `document_convert({ path })` —— 把旧版 `.doc` 或 `.xls` 文件转换为工作区内的可编辑副本。

文档 id 是 Session 作用域的：`<sessionId>:<path>`，即 Office 编辑器已用于草稿的身份。命令发送前，路径按 Session 的工作区根解析并用 `assertWithinRoots` 限制范围。

`fromVersion` 是文档的数字内容版本：文档打开时为 `0`，每次接受的保存加一。`document_read` 返回实时值；调用方版本不一致的编辑会以带 `currentVersion` 的 `conflict` 拒绝。

每次拒绝都是带稳定 `code` 的结构化结果：`conflict`、`unsupported-feature`、`out-of-scope`、`document-missing`、`document-not-open`、`pending-edits`、`bound-exceeded`、`timeout`、`invalid-request` 或 `conversion-failed`。

-----

<a id="understand-the-implementation"></a>
## 理解实现

<details>
<summary>实现内部细节——点击展开</summary>

`src/index.ts` 注册四个工具定义并从 `ctx.officeDocument` 构建 Host 运行时；部署上限来自经验证的插件 `Config`。`src/read.ts` 分发 `{ kind: 'read' }` 命令，并通过 `src/summary.ts` 总结返回的快照。`src/edit.ts` 拥有共享编辑路径：按身份中的工作区相对路径读取文件，使编辑器 tab 与打开事实保持同一个文档 id；从编辑前快照计算逆操作、分发 `{ kind: 'edit', steps }`，然后用命令前观察到的 filesystem 版本与智能体当前的打开 turn 保存返回的快照，并追加一条 `actor: 'agent'` 的 `document/edit` 事实。`src/edit-cells.ts` 与 `src/edit-paragraph.ts` 验证各自的操作列表并构建类型化步骤；`src/convert.ts` 位于 Host 侧。`src/errors.ts` 把 Remote 与 Exchange 失败映射为模型面向的代码。

命令事件不携带快照。被拒绝、未应答或超时的命令不写入任何内容，也不追加编辑事实，因此只有被接受的保存才能改变持久文档内容。

</details>

-----

<a id="further-exploration"></a>
## 延伸阅读

- [Office Exchange](../office-exchange/README.zh.md) —— 这些工具消费的转换能力、Remote 表面与命令 broker。
- [Document group](../README.zh.md) —— 该组中的 Office 转换包。
- [工具目录](../../../docs/tool-catalog.zh.md) —— 生成的模型面向工具参考。

-----

<a id="model-experience"></a>
## 模型体验

### 文档工具

#### 模型看到什么

四个工具 schema。`document_read` 返回有界的结构摘要，包含工作表与单元格计数、按行优先的单元格预览、公式、段落与表格尺寸。`document_edit_cells` 与 `document_edit_paragraph` 返回新的 `contentVersion` 与编辑器应用的操作。`document_convert` 返回转换后的路径与可编辑格式。拒绝返回 `status: "error"`、稳定的 `code` 以及可操作的说明。

#### Token 影响

读取摘要是此包产生的唯一可变大小的模型输入。`summaryCells`（默认 200）限制单元格预览，而计数描述完整单元，因此大型工作簿永远不会淹没请求；被截断的摘要会标记为 `truncated`。

#### KV Cache 影响

无；此包既不组装也不发送 provider 请求。

## 已知限制与延后工作

<a id="known-limitations-and-deferred-work"></a>

- 稳定文档 id 目前等于 `<sessionId>:<path>`。`document` projection（US3）拥有稳定 id 映射与模型面向的打开文档集合；在它落地之前，模型必须从编辑器或先前的读取中获得该 id。
- 这些工具只暴露单元格与段落编辑。`document/edit` 步骤联合还包含行、列、公式与样式步骤：`document_edit_cells` 会发出公式步骤，但还没有工具表面发出行／列插入或删除或样式操作。
- `document_convert` 返回工作区下的转换副本；没有任何机制自动清理 `.dsh/office-tmp/` 转换目录。
- 编辑需要实时编辑器实例。无头 Session 以 `document-not-open` 拒绝，而不是编辑并保存脱离的文件。

<a id="dev-note"></a>
### 开发备注

<details>
<summary>维护者工作上下文——点击展开</summary>

测试挂载真实的 Host 组合（`LocalFileSystem`、`LocalOfficeExchange`、`ToolRuntime`、`AgentLoop`），只 spy `officeDocument.dispatchCommand`，因此保存路径、事件追加与文件限制都真实运行。`tests/harness.ts` 拥有该组合。

</details>
