---
description: "Host 文档 Session 原语：稳定身份绑定、版本守卫的变更记录、持久文档事实，以及打开文档的 projection 与 prompt 区段。"
kind: "package-reference"
---

# @forwork/session-document

[English](README.md) | 中文

## 概述

`dsh-session-document` 在路径移动时保持文档身份稳定，让每次接受的编辑可归属且可逆，并把持久的 `document/*` 事实折叠为打开文档集合。身份绑定把 Session 本地的 `DocumentId` 与定位符及其可用性配对；变更记录把编辑与 actor、版本迁移和逆操作配对；`document` projection 与 prompt 区段读回这些事实。Restore 把 projection 展开为 Client 重新打开的有序集合，恢复规则把已存储草稿与它捕获自的文件匹配。基础 profile 的 `session-document` Host 插件不访问文件系统，也不拥有存储：Session 日志是唯一的持久来源。

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

### 何时使用

当打开文档的路径可能变化、编辑必须在文档层面撤销，或 agent 需要当前打开集合时使用此包：Office 编辑器 Remote、重启后的恢复以及文档工具。基础 profile 挂载 `session-document` 插件，它注册 `document` projection 与打开文档的 prompt 区段；只要手边有活跃 `Session`，就可以使用这些事实助手。调用方拥有文件系统观察；Session 日志拥有持久性。

### 绑定身份

```text
const binding = createDocumentIdentity(documentId, 'C:/work/report.docx')
const afterMove = rebindDocument(binding, 'C:/work/renamed.docx', at, true)
const afterLoss = markDocumentMissing(afterMove, at)
const found = resolveDocumentBinding(bindings, documentId)
```

`documentId` 在每次重绑与每次缺失观察后都保留。`rebindDocument` 接受显式的 `exists` 标志：调用方报告新路径是否解析成功，因此移动到不可用路径会保持 `missing: true`，而不是清除它。缺失的绑定连同历史留在注册表中；它从不关闭文档。此处路径是不透明字符串，因此 filesystem seam 产生的任何定位符都可原样使用。

### 应用并撤销变更

```text
const { record, version } = applyChange({
  documentId,
  actor: 'agent',
  currentVersion: liveVersion,
  fromVersion: claimedVersion,
  operation,
  inverse,
  at,
})
const undo = undoChange(record)
```

`applyChange` 是版本守卫。它不拥有存储，因此调用方把实时文档版本作为 `currentVersion` 传入；只有当 `fromVersion` 等于它时变更才被接受。被接受的变更产生新记录以及 `toVersion = fromVersion + 1`。记录携带结构化步骤及其逆操作，从不携带单元快照，因此变更可从 Session 日志重建。`undoChange` 把一条记录投影为审计与恢复用途的逆操作步骤和要恢复的版本。运行时撤销与重做使用编辑器运行时自身的原生历史，绝不是第二套逆操作栈。

`changeId` 默认使用确定性的 `documentId:toVersion`；显式传入时保持原值。每次拒绝都是带稳定 code 的 `OfficeDocumentError`：

- `stale-version` —— `fromVersion` 不是实时版本；调用方必须重新加载或合并。
- `invalid-operation` —— 操作或逆操作为空，因此变更无法应用或反转。
- `invalid-transition` —— 版本无法描述真实文档，或无法安全推进。

### 记录文档事实

```text
recordDocumentOpen(session, { documentId, path, format, contentVersion })
recordDocumentSave(session, { documentId, actor, contentVersion, turn? })
recordDocumentRebind(session, { documentId, fromPath, toPath })
recordDocumentMissing(session, { documentId, path })
recordDocumentClose(session, { documentId, reason })
```

每个助手追加一条普通 Session 事件并返回它。事实是仅追加的：之后的观察会追加新事实，而不是替换较早的事实。改动留在编辑器内的保存是 `document/save`；它记录作者（`user` 或 `agent`）、已发布版本，以及存在时的致因 turn，不携带操作步骤，也从不伪造逆操作。智能体编辑的步骤与逆操作仍留在 `document/edit`。

### 读取打开集合

`document` projection 持有 `{ order, documents }`：显示顺序，以及每个打开文档的一条条目，包含其身份、当前路径、格式、内容版本、`open` 或 `missing` 状态与不支持特性。`document/open` 插入或更新条目，`document/save` 与 `document/edit` 推进其版本，`document/rebind` 替换路径，`document/missing` 标记为不可用，`document/close` 删除它。未知事件保持同一个 state 引用，因此不运行任何下游工作。线缆视图是同一个值，并带有供 Client 恢复的 schema。只知道当前定位符的调用方（例如重启后的文件读取）用 `openDocumentAtPath` 解析已有条目，而不是从路径派生第二个身份。

该插件的 prompt 区段为正在组装的 Agent 的 Session 渲染打开集合；没有 Agent 或没有打开文档时不渲染内容。每行以条目的精确 `documentId` 开头，文档工具调用正是原样传回该身份。projection 与区段都随挂载 fiber 释放。

### 恢复打开集合

```text
const entries = restoreSet(projection)
const decision = recoveryDecision(draft, readResult.version)
```

`restoreSet` 按 projection 自身的顺序为每个打开文档返回一条条目，携带身份、当前路径、格式与状态。`missing` 文档仍留在集合中：Session 连同历史保持其打开，因此调用方报告其路径，而不是打开一个已失效的 tab。已关闭的文档本就不在 projection 中。

`recoveryDecision` 是唯一的草稿规则：没有草稿时为 `none`，草稿捕获的版本等于当前文件版本时为 `offer`，其余情况为 `stale`。过期草稿只在用户显式选择后提供；此处没有任何逻辑会应用它。

该包还暴露 `./types`（projection 的纯类型及其 `SessionProjectionMap` 键）与 `./client`（对 `./types` 的仅类型重导出），使 Client 包可以绑定 `useProjection('document')` 而无需导入 Host 值模块。

-----

<a id="understand-the-implementation"></a>
## 理解实现

<details>
<summary>实现内部细节——点击展开</summary>

这两个模块都是针对不可变值的纯函数集合；每个函数返回新对象，从不修改输入。

### 源映射

| 文件 | 职责 |
|---|---|
| [`src/document-identity.ts`](src/document-identity.ts) | `DocumentIdentityBinding`，及其 create、rebind、missing 与 resolve 迁移，以及共享的 `documentIdFor` 身份 |
| [`src/document-changes.ts`](src/document-changes.ts) | `ChangeRecord`、`ChangeActor`、`OfficeDocumentError`、版本守卫与撤销原语 |
| [`src/document-facts.ts`](src/document-facts.ts) | 五个 `document/*` 事实的仅追加写入器 |
| [`src/document-projection.ts`](src/document-projection.ts) | `document` projection fold、state 与线缆 schema |
| [`src/types.ts`](src/types.ts) | 纯条目／state／视图类型，以及 `document` projection 键声明 |
| [`src/client.ts`](src/client.ts) | 面向 Client 程序的 `./types` 仅类型重导出 |
| [`src/document-restore.ts`](src/document-restore.ts) | `restoreSet`：Client 重新打开的有序、带缺失标记的集合 |
| [`src/document-recovery.ts`](src/document-recovery.ts) | `recoveryDecision`：提供、过期或忽略一份已存储草稿 |
| [`src/plugin.ts`](src/plugin.ts) | `session-document` 插件：projection 注册与打开文档的 prompt 区段 |
| [`src/index.ts`](src/index.ts) | 包入口，重新导出各模块、prompt 渲染器与插件的 `name`／`inject`／`apply` |

</details>

-----

<a id="further-exploration"></a>
## 延伸阅读

- [Session 包索引](../README.zh.md) —— 该库支撑的持久化、projection 与标题包。
- [Session projection 子系统](../../../docs/subsystems/session-projection.zh.md) —— `document` projection 如何折叠发出的 `document/*` 事实。
- [文档事件契约](../../../specs/001-office-agent-framework/contracts/document-events.contract.md) —— 这些值所供给的事件词汇。

-----

<a id="model-experience"></a>
## 模型体验

### 文档迁移

#### 模型看到什么

当至少打开一个文档时，prompt 区段渲染一份简短的英文列表：由位置顺序体现的打开文档数量、每个文档的 Session 作用域 `<sessionId>:<path>` id、其文件名、其格式以及 `missing` 标记。该 id 正是文档工具接受的身份，因此模型无需额外读取即可寻址打开文档。没有 Agent 或没有打开文档时，区段不渲染任何内容；文档工具拥有自己的模型可见 schema 与结果。

##### 打开文档 prompt 区段

```markdown
Open documents in this session:
1. document-plugin:work/budget.xlsx (budget.xlsx, xlsx)
2. document-plugin:work/report.docx (report.docx, docx, missing)
```

#### Token 影响

每个打开文档一行，再加一行标题，且仅在文档打开期间存在。每行的 id 重复完整路径，因此行长随路径长度增长；超过 40 个字符的文件名会截断为前 37 个字符加 `...`，而 id 保持精确。空打开集合增加零 token。

#### KV Cache 影响

只有当 `document/*` 事实改变打开集合，或改变某条目的路径、格式或状态时，列表才变化；每个保持打开的文档都保持稳定位置。

## 已知限制与延后工作

<a id="known-limitations-and-deferred-work"></a>

- **文件系统观察由调用方提供。** `rebindDocument` 信任 `exists` 标志，`markDocumentMissing` 信任其调用方；探测路径的注册表属于消费方，不属于此包。Office 编辑器会从其资源失败中报告文件缺失，并在变更跟踪器于新路径观察到同一内容版本时请求 Remote 重绑定；没有任何打开的 tab 或 feed 观察到的移动会让文档留在旧路径，直到它重新打开。
- **恢复与草稿恢复止于决策。** `restoreSet` 展开 projection，`recoveryDecision` 对草稿分类；打开 tab、每个 Session 只执行一次、报告缺失路径以及保存后清除草稿，都属于拥有这些 tab 的 Client。
- **projection 不携带不支持特性事实。** `document/open` 目前没有承载它们的字段，因此条目初始为空；在事实携带该列表之前，Exchange 读取结果仍是保存拒绝的权威。
- **版本守卫不串行化写入方。** 调用方必须从唯一的权威文档所有者传入实时 `currentVersion`；该库检测不匹配，但无法阻止两个并发调用读到同一版本。
- **记录有效性只做结构检查。** 该库要求非空的操作与逆操作，但无法证明逆操作恢复先前内容；同时产生两者的编辑提供方负责精确可逆性。

<a id="dev-note"></a>
### 开发备注

<details>
<summary>维护者工作上下文——点击展开</summary>

无。

</details>
