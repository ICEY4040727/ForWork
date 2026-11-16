---
description: "Office Exchange 能力 seam：OOXML 与单元快照的映射、不支持特性检测，以及由 Office 编辑器与 agent 文档工具共享的旧版转换。"
kind: "package-reference"
---

# @forwork/office-exchange

[English](README.md) | 中文

## 概述

定义 Office 文件进出可编辑单元模型的契约。该 seam 命名转换与丢失检测；提供方负责 ExcelJS、SheetJS、docx 与 LibreOffice 的具体机制。编辑器与 agent 工具消费 `ctx.officeExchange`，不直接导入转换库。

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

该包默认导出 `LocalOfficeExchange`，一个注册 `ctx.officeExchange` 的 Cordis 插件。挂载时给出部署边界，之后通过 seam 读写：

```ts
import LocalOfficeExchange from '@forwork/office-exchange'

declare const ctx: import('@forwork/cordis').Context
declare const path: string
declare const signal: AbortSignal

await ctx.plugin(LocalOfficeExchange, {
  maxFileBytes: 25 * 1024 * 1024,
  maxCells: 2_000_000,
  maxConversionMs: 60_000,
  commandTimeoutMs: 5_000,
})

const { format, snapshot, unsupported } = await ctx.officeExchange.read(path, signal)
```

`read(path)` 返回检测到的格式、可编辑的单元快照，以及所有无法经受编辑并保存的特性。`write(snapshot, format)` 把快照编码回文件字节但不写入。`detectUnsupported(path)` 报告可能丢失的特性，不把内容读入编辑器。`convertLegacy(path, signal?, outputDir?)` 把 `.doc/.xls/.ppt` 文件转换为可编辑的等价文件；省略 `outputDir` 时写入任意工作区之外的新临时目录。

支持的格式：

- `.xlsx` 与 `.docx` 可编辑：`read` 与 `write` 将它们映射为 Univer 工作簿与文档快照。
- `.doc`、`.xls` 与 `.ppt` 必须先经 `convertLegacy` 转换再 `read`；直接读取会报 `unsupported-format`。
- `.pptx` 仅供预览：`detectUnsupported` 检查它，而 `read` 与 `write` 报 `unsupported-format`。

每个失败都是带稳定 `code` 的 `OfficeExchangeError`；边界在挂载时解析一次，无效值会显式报错。

`formatOf(path)` 按扩展名分类路径；`isEditableFormat` 区分可直接编辑的 `.docx/.xlsx` 与需要转换的格式；`isLegacyFormat` 指出旧版格式集合。

### Remote 表面

该插件还挂载 `ctx.officeDocument`，一个在同一 seam 之上的 Session 作用域 Remote 表面。它生成的 Client 命名空间是 `remote.officeDocument`；把该包挂载进 Host 组合，并把 `@forwork/office-exchange/remote` 加入 Client 组装。

该 Remote 从请求头派生的作用域解析实时 Session，并在其中记录 `document/*` 事实；`session-document` 插件把它们折叠为打开集合。

- `read(workspaceFileScope, path)` 从 Session 的工作区根解析路径，用 `assertWithinRoots` 限制范围，在提供方的 `maxFileBytes` 边界内通过 filesystem seam 读取，并返回 `{ format, snapshot, unsupported, version }`。本进程中某文档的第一次读取会追加一次 `document/open`；之后的读取不更新任何事实。
- `save(workspaceFileScope, path, snapshot, format, expectedVersion)` 通过提供方序列化快照，并以 filesystem `replaceIfVersion` 守卫发布。自读取以来发生变化的文件会以 `office-document/conflict` 拒绝；写入从不退化为无条件保存。被接受的保存追加带 `actor: 'user'` 与下一内容版本的 `document/save`；内存中的用户改动留在编辑器内，因此不记录操作步骤或逆操作。
- `closeDocument(workspaceFileScope, documentId, reason)` 追加 `document/close`，并丢弃进程本地的打开条目。Host 重启后关闭的 tab 仍会留下持久事实。
- `rebindDocument(workspaceFileScope, documentId, toPath)` 追加 `document/rebind` 并保留身份；旧路径与新路径都必须解析到工作区根内。
- `markDocumentMissing(workspaceFileScope, documentId)` 追加一次 `document/missing` 并保持文档打开；之后的保存、编辑或重绑会把它重新标记为可用。
- `detectUnsupported(workspaceFileScope, path)` 报告易丢失的特性，不返回内容。
- `convertLegacy(workspaceFileScope, path)` 返回 `<workspaceRoot>/.dsh/office-tmp/convert-<hash>-<random>/` 内的转换后文件路径，因此同一个 Remote 可以在不变的根检查下 `read` 它。每次调用获得自己的目录。
- `writeDraft(workspaceFileScope, documentId, documentPath, contentVersion, snapshot)` 持久化一份恢复草稿并返回 `{ savedAt }`。草稿位于 `<workspaceRoot>/.dsh/office-tmp/drafts/<sha256(documentId)>.json`；摘要确保调用方选择的 id 无法命名草稿目录之外的路径。它替换已有草稿，且从不写入文档路径本身。
- `readDraft(workspaceFileScope, documentId)` 返回 `{ documentId, path, contentVersion, snapshot, savedAt }`，文档没有草稿时返回 `null`。格式错误的草稿以 `office-document/conversion-failed` 拒绝。
- `clearDraft(workspaceFileScope, documentId)` 删除草稿；对没有草稿的文档执行清除是无操作。
- `completeCommand(workspaceFileScope, requestId, result)` 用实时编辑器的答案结算一个待处理的关联命令。该提供方还挂载 `ctx.officeCommandBroker`，这个服务拥有待处理请求、其超时（`commandTimeoutMs`）以及每个文档最后观察到的数字内容版本。未知、迟到或属于其他 Session 的结果会被忽略并记录日志，绝不应用。

`completeCommand` 应答由 Host 工具层追加的 `document/command` Session 事件；此包本身不追加该事件。结果联合为 `read`、`applied`、`pending-edits`、`conflict`、`unsupported-feature` 或 `document-not-open`；读取或已应用的答案携带实时快照与数字内容版本，冲突携带 `currentVersion`。Host 只在 `applied` 答案后写文件。

草稿与版本绑定：记录携带快照捕获自的 `contentVersion`，因此恢复路径可以在提供恢复前与文件的当前版本比较。草稿目录与解析出的草稿文件都通过 `assertWithinRoots`，指向工作区根之外的链接会被拒绝，而不是穿透写入。

`OfficeDocumentSnapshot` 是带快照 kind 标记的 JSON 投影：Univer 的原生接口允许无约束自定义字段，JSON 线缆无法承载。Host 在序列化前验证该投影，并把每个失败映射到 `office-document/*` Remote 代码（`unsupported-format`、`unsupported-feature`、`out-of-scope`、`bound-exceeded`、`conversion-failed`、`conflict`）。

### 失败策略

失败是带稳定 `code` 的 `OfficeExchangeError` 值。无法打开文件的读取报告带路径的 `conversion-failed`，而不是泄漏 Node 错误。读取时非空的 `unsupported` 列表意味着写入路径必须拒绝，而不是保存有损文件；产品会指出违规特性并提供替代方案。

-----

<a id="understand-the-implementation"></a>
## 理解实现

<details>
<summary>实现内部细节——点击展开</summary>

`src/service.ts` 声明抽象的 `OfficeExchange` 服务、context 增强，以及由包入口重新导出的 seam 词汇。`src/types.ts` 保存 JSON 安全的线缆与 seam 词汇；`src/snapshots.ts` 保存 Univer 快照类型，使 Remote Client 永不加载编辑器核心。`src/remote.ts` 实现 `officeDocument` Remote 表面：Session 作用域解析、工作区限制、版本传递、JSON 投影、工作区作用域的转换目录、恢复草稿存储、命令分发与结算，以及其余的保存保证。`src/commands.ts` 拥有 `officeCommandBroker` 服务及其待处理请求注册表。该 Remote 还要求 `sessions` 服务：它只向活跃 Session 追加文档事实。`src/provider/index.ts` 在纯模块 `provider/xlsx.ts`、`provider/docx.ts`、`provider/detect.ts` 与 `provider/legacy.ts` 之上实现 `LocalOfficeExchange`：它拥有文件读取、格式分发、类型化读取失败、调用方指定的旧版转换目录，以及经验证的 `Config` schema。Office 编辑器与 agent 工具仍是消费方；此包不注册任何工具，Remote 唯一的 Session 写入就是上述 `document/*` 事实。

</details>

-----

<a id="further-exploration"></a>
## 延伸阅读

- [Document group](../README.zh.md) —— 该组中的 Office 转换包。
- [Office 转 PDF](../office-to-pdf/README.zh.md) —— 用于旧版格式的 LibreOffice 转换提供方。

-----

<a id="model-experience"></a>
## 模型体验

### Exchange seam

#### 模型看到什么

没有直接内容；agent 文档工具消费 `ctx.officeExchange`，并拥有自己的模型可见 schema 与结果。

#### Token 影响

零直接 token；该 seam 向消费方返回快照与特性列表，不添加 prompt 或结果文本。

#### KV Cache 影响

无；该 seam 既不组装也不发送 provider 请求。

## 已知限制与延后工作

<a id="known-limitations-and-deferred-work"></a>

- `LocalOfficeExchange` 只读写本地路径；工作区限制属于提供路径的工具层与 Session 层。
- 旧版转换需要 LibreOffice kit；引擎不可用时转换失败，不会回退。
- PowerPoint 编辑在加入幻灯片单元及其转换之前为只读。
- `convertLegacy` 把转换后的可编辑文件保留在 `<workspaceRoot>/.dsh/office-tmp/` 下，`officeDocument` Remote 可以在同一根检查下读取。失败的转换会删除它刚创建的目录；成功的转换会保留到用户或 Session 删除为止，目前没有任何机制自动清理。
- 恢复草稿会累积在 `<workspaceRoot>/.dsh/office-tmp/drafts/` 下，直到 `clearDraft` 删除；目前还没有任何机制为不再打开的文档清理草稿。
- Client 通过 `markDocumentMissing` 报告文件缺失，但目前没有 Client 触发器检测移动：`rebindDocument` 暴露在 Remote 表面上却没有自动调用方。打开集合是进程本地的，用于去重，因此 Host 重启后下一次读取会再次追加 `document/open`；projection 保留文档的位置与版本。
- 该 Remote 只为活跃 Session 记录事实。为未加载 Session 执行的读取以 `gateway/bad-request` 拒绝，而不是在没有持久化保证的情况下读取。

<a id="dev-note"></a>
### 开发备注

<details>
<summary>维护者工作上下文——点击展开</summary>

许可检查（2026-11-14）：`@univerjs/core` 是 Apache-2.0，`exceljs` 是 MIT。`@univerjs/core` 保持为 peer 依赖，因为此包只导入其类型并在声明中输出；`exceljs` 是 `.xlsx` 提供方的运行时依赖。

</details>
