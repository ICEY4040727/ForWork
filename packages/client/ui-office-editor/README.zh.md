---
description: "右侧 Sidebar 中内嵌的 Office 文档编辑 tab，包含 Client 与 Host 插件面。"
kind: "package-reference"
---

# @forwork/client-ui-office-editor

[English](README.md) | 中文

## 概述

右侧 Sidebar 中内嵌 Office 文档编辑器的 Client 插件包。它为 Session 作用域的 `.docx` 与 `.xlsx` 地址注册 `office` tab 类型，通过 `officeDocument` Remote 读取和保存，并在 adapter 接缝后挂载 Univer Sheets 或 Docs 核心 preset。编辑会以空闲防抖方式写入恢复草稿，重新打开文档时只提供草稿供选择，不会静默应用。打开旧版 `.doc`、`.xls` 或 `.ppt` 源文件时会先转换，再在新 tab 中打开副本：可编辑副本在此打开并显示标注原始文件的提示，转换后的演示文稿仍由只读预览处理。

## 目录

- [使用这个包](#use-this-package)
- [理解实现](#understand-the-implementation)
- [进一步探索](#further-exploration)
- [模型体验](#model-experience)
- [已知限制与延后工作](#known-limitations-and-deferred-work)
- [开发备注](#dev-note)

-----

<a id="use-this-package"></a>
## 使用这个包

Web bundle 负责挂载这个包。在右侧 Sidebar 打开 `.docx` 或 `.xlsx` 文件后，地址由 `office` 类型接管，并打开可编辑的表格或段落表面。表格与文字处理文档在 adapter 后分别使用 Univer preset。保存通过 `officeDocument` Remote 按读取时观察到的版本写入；版本过期会显示为冲突。报告了不支持特性的文件不能保存，面板会列出将丢失的特性。

针对当前打开文档的 agent command 通过 Session 的 `document/command` 事件流进入该 tab。该 tab 返回实时快照与数字内容版本，通过已挂载的 adapter 应用编辑步骤；版本过期时回复冲突，存在未保存的用户编辑时拒绝，并忽略面向其他文档的命令。被接受的编辑之后的保存保持同一个已挂载单元，因此编辑器的撤销仍能还原智能体的改动。

旧版 `.doc`、`.xls` 或 `.ppt` 源文件通过同一个 Remote 转换，并在该 tab 的位置重新打开副本；原文件永远不会被写入。转换后的 tab 显示一条可关闭的提示，标注原始文件。正文还会列出该 Session 中打开的 Office 文档，点击即可聚焦。

外部文件变化会重新载入干净的 tab，并在丢弃未保存编辑前询问。文件移动会根据打开 tab 的文件 feed 重新绑定：当文档路径消失、而另一个观察在同一文件版本的新路径上报告时，该 tab 会请求 Session 记录 `document/rebind`。挂载中的 tab 保留文档身份，同时把保存目标移到新路径；重启后按投影保留的身份重新打开。关闭有未保存编辑的 tab 会被 Sidebar 的清理失败路径否决，随后面板给出保存并关闭、丢弃并关闭或取消。

当单元有未保存编辑时，停止编辑 2 秒后会写入一份恢复草稿。重新打开文档时只读取一次草稿并提供恢复：草稿来自当前文件版本时直接提供，来自更早版本时标记为过期，并且仍只在用户选择时应用。保存成功和丢弃都会删除已存储的草稿。

Session 中的第一个触发点会读取 `document` projection，并按 projection 顺序重新打开该 Session 仍持有的文档。触发点既可以是 Session 中挂载的 Office 正文，也可以是该 Session 成为主视图 Session，因此没有打开 Office tab 的 Session 也会恢复其集合。每个 Session 在每个窗口只执行一次；已经有 Office tab 的文档不会再次打开，projection 报告为缺失的文档会列入可关闭的提示，而不是打开死 tab。每个恢复的文档随后走上文普通的打开路径，因此其存储的草稿仍由同一套恢复界面提供。

-----

<a id="understand-the-implementation"></a>
## 理解实现

<details>
<summary>实现内部细节——点击展开</summary>

`src/index.ts` 是 Host 伴生面，为 Client loader 定义一个空的 apply 函数。`src/client/index.ts` 是浏览器插件：它把 `officeDefinition()` 注册进 `ctx.sidebarRightTabs`，注册字典、`sidebar.right.pane.tab` 下的正文、`sidebar.right.pane.tab.title` 下的 chip 标题，并通过 `ctx.sidebarRight.registerCloseHandler` 注册脏关闭否决。定义让正文在切换 tab 时保持挂载，因此未保存的编辑不会丢失。

`definition.ts` 声明 `office` 类型、它的 `.docx`/`.xlsx` pattern、需要转换的旧版 `.doc`/`.xls`/`.ppt` 源以及 Session 文件否决。`OfficePanel.tsx` 通过 `useResource` 读取文件元数据，对注入的 `officeDocument` Remote 调用 `useOfficeDocument`，并且只在读取成功后挂载 adapter。`useOfficeDocument.ts` 负责读取与保存状态、期望版本检查、失败分类、外部变化规则、恢复提供，以及通过 tab 自身的 `openResource` 重新打开副本的旧版转换。`command-bridge.ts` 把一个 tab 订阅到 Session 追加的 `document/command` 事实，在挂载期间通过 `officeDocument.registerCommandTarget` 注册实时目标并在卸载时释放，按实时文档身份、数字内容版本与脏状态对每个请求恰好结算一次，通过 adapter handle 应用编辑步骤，并通过 `officeDocument.completeCommand` 结算 Host 调用。`drafts.ts` 负责空闲防抖调度器：每个面板创建一个实例，编辑会替换待写入候选，`flush` 立即写入，`cancel` 等待进行中的写入，任何拒绝或失败都会进入 hook 的草稿失败状态。`changes.ts` 负责面板与关闭处理器共享的脏状态和关闭警告 snapshot store；`notices.ts` 按转换后 tab 的 content id 保存转换副本记录；`restore.ts` 保存每个 Session 的恢复记录，并提供按 projection 顺序打开缺失文档的 hook。`DocumentPanel.tsx` 从 `sidebarRight.openTabs` 列出该 Session 中打开的 `office` tab。`target.ts` 把地址映射为 Remote 路径，命名共享的文档身份，并用 workspace-path 助手构造转换副本地址。`face.ts` 为 slot 注册绑定 Remote、adapter、命令桥、这些 store 与打开 tab 清单。

`adapter.ts` 负责运行时边界：`OfficeAdapter`、`OfficeEditorHandle`、`createAdapter(runtimeFactory?)` 以及一个可控的测试 adapter。默认工厂动态导入 `editor/sheets.ts` 或 `editor/docs.ts`；这些模块用 `UniverSheetsCorePreset` 或 `UniverDocsCorePreset`、各自的 en-US locale 默认导出与 CSS 构造 `createUniver`，从线缆快照创建单元，并从 preset 的命令信号报告变更。`editor/sheet-operations.ts` 在运行任何命令之前，把每一种已声明的步骤类型（cell、formula、row、column 与 style）校验并翻译为 Univer 命令，因此被拒绝的步骤不会改动工作簿。没有模块在顶层导入 Univer，因此 jsdom 测试永远不会加载它。撤销与重做使用 preset 头部的标准控件；handle 转发同一套原生历史。

`tsconfig.host.json` 与 `tsconfig.client.json` 分别编译两个插件面，`tsdown.config.ts` 产出浏览器 bundle 以及按需请求的 sheets 与 docs chunk。

</details>

-----

<a id="further-exploration"></a>
## 进一步探索

- [ui-sidebar-right](../ui-sidebar-right/README.zh.md) —— 右侧 Sidebar 及其 tab 类型注册表。
- [ui-sidebar-documentpreview](../ui-sidebar-documentpreview/README.zh.md) —— 编辑器旁的只读文档预览。
- [Client group](../README.zh.md) —— 该组中的浏览器包。

-----

<a id="model-experience"></a>
## 模型体验

### 编辑器表面

#### 模型看到什么

没有内容。浏览器插件的 `apply` 不注册 tool schema、prompt 区段或 tool 结果文本。

#### Token 影响

零直接 token；这个包不向模型请求贡献任何内容。

#### KV Cache 影响

无；这个包既不组装也不发送 provider 请求。

## 已知限制与延后工作

<a id="known-limitations-and-deferred-work"></a>

- 真实的 Univer 挂载需要浏览器（canvas、ResizeObserver、preset 样式）；单元测试替换运行时工厂，真实挂载由组装后的浏览器冒烟负责。
- 关闭警告是 Sidebar 否决移除后显示在面板内的横幅，不是原生模态框。正文在隐藏期间保持挂载，因此该 tab 被激活时会看到待处理的提示。
- 编辑器界面只携带 preset 的 en-US locale。
- `.ppt` 转换出的副本是 `.pptx`，在文档预览中仍为只读，因此演示文稿不会在此编辑器中打开。
- 转换副本提示来自本窗口的插件状态，打开文档列表来自 Sidebar 的打开 tab 清单。
- 恢复由 Session 中的第一个触发点执行：挂载的 Office 正文，或该 Session 成为主视图 Session。激活路径读取 Session 列表缓存的 `document` projection，并通过 Sidebar 的按 Session 打开路径打开资源，因此不需要 Office tab。Sidebar 会持久化每个 Session 的 tab 布局，因此重启通常也会挂载这些正文。
- 单次恢复记录按 Session 保存且只存在于本窗口；恢复执行后才到达的 projection 不会重新打开任何内容。
- 恢复草稿由 `officeDocument` Remote 存放在 `<workspaceRoot>/.dsh/office-tmp/drafts/`，只在文档打开时读取。提供恢复的时机是每次打开 tab，因此外部变化后的重新读取不会再次检查草稿。
- 移动检测需要文件 feed 对新路径的另一次观察，例如在第二个 tab 中打开它。如果没有任何打开的 feed 报告重命名，tab 会保留在旧路径，直到文档重新打开。
- 内容版本由 Host 记录的保存、agent command 与本地保存共享。Host 自身保存触发重载时，会采用 Host 报告的版本，而不是把变化呈现为外部编辑；有未保存编辑的单元仍会询问。该 tab 已应用的 agent command 之后的保存会在不重新读取、不重新挂载的情况下推进该版本，因此编辑器的原生撤销历史得以保留。
- Univer core 与 presets 使用 Apache-2.0；这个包把 `@univerjs/core`、`@univerjs/presets`、`@univerjs/preset-sheets-core` 与 `@univerjs/preset-docs-core` 固定在 1.4.0 供编辑器表面使用。

<a id="dev-note"></a>
### 开发备注

<details>
<summary>维护者工作上下文——点击展开</summary>

无。

</details>
