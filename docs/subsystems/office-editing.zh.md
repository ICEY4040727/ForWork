# Office 编辑

[English](office-editing.md) | 中文

Office 编辑把 Session 中打开的 Word 或 Excel 文件变成实时编辑单元，在多次编辑之间保持该单元的身份与版本，并让智能体通过工具而不是直接操作文件来读取和编辑该单元。[office-exchange](../../packages/document/office-exchange/README.zh.md) 在文件与单元快照之间转换，并承载 `officeDocument` Remote；[session-document](../../packages/session/session-document/README.zh.md) 拥有文档身份与 `document/*` Session 事实；[tool-office](../../packages/document/tool-office/README.zh.md) 注册智能体工具；[ui-office-editor](../../packages/client/ui-office-editor/README.zh.md) 渲染承载实时编辑单元的右侧 Sidebar tab。[Office Exchange](office-exchange.zh.md) 拥有 seam 词汇、格式分类、Remote 方法与挂载上限；本页说明各参与方如何协作。

## 所有权

| 所有者 | 职责 |
|---|---|
| [office-exchange](../../packages/document/office-exchange/README.zh.md) | `ctx.officeExchange` 转换 seam、`LocalOfficeExchange`、`officeDocument` Remote 与 `officeCommandBroker` 命令注册表 |
| [session-document](../../packages/session/session-document/README.zh.md) | `DocumentId` 绑定、版本守卫的变更记录、`document/*` 事实写入器、`document` 投影，以及打开集合恢复与草稿恢复规则 |
| [tool-office](../../packages/document/tool-office/README.zh.md) | `document_read`、`document_edit_cells`、`document_edit_paragraph` 与 `document_convert` |
| [ui-office-editor](../../packages/client/ui-office-editor/README.zh.md) | `office` Sidebar tab、Univer adapter、保存与冲突处理、草稿调度与命令应答 |
| [ui-sidebar-right](../../packages/client/ui-sidebar-right/README.zh.md) | tab 类型注册、每 Session 的 tab 布局，以及未保存关闭的否决 |
| [ui-sidebar-documentpreview](../../packages/client/ui-sidebar-documentpreview/README.zh.md) | 只读预览，包括通过 Host 渲染 PDF 呈现的 `.pptx` |

## 格式与可编辑集合

`.docx` 与 `.xlsx` 可编辑：编辑器把它们挂载为 Univer workbook 或 document 单元，保存时通过 Exchange 提供方序列化回文件。旧格式 `.doc`、`.xls` 或 `.ppt` 源文件先转换，再打开副本；原文件绝不会被写入。`.pptx` 永远不会成为可编辑单元：`detectUnsupported` 可以检查演示文稿，`read` 与 `write` 以 `unsupported-format` 拒绝，只读文档预览则把它渲染为 Host PDF。

一次读取会报告在编辑并保存后会丢失的全部 `UnsupportedFeature`。该列表非空期间，编辑器与工具都会拒绝保存，而不是写入有损文件。[Office Exchange](office-exchange.zh.md) 定义格式联合、单元快照与特性类别。

## 身份、版本与守卫保存

文档的 Session 级身份是 `<sessionId>:<path>`，由 `documentIdFor` 生成。身份在 `document/rebind` 与 `document/missing` 之后保持不变；路径是可能变化的定位符，文件缺失时文档仍留在打开集合中并保留其历史。

两个版本值回答不同的问题。`officeDocument` 读取返回不透明的文件系统版本，`save` 把它作为 `expectedVersion` 传回：文件系统的 `replaceIfVersion` 守卫仅在文件仍持有该版本时发布写入，否则 Remote 抛出 `office-document/conflict`。数字 `contentVersion` 在文档打开时为 `0`，每次接受的保存加一；智能体命令用 `fromVersion` 与它比较，`document` 投影报告它。

文件在外部变化时，编辑器会重新加载干净的 tab；脏 tab 则在丢弃编辑前提示。该 tab 已应用的 agent command 之后的保存不是外部变化：实时单元已持有被保存的快照，因此 tab 原地采用已发布的文件版本，编辑器的原生撤销历史保留这次智能体编辑。保存没有无条件回退，带有不支持特性的文档既不能从编辑器保存，也不能从工具保存。

## 数据路径

| 流程 | 顺序 |
|---|---|
| 打开 | `officeDocument.read` 把格式、快照、不支持特性与文件版本返回给编辑器 adapter；本进程中文档的第一次读取只追加一次 `document/open` |
| 用户保存 | 脏单元携带其读取版本与 `user` actor 调用 `officeDocument.save`；守卫写入追加 `document/save` |
| 智能体读取 | `document_read` 通过 `officeCommandBroker` 派发一条 `read` 命令；实时 tab 以快照与数字版本应答，工具返回有界摘要 |
| 智能体编辑 | 工具按身份中的工作区相对路径读取文件以计算逆操作，把步骤作为一条命令派发，在收到 `applied` 应答后以 `agent` actor 与打开 turn 执行一次守卫保存（追加 `document/save`），并追加携带步骤与逆操作的 `document/edit`，随后报告新的数字版本 |
| 旧格式打开 | `convertLegacy` 在 `.dsh/office-tmp/` 下写入可编辑副本；tab 打开该副本并显示提示，原文件保持不变 |
| 草稿恢复 | 空闲的脏单元调用 `writeDraft`；下次打开调用 `readDraft` 并应用恢复决策 |
| 恢复 | 缓存的 `document` 投影提供 `restoreSet`，tab 按投影顺序重新打开；缺失文档通过提示报告 |

## Session 事实与 document 投影

七个 `document/*` 事件承载持久记录：`open`、`save`、`edit`、`rebind`、`missing`、`close` 与 `command`。它们是仅追加的。`document/command` 记录请求而非变更；每次接受的写入追加带 actor 与致因 turn（若有）的 `document/save`，接受的智能体变更还会追加携带精确逆操作的 `document/edit`。打开身份是工作区相对定位符，因此同一文件的绝对与相对拼写是同一个文档。`session-document` 定义 payload 与事实写入器。

`document` 投影把这些事实折叠为 `{ order, documents }`：显示顺序，以及每个打开文档的一条条目，包含其身份、当前路径、格式、数字内容版本、`open` 或 `missing` 状态与不支持特性。`document/open` 插入或更新条目，`document/save` 与 `document/edit` 推进其版本，`document/rebind` 替换路径，`document/missing` 标记为不可用，`document/close` 删除它；无关事件返回同一个 state 引用。该插件还把打开集合渲染为 prompt 区段，并向 Client 暴露 `restoreSet` 与 `recoveryDecision`。[Session Projection](session-projection.zh.md) 拥有投影 seam，[persistence catalog](../persistence-catalog.zh.md) 承载生成的事件类型。

## Sidebar 编辑器

`office` tab 类型认领 basename 以 `.docx`、`.xlsx`、`.doc`、`.xls` 或 `.ppt` 结尾的 Session 级 `dsh-resource://file/session/<sessionId>/<path>` 地址；`.pptx` 被有意排除。tab 在切换时保持正文挂载，因此实时编辑在导航后仍然存活。正文通过 `officeDocument.read` 读取，并在 adapter 后挂载 Univer Sheets 或 Docs preset。

脏状态驱动三种行为：一次编辑在 2 s 内没有后续编辑后调度一份恢复草稿；Sidebar 关闭被否决，随后面板提供保存并关闭、丢弃并关闭或取消；外部变化在丢弃前提示。`DocumentPanel` 列出 Session 打开的 Office tab，并在点击时聚焦其中一个。旧格式打开会转换源文件，并在原 tab 位置重新打开副本，同时显示标出原文件的可关闭提示。

恢复在每个 Session 每窗口运行一次，触发条件是第一个挂载的 Office 正文或 Session 成为主视图 Session。它读取缓存的 `document` 投影，并通过 Sidebar 的每 Session 打开路径按投影顺序打开每个文档；每个地址都绑定到投影的 `documentId`，因此重绑定到新路径的文档会按 Session 记录的身份重新打开。投影报告缺失的文档出现在可关闭提示中，而不是一个失效的 tab。

## 智能体命令桥

读取与编辑工具从不直接修改实时单元。读取派发一条 `read` 命令；编辑先通过 `officeDocument.read` 读取文件以计算逆操作，再派发步骤。提供方的 `officeCommandBroker` 追加一条 `document/command` 并等待实时 tab；tab 恰好应答一次，状态为 `read`、`applied`、`pending-edits`、`conflict`、`unsupported-feature` 或 `document-not-open`。超时的请求不写入任何内容。broker 只接受记录了该请求的 Session 的应答，并忽略且记录未知、迟到或来自其他 Session 的应答。

接受的智能体编辑只保存一次。工具校验编辑器的数字版本为 `fromVersion + 1`，用命令前观察到的文件版本保存返回的快照，把保存记录为带打开 turn 的 `actor: 'agent'`，并追加携带步骤与逆操作的 `document/edit`。拒绝不会改变文件与日志。[tool-office](../../packages/document/tool-office/README.zh.md) 列出面向模型的结果代码。

## 草稿与恢复

单元为脏期间，每次编辑替换一个待写草稿候选，调度器在 2 s 内没有后续编辑后写入它。`writeDraft` 把线缆快照及其捕获自的文件版本存储到 `<workspaceRoot>/.dsh/office-tmp/drafts/<sha256(documentId)>.json`；它从不写入文档路径。重新打开文档时只读取一次草稿：捕获自当前文件版本的草稿直接提供，较旧的草稿标记为过期，仅在用户选择时应用。成功保存与丢弃都会删除草稿，且没有逻辑会清理仍保持关闭的文档的草稿。

## 旧格式转换

`convertLegacy` 在 Session 工作区内运行。它创建 `<workspaceRoot>/.dsh/office-tmp/convert-<hash>-<random>/` 并把可编辑副本写入其中，因此同一 Remote 可以在不变的根检查下读取结果。每次调用获得独立目录，转换失败会删除它。原文件绝不会被修改。编辑器在 tab 中打开可编辑副本并显示提示；转换后的 `.ppt` 变为 `.pptx`，留在只读预览中。

## 组合

[base bundle](../../packages/bundle/base/README.zh.md) 挂载 `office-exchange`、`session-document` 与 `tool-office`。[Web 应用 bundle](../../packages/bundle/web-app/README.zh.md) 挂载 `ui-office-editor`，[api/remotes](../../packages/api/remotes/README.zh.md) 暴露生成的 `remote.officeDocument` 命名空间。`tool-office` 需要提供方的命令 broker：挂载 Remote 而不挂载 broker 的组合仍可读取文件，但每条命令都以 `conversion-failed` 拒绝。`ctx.officeExchange`、`ctx.officeDocument` 与 `ctx.officeCommandBroker` 的生成服务参考位于 [Office Exchange](office-exchange.zh.md#cordis-surface)。
