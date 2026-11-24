---
description: "Office Exchange 能力 seam：OOXML 与单元快照的映射、原地写入源包、保留与耦合构造拒绝，以及由 Office 编辑器与 agent 文档工具共享的旧版转换。"
kind: "package-reference"
---

# @forwork/office-exchange

[English](README.md) | 中文

## 概述

定义把 Office 文件移入可编辑单元模型并写回其来源包的契约。该 seam 命名转换、保留、拒绝与保存处置；提供方负责 ExcelJS、docx 与 LibreOffice 的具体机制。编辑器与 agent 工具消费 `ctx.officeExchange`，不直接导入转换库。

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
  maxSourceAnchors: 100_000,
  maxSourceAnchorBytes: 8 * 1024 * 1024,
})

const { format, snapshot, unsupported } = await ctx.officeExchange.read(path, signal)
```

`read(path)` 返回检测到的格式、可编辑的单元快照，以及整包重新生成会丢失的每个特性。`write(snapshot, format, source?, signal?)` 把快照编码回文件字节但不写入，并返回 `{ bytes, disposition }`；`source` 是调用方按期望版本读取的源包，它与快照 sidecar 同时存在时选择原地写入器。`detectUnsupported(path)` 报告易丢失特性，不把内容读入编辑器。`detectPreserved(bytes)` 为调用方已持有的字节分类出原样保留类别与耦合构造标记。`convertLegacy(path, signal?, outputDir?)` 把 `.doc/.xls/.ppt` 文件转换为可编辑的等价文件；省略 `outputDir` 时写入任意工作区之外的新临时目录。

支持的格式：

- `.docx`、`.docm`、`.xlsx` 与 `.xlsm` 可编辑：`read` 把它们映射为 Univer 文档与工作簿快照，`write` 把被编辑单元拼接回源包。原地保存会保留宏启用文件的 VBA 项目与宏启用的主内容类型。
- `.doc`、`.xls` 与 `.ppt` 必须先经 `convertLegacy` 转换再 `read`；直接读取会报 `unsupported-format`。
- `.pptx` 仅供预览：`detectUnsupported` 检查它，而 `read` 与 `write` 报 `unsupported-format`。

每个失败都是带稳定 `code` 的 `OfficeExchangeError`；边界在挂载时解析一次，无效值会显式报错。

`formatOf(path)` 按扩展名分类路径；`isEditableFormat` 区分可直接编辑的 `.docx/.docm/.xlsx/.xlsm` 与需要转换的格式；`isLegacyFormat` 指出旧版格式集合。

### 原地保存

当快照携带 `data.officeExchange.source` 且调用方提供源字节时，`write` 从源包组装输出。编辑未改变的每个条目都从其源 ZIP 本地记录复制，因此压缩数据、CRC、时间戳、扩展字段、压缩方法、条目顺序与数据描述符保持逐字节一致。未改变任何单元的保存会原样返回源字节。

在部件内部只有被编辑单元的 XML 区间会改变：一个段落、一个表格单元格或一个工作表 `<c>` 元素。正文级内容控件锚定其整个 `w:sdt` 元素，因此编辑该单元会被拒绝，而相邻编辑不会触碰它。编辑区间之外的每个字节都从源文件复制。工作表行或列的插入与删除表示为行或单元格区间操作，而不是整表重写；当工作表携带合并区域、条件格式区域、数据验证区域、autoFilter 区域或 Excel 表格地址区域时，该移动会被拒绝并指出该构造，因为这些区域的坐标不会随被移动的行或列一起变化；声明的表格部件若其关系、目标部件或 `ref` 无法解析为区域，该移动也会被拒绝，而不是留下过期的表格引用。必需的样式、编号定义、关系或内容类型声明会追加到所属部件的结束标签之前，被编辑的单元格样式只追加其所需的样式组件；不会重新生成任何辅助部件，`word/settings.xml` 从不写入，工作表字符串改动写成内联字符串，使 `xl/sharedStrings.xml` 保持不动。

源 sidecar 记录 `schemaVersion: 1`、打开时的 `format`、源 `packageDigest` 与 `byteLength`、`anchorMode`、`sourceParts`，以及逐单元 `anchors` 或逐工作表 `sheets` 之一。每个锚点给出单元 kind、所属部件、模型 key、未压缩部件中的字节区间，以及规范单元的摘要。它经过编辑器的快照序列化、JSON 线缆与恢复草稿后仍然存在。`maxSourceAnchors`（默认 100,000）与 `maxSourceAnchorBytes`（默认 8 MiB）约束 payload；超过逐单元格数量的工作簿存储工作表摘要并在保存时重建单元格锚点，无法锚定的源文件会在没有 sidecar 的情况下读取。

当后续保存提供的源与 sidecar 的摘要或长度不一致时，写入器会在规划前重新锚定该源，因此无需重新读取也能连续保存两次。无法重新锚定的源会被拒绝。没有 sidecar 的快照，或未提供源的调用方，使用旧版重新生成写入器；该写入器从模型重建整包，并报告不声称任何逐字条目的处置。`.docm` 或 `.xlsm` 目标会拒绝该路径，因为重新生成会写出不含 VBA 项目的无宏包。

### 保留与拒绝

`read` 在 `preserved` 中报告原样保留的构造类别，在 `unsupported` 中报告整包重新生成会丢失的内容；两份列表都不会阻止保存。原地写入器在被编辑单元之外原样携带这些类别：宏、图表、数据透视表、Excel 表格、定义名称、计算链、批注、嵌入对象、绘图、形状、主题、字体表、Web 设置、自定义 XML、外部链接、域、书签、内容控件与文档属性。

当编辑与含义或锚点依赖被改内容的构造相交时，保存会被拒绝：覆盖该单元的修订或批注区间、包裹被编辑文字的内容控件、拥有该文字的域、富文本共享字符串、共享或数组公式、改动的工作表超链接，或含义不明的 `mc:AlternateContent` 分支。写入器无法定位或表达的区间，以及它无法表示的样式组件，也会被拒绝。拒绝会指出原因及其位置，不写入任何内容，也不记录保存事实。与改动无关的未映射属性会重新输出在被拼接的单元内，而不是触发拒绝，包括单元格样式索引、表格单元格与行属性、未映射的段落与文字属性、单元格宽度与适应文字。

只有当工作表携带上述地址区域时，结构性的行或列移动才会被拒绝；没有锚定源的宏启用目标会拒绝保存，而不是退回重新生成写入器。文件中其他位置的原样保留构造绝不会拒绝编辑。

### 保存处置

`write` 返回 `SaveDisposition`：`schemaVersion: 1`、精确的 `verbatimCount`，以及 `spliced`、`added` 与 `removed` 条目各一个 `{ count, names }` 组。计数是精确的；每个 `names` 列表都是排序后的 ASCII，最多 256 个。Remote 把同一个值作为可选属性记录到 `document/save`，因此较早的保存或旧版重新生成会读作缺席而不是无效。

### 不支持特性的检测

`read` 与 `detectUnsupported` 报告整包重新生成会丢失的每个构造。对原地保存而言该列表是信息性的：携带该构造本身不会导致拒绝，只有与编辑相交的耦合构造，或写入器无法保留的区间，才会拒绝。

- 文档：宏、修订记录、批注、图表、绘图、形状与嵌入对象。
- 工作簿：图表、批注部件与非图片绘图部件、宏与嵌入对象。

文字格式（字体、字号、颜色、突出显示或底纹、删除线、全大写、小型大写、垂直对齐、各种下划线以及文字边框）、段落格式（对齐、间距、行距、缩进、制表位、大纲级别、底纹以及段落边框）、`word/styles.xml` 的文档默认值、粗体、斜体、下划线、标题样式、节页面设置（页面大小、方向、页边距、页眉/页脚距离、分栏以及分节类型）、带默认、首页与偶数页引用的页眉和页脚文字、带 `word/numbering.xml` 定义与逐段 `w:numPr` 的列表、带外部关系目标或 `w:anchor` 书签的超链接、表格结构（合并单元格、单元格边框、底纹、垂直对齐与边距、行高与标题行重复、表格宽度、对齐、缩进与布局，以及单元格内的嵌套表格）、数字格式、填充、工作簿单元格边框、水平与垂直对齐、自动换行与缩小字体填充、文字旋转、删除线、上标与下标、冻结窗格、工作表缩放、从右到左视图、行列标题可见性、工作表标签颜色、深度隐藏的工作表状态、行与列样式、合并区域、行高与列宽都能往返，因此检测不会报告它们。

工作簿条件格式、数据验证、单元格超链接、单元格图片与工作表背景图片都会映射进表格插件模型并在保存时写回。条件格式规则、验证规则与绘图通过 `SHEET_CONDITIONAL_FORMATTING_PLUGIN`、`SHEET_DATA_VALIDATION_PLUGIN` 与 `SHEET_DRAWING_PLUGIN` 资源持久化在工作簿快照中；超链接目标持久化在所属单元格的富文本自定义区间里。检测不再报告这些内容。

文档图片与脚注/尾注会被保留而不是映射：每个绘图、图片或注释引用所在的 run 都与其媒体部件、关系、注释部件和内容类型声明一起逐字节捕获，并在写入时原样写回。它们不会在编辑器中渲染，也不会成为可编辑内容；周围文字仍然可编辑。

文件以“携带所有原样保留构造、而非拒绝”的方式打开。常见的文字与段落格式会映射进文档模型并在保存时写出，因此编辑会保留它们而不会被拒绝；与含义或锚点依赖被改内容的构造相交的编辑会被拒绝并指出该构造。

### Remote 表面

该插件还挂载 `ctx.officeDocument`，一个在同一 seam 之上的 Session 作用域 Remote 表面。它生成的 Client 命名空间是 `remote.officeDocument`；把该包挂载进 Host 组合，并把 `@forwork/office-exchange/remote` 加入 Client 组装。

该 Remote 从请求头派生的作用域解析实时 Session，并在其中记录 `document/*` 事实；`session-document` 插件把它们折叠为打开集合。

- `read(workspaceFileScope, path)` 从 Session 的工作区根解析路径，用 `assertWithinRoots` 限制范围，在提供方的 `maxFileBytes` 边界内通过 filesystem seam 读取，并返回 `{ format, snapshot, preserved, unsupported, version }`。本进程中某文档的第一次读取会追加一次 `document/open`；之后的读取不更新任何事实。打开事实记录工作区相对定位符，因此同一文件的绝对与相对拼写共享同一个文档身份。
- `save(workspaceFileScope, path, snapshot, format, expectedVersion, attribution)` 要求文件仍持有 `expectedVersion`，在字节边界内读取其字节，再次校验版本，并从这些字节通过提供方写入。与耦合构造相交的写入以 `office-document/unsupported-feature` 拒绝并指出该构造，不追加任何事实；自读取以来发生变化的文件会以 `office-document/conflict` 拒绝，写入从不退化为无条件保存。输出通过文件系统的 `replaceIfVersion` 守卫发布。被接受的保存追加带 `attribution.actor`（编辑器保存为 `user`，文档工具为 `agent`）、下一内容版本与写入器处置的 `document/save`；智能体在 turn 内保存时还会记录 `attribution.turn`。内存中的改动留在编辑器内，因此不记录操作步骤或逆操作。
- `closeDocument(workspaceFileScope, documentId, reason)` 追加 `document/close`，并丢弃进程本地的打开条目。Host 重启后关闭的 tab 仍会留下持久事实。
- `rebindDocument(workspaceFileScope, documentId, toPath)` 追加 `document/rebind` 并保留身份；旧路径与新路径都必须解析到工作区根内。
- `markDocumentMissing(workspaceFileScope, documentId)` 追加一次 `document/missing` 并保持文档打开；之后的保存、编辑或重绑会把它重新标记为可用。
- `detectUnsupported(workspaceFileScope, path)` 报告易丢失的特性，不返回内容。
- `convertLegacy(workspaceFileScope, path)` 返回 `<workspaceRoot>/.dsh/office-tmp/convert-<hash>-<random>/` 内的转换后文件路径，因此同一个 Remote 可以在不变的根检查下 `read` 它。每次调用获得自己的目录。
- `writeDraft(workspaceFileScope, documentId, documentPath, contentVersion, snapshot)` 持久化一份恢复草稿并返回 `{ savedAt }`。草稿位于 `<workspaceRoot>/.dsh/office-tmp/drafts/<sha256(documentId)>.json`；摘要确保调用方选择的 id 无法命名草稿目录之外的路径。它替换已有草稿，且从不写入文档路径本身。
- `readDraft(workspaceFileScope, documentId)` 返回 `{ documentId, path, contentVersion, snapshot, savedAt }`，文档没有草稿时返回 `null`。格式错误的草稿以 `office-document/conversion-failed` 拒绝。
- `clearDraft(workspaceFileScope, documentId)` 删除草稿；对没有草稿的文档执行清除是无操作。
- `completeCommand(workspaceFileScope, requestId, result)` 用实时编辑器的答案结算一个待处理的关联命令。该提供方还挂载 `ctx.officeCommandBroker`，这个服务拥有待处理请求、其超时（`commandTimeoutMs`）以及每个文档最后观察到的数字内容版本。未知、迟到或属于其他 Session 的结果会被忽略并记录日志，绝不应用。 `registerCommandTarget(workspaceFileScope, documentId, registrationId)` 与 `unregisterCommandTarget(workspaceFileScope, documentId, registrationId)` 记录哪些实时 tab 可以应答：没有已注册目标的文档会立即以 `document-not-open` 答复，只有已注册但保持沉默的目标才会进入超时。

`completeCommand` 应答由 Host 工具层追加的 `document/command` Session 事件；此包本身不追加该事件。结果联合为 `read`、`applied`、`pending-edits`、`conflict`、`unsupported-feature` 或 `document-not-open`；读取或已应用的答案携带实时快照与数字内容版本，冲突携带 `currentVersion`。Host 只在 `applied` 答案后写文件。

草稿与版本绑定：记录携带快照捕获自的 `contentVersion`，因此恢复路径可以在提供恢复前与文件的当前版本比较。草稿目录与解析出的草稿文件都通过 `assertWithinRoots`，指向工作区根之外的链接会被拒绝，而不是穿透写入。

`OfficeDocumentSnapshot` 是带快照 kind 标记的 JSON 投影：Univer 的原生接口允许无约束自定义字段，JSON 线缆无法承载。Host 在序列化前验证该投影，并把每个失败映射到 `office-document/*` Remote 代码（`unsupported-format`、`unsupported-feature`、`out-of-scope`、`bound-exceeded`、`conversion-failed`、`conflict`）。

### 失败策略

失败是带稳定 `code` 的 `OfficeExchangeError` 值。无法打开文件的读取报告带路径的 `conversion-failed`，而不是泄漏 Node 错误。读取时非空的 `unsupported` 列表意味着整包重新生成会丢失内容；原地保存会携带它，只有与编辑相交的耦合构造才会拒绝。ZIP64、加密、分卷或其他不受支持的 OOXML 容器报告 `conversion-failed` 且不写入任何内容。旧格式转换超过 `maxConversionMs` 时会中止引擎并报告 `too-large`，源文件保持不变。

-----

<a id="understand-the-implementation"></a>
## 理解实现

<details>
<summary>实现内部细节——点击展开</summary>

`src/service.ts` 声明抽象的 `OfficeExchange` 服务、context 增强，以及由包入口重新导出的 seam 词汇。`src/types.ts` 保存 JSON 安全的线缆与 seam 词汇；`src/snapshots.ts` 保存 Univer 快照类型，使 Remote Client 永不加载编辑器核心。`src/remote.ts` 实现 `officeDocument` Remote 表面：Session 作用域解析、工作区限制、版本传递、JSON 投影、工作区作用域的转换目录、恢复草稿存储、命令分发与结算，以及其余的保存保证。`src/commands.ts` 拥有 `officeCommandBroker` 服务及其待处理请求注册表。该 Remote 还要求 `sessions` 服务：它只向活跃 Session 追加文档事实。`src/detect.ts` 分类原样保留构造与耦合标记。`src/package/` 承载 Host-only 的 ZIP/OPC 层：读取器索引本地与中央记录，写入器复制未改动记录并重建中央目录，`ooxml-package.ts` 暴露部件映射、dirty 集合与无操作快速路径。`src/source/` 承载源区间层：命名空间感知的字节扫描器、单元摘要、sidecar 校验与拼接计划。`src/provider/index.ts` 在纯模块 `provider/xlsx.ts`、`provider/docx.ts`、`src/detect.ts` 与 `provider/legacy.ts` 之上实现 `LocalOfficeExchange`：它拥有文件读取、格式分发、写入器选择、类型化失败、调用方指定的旧版转换目录，以及经验证的 `Config` schema。Office 编辑器与 agent 工具仍是消费方；此包不注册任何工具，Remote 唯一的 Session 写入就是上述 `document/*` 事实。

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
- 原地保存需要源包。没有 `data.officeExchange.source` 的快照，或未提供源字节的调用方，使用旧版重新生成写入器；该路径从模型重建整包，不声称任何逐字条目。`.docm` 或 `.xlsm` 目标会拒绝该路径，因为重新生成会丢弃 VBA 项目并写出无宏包。
- ZIP64 归档、加密条目、分卷归档，以及存储或 deflate 之外的压缩方法以 `conversion-failed` 拒绝，而不是重写。文件大小、单元格数量、条目数量与锚点 payload 边界失败时不写入。
- 辅助部件仅追加：必需的 `w:style` 追加到 `word/styles.xml`，必需的列表定义追加到 `word/numbering.xml`，必需的关系或内容类型声明追加到其结束标签之前。`word/settings.xml` 从不写入。被编辑的单元格样式会复用匹配的 `<xf>`，或只向 `xl/styles.xml` 追加所需的 `<numFmt>`、`<font>`、`<fill>`、`<border>` 与 `<xf>` 条目，并更新该集合的 `count`；现有组件字节保持不变。写入器无法表达的模型样式组件、源 `<xf>` 扩展、主题颜色引用、对角线边框、空的自闭合样式表，或没有 `xl/styles.xml` 的包会拒绝并指出该组件。改动单元格超链接会被拒绝，因为工作表关系不会被重新生成。
- Word 表格行或列尺寸变化、Word 单元重排，或无法定位或含义不明的源区间会被拒绝并指出构造。被编辑单元内独立的未映射属性会被携带：单元格样式索引、表格单元格与行属性、未映射的段落与文字属性、单元格宽度与适应文字。
- Office 数学公式（`m:oMath` 与 `m:oMathPara`）在编辑发生在别处时逐字节保留，但不会被映射，也不可编辑；编辑包含内联公式的段落会以 `unsupported-feature` 拒绝并指出该数学元素。
- 工作表行或列移动在工作表携带合并区域、条件格式区域、数据验证区域、autoFilter 区域或 Excel 表格地址区域时会被拒绝并指出该构造；声明的表格部件若其关系、目标部件或 `ref` 无法解析，也会因同样的原因拒绝移动，而不会留下过期引用；普通工作表按行或单元格区间操作移动。
- 保存的文件从源本地记录组装，并作为一个完整文件通过守卫替换发布；不会在磁盘上原地修补。
- 节的页面边框、页码、文档网格、行号以及竖排文字方向不会被映射。页眉和页脚中的绘图与域也不会被映射；页眉和页脚图片与正文图片一样以不透明方式保留而不渲染，其中的列表与超链接采用与正文相同的映射。脚注和尾注正文作为完整部件以不透明方式保留，因此其文字不可编辑。原地保存会保留源文件携带的每个 `word/numbering.xml` 定义，包括没有任何 `w:num` 引用的定义；重新生成写入器会丢弃未被引用的抽象编号定义，因为它没有可挂接的列表实例。超链接关系目标始终按外部目标写出。
- 表格单元格的适应文字、隐藏标记与文字方向属性、`w:tblCellSpacing`、行的 `w:gridBefore`/`w:gridAfter` 偏移、浮动表格位置以及表格样式的条件格式不会被映射进编辑器模型。原地保存会在未编辑内容中保留它们，并在被编辑单元中重新输出与改动无关的部分。单元格的 `w:insideH`/`w:insideV` 边框仅在缺少对应外边时作为回退使用，并会按显式单元格边写出；单元格边框的线型变体会归入模型的实线、点线或虚线集合；自动表格宽度会按其解析后的网格宽度写出。
- 主题部件与由主题派生的样式不会被映射：没有直接取值的主题字体或颜色引用不会成为可编辑格式。原地保存原样复制主题部件与现有样式定义，只追加它需要的样式定义；重新生成写入器只重建 Normal 与标题样式。引用某样式的段落会保留其 `pStyle` id，其直接的文字与段落格式仍会保留。
- 工作簿的行与列大纲级别在 Univer 模型中没有对应字段，不可编辑；重新生成写入器会丢弃它们，原地保存则在被编辑单元格区间之外复制工作表。工作表滚动位置与打印设置同理。Excel 没有独立的“裁剪”换行标志，因此模型中使用裁剪策略的单元格会按不换行写出；Excel 的 `fill` 以及垂直 `distributed`/`justify` 对齐会映射到最接近的 Univer 对齐。
- 唯一值、重复值与平均值条件格式在 ExcelJS 中没有写入器，因此重新生成写入器会写出等价的 `COUNTIF` 或 `AVERAGE` 表达式。Excel 内置集合之外的图标集不会渲染图标。ExcelJS 不保留规则的文本属性、命中即停（stop-if-true）以及数据条的渐变、边框、方向与轴选项，因此这些选项不可编辑。
- ExcelJS 读取时不保留超链接的屏幕提示，因此被编辑的链接只有 URL 与显示文字能往返；对另一单元格的原地保存会复制未触碰的链接。列表、整数、小数、日期、时间、文本长度、自定义与任意以外的数据验证类型（复选框与多选列表）在保存时会被拒绝，而不是被丢弃。
- 单元格图片以 base64 表格绘图保存，并按单元格锚点写回；工作表背景使用相同存储。绝对锚点图片与非图片绘图元素作为 `drawing` 内容原样保留，保持不可编辑。png、jpeg、gif 之外的媒体会随扩展名保留，但重新生成写入器会把内容类型写成 `image/<extension>`，部分应用可能无法识别。
- `convertLegacy` 把转换后的可编辑文件保留在 `<workspaceRoot>/.dsh/office-tmp/` 下，`officeDocument` Remote 可以在同一根检查下读取。失败的转换会删除它刚创建的目录；成功的转换会保留到用户或 Session 删除为止，目前没有任何机制自动清理。
- 恢复草稿会累积在 `<workspaceRoot>/.dsh/office-tmp/drafts/` 下，直到 `clearDraft` 删除；目前还没有任何机制为不再打开的文档清理草稿。
- Client 通过 `markDocumentMissing` 报告文件缺失，但目前没有 Client 触发器检测移动：`rebindDocument` 暴露在 Remote 表面上却没有自动调用方。打开集合是进程本地的，用于去重，因此 Host 重启后下一次读取会再次追加 `document/open`；projection 保留文档的位置与版本。
- 该 Remote 只为活跃 Session 记录事实。为未加载 Session 执行的读取以 `gateway/bad-request` 拒绝，而不是在没有持久化保证的情况下读取。

<a id="dev-note"></a>
### 开发备注

<details>
<summary>维护者工作上下文——点击展开</summary>

许可检查（2026-11-18）：`@univerjs/core` 与四个表格特性插件（`@univerjs/data-validation`、`@univerjs/drawing`、`@univerjs/sheets-conditional-formatting`、`@univerjs/sheets-data-validation`、`@univerjs/sheets-drawing`、`@univerjs/sheets-hyper-link`）是 Apache-2.0，`exceljs` 是 MIT。`@univerjs/core` 保持为 peer 依赖，因为此包只导入其类型并在声明中输出；表格特性插件由编辑器 Client 挂载。`exceljs` 是 `.xlsx` 提供方的运行时依赖。

</details>
