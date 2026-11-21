# Office 交换

[English](office-exchange.md) | 中文

[document 包族](../../packages/document/README.zh.md) 在磁盘字节与 Office 编辑器及智能体文档工具使用的可编辑单元快照之间移动 Office 文件。[office-exchange](../../packages/document/office-exchange) 声明 `ctx.officeExchange` 能力 seam，提供 `LocalOfficeExchange` 提供方，并承载会话级的 `officeDocument` Remote。该包不注册面向模型的工具，也不追加 Session 事件。[Office 编辑器 Client](../../packages/client/ui-office-editor/README.zh.md) 通过 `remote.officeDocument` 读取和保存，面向模型的 [Office 文档工具](../../packages/document/tool-office/README.zh.md) 使用该 seam 与 Remote；[Office 编辑](office-editing.zh.md) 端到端说明该子系统。

## 所有权

| 所有者 | 职责 |
|---|---|
| [office-exchange](../../packages/document/office-exchange) | `ctx.officeExchange`：读取、写入、不支持特性检测与旧格式转换；`officeDocument` Remote |
| [api/remotes](../../packages/api/remotes/README.zh.md) | 挂载生成的 `remote.officeDocument` Client 命名空间 |
| [base bundle](../../packages/bundle/base/README.zh.md) | 以部署上限挂载 `office-exchange` |

## 格式与单元快照

`OfficeFormat` 为 `docx`、`docm`、`xlsx`、`xlsm`、`pptx`、`doc`、`xls` 或 `ppt`。可编辑集合是 `EditableOfficeFormat`（`docx`、`docm`、`xlsx`、`xlsm`）；旧格式集合是 `LegacyOfficeFormat`（`doc`、`xls`、`ppt`）。

| 格式 | 操作 |
|---|---|
| `.docx`、`.docm`、`.xlsx`、`.xlsm` | 可编辑：`read` 将文件字节映射为 `UnitSnapshot`；`write` 把被编辑单元拼接回源包。原地保存会保留宏启用文件的 VBA 项目与宏启用的主内容类型。 |
| `.doc`、`.xls`、`.ppt` | 旧格式：`convertLegacy` 生成可编辑文件；对原文件调用 `read` 以 `unsupported-format` 拒绝 |
| `.pptx` | 仅预览：`detectUnsupported` 可检查；`read` 与 `write` 以 `unsupported-format` 拒绝 |

`UnitSnapshot` 是编辑器原生的可序列化快照，是以 `{ kind: 'workbook', data: IWorkbookData }` 与 `{ kind: 'document', data: IDocumentData }` 构成的可辨识联合。`read` 返回 `ExchangeReadResult`，包含检测到的 `format`、解析出的 `snapshot`，以及整包重新生成会丢失的全部 `UnsupportedFeature`。每个 `UnsupportedFeature` 携带 `kind`——`macro`、`tracked-changes`、`comment`、`chart`、`drawing`、`shape`、`formula-recalc`、`embedded-object` 或 `other`——以及可选的、面向人类的 `detail`。非空列表描述源文件携带的内容，本身绝不拒绝保存。

`formatOf(path)` 按扩展名分类路径，`isEditableFormat` 收窄到 `docx`/`docm`/`xlsx`/`xlsm`，`isLegacyFormat` 收窄到 `doc`/`xls`/`ppt`。

## 原地保存

`ctx.officeExchange.write(snapshot, format, source?, signal?)` 返回 `OfficeWriteResult`（`bytes`、`disposition`），但不写文件。当快照带有 sidecar 且调用方提供源包时，写入器从源字节组装输出：编辑未改变的每个条目都从其源 ZIP 本地记录复制，因此压缩数据、CRC、时间戳、扩展字段、压缩方法、条目顺序与数据描述符保持逐字节一致。未改变任何单元的保存会原样返回源字节。

在编辑改变的部件内部，只替换被编辑单元的 XML 区间：一个段落、一个表格单元格或一个工作表 `<c>` 元素。正文级内容控件锚定其整个 `w:sdt` 元素，因此编辑该单元会被拒绝，而相邻编辑不会触碰它。编辑区间之外的每个字节都从源文件复制，工作表行或列的插入与删除表示为区间操作，而不是整表重写；当工作表携带合并区域、条件格式区域、数据验证区域、autoFilter 区域或 Excel 表格地址区域时，该移动会被拒绝并指出该构造，因为这些区域的坐标不会随被移动的行或列一起变化。必需的样式、编号定义、关系或内容类型声明会追加到所属部件的结束标签之前，而不是重新生成该部件；被编辑的单元格样式会复用匹配的 `<xf>`，或向 `xl/styles.xml` 追加所需的样式组件并更新其计数。`word/settings.xml` 从不写入，工作表字符串改动写成内联字符串，使 `xl/sharedStrings.xml` 保持不动。

`read` 在可编辑数据对象内的 `data.officeExchange.source` 附加源 sidecar。它记录 `schemaVersion: 1`、`format`、源 `packageDigest` 与 `byteLength`、`anchorMode`（`per-unit` 或 `sheet-summary`）、`sourceParts`，以及 `anchors` 或 `sheets` 之一。每个锚点给出单元 `kind`、所属 `part`、模型 `key`、未压缩部件中的 `start` 与 `end` 字节偏移，以及规范单元投影的摘要。该 sidecar 经过 `unit.save()`、线缆与恢复草稿后仍然存在，因此编辑器往返之后的保存仍能找到其锚点。

`maxSourceAnchors`（默认 100,000）与 `maxSourceAnchorBytes`（默认 8 MiB）约束该 sidecar。超过逐单元锚点数量的工作簿记录 `sheet-summary` 条目，并在保存时从源重建逐单元格锚点。无法定位单元、或锚点 payload 仍超出预算的源文件会在没有 sidecar 的情况下被读取，因此通过重新生成写入器保存。当保存的源与 sidecar 的 `packageDigest` 或 `byteLength` 不再一致时，写入器会在规划前重新锚定该源，因此刚写入文件的第二次保存无需重新读取即可工作。无法重新锚定的源以 `conversion-failed` 拒绝；文件系统版本守卫仍是冲突的最终权威。

没有 `data.officeExchange.source` 的快照，或未提供源包的调用方，使用旧版重新生成写入器。该写入器从可编辑模型重建整包，并返回不声称任何逐字条目的处置。`.docm` 或 `.xlsm` 目标会拒绝该路径，因为重新生成会写出不含 VBA 项目的无宏包。

## 保留与拒绝

`detectPreserved(bytes, config?)` 对源包分类，但不决定保存是否可以进行。`PreservedConstructKind` 命名原地保存在被编辑单元之外原样携带的内容：`macro`、`chart`、`pivot-table`、`excel-table`、`defined-names`、`calc-chain`、`comment`、`embedded-object`、`drawing`、`shape`、`theme`、`font-table`、`web-settings`、`custom-xml`、`external-link`、`field`、`bookmark`、`content-control` 与 `document-properties`。仅仅存在这些内容绝不拒绝保存。

`CoupledConstructKind` 命名含义或锚点依赖被编辑内容的构造：`tracked-change`、`comment-range`、`content-control-wrap`、`field-owned-text`、`rich-text-shared-string`、`shared-formula`、`array-formula` 与 `ambiguous-alternate-content`。写入器针对被编辑单元的区间进行规划；当其中之一与区间相交时，返回拒绝并指出构造及其部件或字节偏移。改动的工作表超链接也会被拒绝，因为工作表关系部件不会被重新生成。当工作表携带合并区域、条件格式 `sqref`、数据验证 `sqref`、autoFilter 区域或 Excel 表格地址区域时，行或列移动会以该构造被拒绝，因为这些区域不会随之移动。没有锚定源的 `.docm` 或 `.xlsm` 保存会以 `unsupported-feature` 拒绝，而不是重新生成无宏包。拒绝不写入任何内容、不记录保存事实，且源字节保持不变。这与原样保留类别不同：文件包含该构造，但用户仍可编辑别处。

被编辑区间内不影响改动内容的未映射属性会合并进拼接后的单元，而不是触发拒绝：单元格的 `s` 样式索引、`w:tcPr`/`w:trPr` 子元素、未映射的 `w:pPr`/`w:rPr` 子元素、单元格宽度与适应文字。只有当编辑会改变其含义或使锚点失效时，构造才是耦合的。

## Seam 操作与失败

`ctx.officeExchange.read(path, signal?)` 返回 `ExchangeReadResult`。`ctx.officeExchange.write(snapshot, format, source?, signal?)` 返回 `OfficeWriteResult`，但不写文件。`detectUnsupported(path, signal?)` 报告易丢失特性而不返回内容。`detectPreserved(bytes, config?)` 为调用方已持有的字节报告原样保留类别与耦合标记。`convertLegacy(path, signal?, outputDir?)` 返回可编辑文件的路径；省略 `outputDir` 时写入任意工作区之外的临时目录。

`OfficeExchangeConfig` 约束每个提供方操作：`maxFileBytes`（默认 25 MiB）、`maxCells`（2,000,000）、`maxConversionMs`（60,000 毫秒）、`commandTimeoutMs`（5,000 毫秒）、`maxSourceAnchors`（100,000）与 `maxSourceAnchorBytes`（8 MiB）。提供方在挂载时一次性解析这些上限，并以 `RangeError` 拒绝非正数或非整数值，而不是截断工作。超过 `maxConversionMs` 的旧格式转换会被中止并报告为 `too-large`，源文件保持不变。每次操作失败都是 `OfficeExchangeError`，其稳定 `code` 为 `unsupported-format`、`unsupported-feature`、`too-large`、`conversion-failed`、`not-implemented` 或 `out-of-scope`。

## officeDocument Remote

Remote 携带 `OfficeDocumentSnapshot`，即单元快照的 JSON 投影：`{ kind: 'workbook' | 'document', data: JsonValue }`。Univer 原生接口允许 JSON 无法承载的任意自定义字段，因此 Host 在序列化前校验该投影，并在保存时恢复原生快照；非 JSON 值以 `conversion-failed` 拒绝。

| 方法 | 行为 |
|---|---|
| `read(workspaceFileScope, path)` | 从 Session 工作区根解析 `path`，将其限制在该根内，通过文件系统 seam 在 `maxBytes` 限制下读取，并返回 `OfficeDocumentReadResult`（`format`、`snapshot`、`preserved?`、`unsupported`、`version`）；打开事实记录工作区相对定位符 |
| `save(workspaceFileScope, path, snapshot, format, expectedVersion, attribution)` | 在 `expectedVersion` 读取源，从这些字节写入，并通过文件系统的 `replaceIfVersion` 守卫发布；返回 `OfficeDocumentSaveResult`（`version`、可选 `disposition`），并在 `document/save` 上记录 `DocumentSaveAttribution`（`actor`、可选 `turn`）与处置 |
| `detectUnsupported(workspaceFileScope, path)` | 报告易丢失特性而不返回内容 |
| `convertLegacy(workspaceFileScope, path)` | 返回位于全新工作区作用域目录中的转换后可编辑文件路径 |

`read` 记录文件的不透明内容版本，并在解析后重新 stat；读取期间发生变化的文件会被拒绝，而不是返回绑定到过期版本的快照。`save` 绝不回退为无条件写入：当文件在读取后发生变化时，文件系统报告 `FS_STALE_VERSION`，Remote 抛出 `office-document/conflict`。与快照种类不匹配的 `format`，或非可编辑格式，同样会被拒绝，而不是写入部分文件。与耦合构造相交的写入以 `office-document/unsupported-feature` 拒绝并指出该构造，且不追加保存事实。

`SaveDisposition` 随保存结果与 `document/save` 事实传递。它携带 `schemaVersion: 1`、精确的 `verbatimCount`，以及 `spliced`、`added` 与 `removed` 条目各一个 `{ count, names }` 组。计数是精确的；名字是排序后的 ASCII 前缀，每组最多 256 个。该属性是可选的，因此处置存在之前记录的保存事实，或旧版重新生成，会读作缺席而不是无效。

Remote 将 seam 与文件系统失败映射为稳定的 Remote 代码：`unsupported-format`、`unsupported-feature`、`out-of-scope`、`bound-exceeded`、`conversion-failed` 和 `conflict`，每种都携带出错的 `path`。调用方信号被中止时以 `gateway/cancelled` 拒绝。`convertLegacy` 创建 `<workspaceRoot>/.dsh/office-tmp/convert-<hash>-<random>/` 并转换到其中，因此同一 Remote 可以在不变的根检查下读取结果。每次调用获得独立目录；转换失败会删除该目录，成功则保留到用户或 Session 删除为止。

## 挂载与组合

`LocalOfficeExchange` 校验其 `Config`，并以 `maxBytes` 等于 `maxFileBytes` 挂载 `OfficeDocumentRemote`。[base bundle](../../packages/bundle/base/README.zh.md) 在 `office-exchange` 行中挂载该提供方，配置为 `maxFileBytes: 26214400`、`maxCells: 2000000` 与 `maxConversionMs: 60000`。Client 组合通过 [api/remotes](../../packages/api/remotes/README.zh.md) 挂载来自 `@forwork/office-exchange/remote` 的生成描述符，并将其暴露为 `remote.officeDocument`。

<!-- BEGIN GENERATED cordis-surface (gen-cordis-catalog.ts) — do not edit between markers -->

<a id="cordis-surface"></a>

## Cordis API

Generated from source by `scripts/gen-cordis-catalog.ts` (verified fresh by `pnpm run verify-cordis-catalog` in doc-sync; regenerate with `pnpm run gen-cordis-catalog`) — the language sides differ only in locale-specific paired document paths. Signature blocks use a `ts cordis-catalog` fence and keep the original source JSDoc; dispatch modes are defined in the [primer](../cordis-primer.zh.md#dispatch-modes), and the framework-inherited `ctx` API lives in [cordis-api/inherited.md](../cordis-api/inherited.md).

<a id="ctxofficecommandbroker--officecommandbroker"></a>

### `ctx.officeCommandBroker` — `OfficeCommandBroker`

Owns pending document commands, the live target registrations, and the last observed content version per document. One instance serves every Session in its composition; a result is accepted only for the Session that recorded the request.

```ts cordis-catalog
/**
 * Append one `document/command` request and wait for the live editor's
 * answer.
 * A document with no registered live target is answered as
 * `document-not-open` without appending a command or starting the timeout.
 * @param request - session, document, operation, and expected version.
 * @param signal - caller cancellation; aborting withdraws the pending request.
 * @returns the editor's result, `document-not-open` when no target is registered, or
 * `timeout` when a registered target did not answer in time.
 * @throws the reason carried by an aborted `signal`, or an append failure.
 */
async request(request: DocumentCommandRequest, signal: AbortSignal): Promise<OfficeCommandOutcome>

/**
 * Resolve one pending request with the live editor's result.
 *
 * A result for an unknown request id, a request from another Session, or a
 * request that already timed out is ignored and reported as a diagnostic;
 * late answers never change durable state.
 * @param scope - Session identity that produced the result.
 * @param requestId - correlation id copied from the command event.
 * @param result - the live editor's answer.
 */
complete(scope: WorkspaceFileScope, requestId: string, result: OfficeCommandResult): void

/**
 * Register one live editor target that can answer commands for a document.
 *
 * A registration is identified by the caller's id so an attach after a detach
 * for the same document does not retire the newer target.
 * @param documentId - stable identity of the document.
 * @param registrationId - caller-minted identity of this attachment.
 */
attachTarget(documentId: DocumentId, registrationId: string): void

/**
 * Release one live editor target registration.
 * @param documentId - stable identity of the document.
 * @param registrationId - identity passed to {@link attachTarget}.
 */
detachTarget(documentId: DocumentId, registrationId: string): void

/**
 * Whether any live editor target is registered for one document.
 * @param documentId - stable identity of the document.
 * @returns `true` when at least one attachment can answer commands.
 */
hasTarget(documentId: DocumentId): boolean

/**
 * Read the last content version reported for one document.
 * @param documentId - stable identity of the document.
 * @returns the last reported version, or `0` before any command reported one.
 */
versionOf(documentId: DocumentId): number

/**
 * Record the content version one document currently holds.
 * @param documentId - stable identity of the document.
 * @param version - non-negative version reported by the live editor.
 */
remember(documentId: DocumentId, version: number): void
```

Types: [DocumentId](core.zh.md)

Source: [`packages/document/office-exchange/src/commands.ts`](../../packages/document/office-exchange/src/commands.ts)

<a id="ctxofficedocument--officedocumentremote"></a>

### `ctx.officeDocument` — `OfficeDocumentRemote`

Host Remote service backing the generated `remote.officeDocument` namespace. Reads return a JSON snapshot bound to the file's observed version; saves replace only that version, so a concurrent external edit is reported as a conflict rather than overwritten.

The service is the producer of `document/*` facts for the editor: a first read appends `document/open` once per Session, a successful save appends `document/save`, and the lifecycle endpoints append their matching facts. The open set is owned by this service per Session; no module-level state is used.

```ts cordis-catalog
/**
 * Read an editable Office file into a JSON snapshot bound to its version.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param path - absolute or workspace-relative path; it must resolve inside the workspace root.
 * The recorded document identity uses the same workspace-relative locator the editor tab uses,
 * so two spellings of one file never open two documents.
 * @param signal - caller cancellation.
 * @returns the detected format, snapshot, preserved construct classes,
 * unsupported features, and observed version.
 * @throws RemoteError when the path is out of scope, the format is not editable, or the read fails.
 */
@Remote async read( workspaceFileScope: WorkspaceFileScope, path: string, signal: AbortSignal, ): Promise<OfficeDocumentReadResult>

/**
 * Serialize a snapshot and replace the file only at the caller's version.
 *
 * The file is read and stat-verified at `expectedVersion` before the writer
 * runs, so the bytes the writer receives are the bytes the caller read. A
 * preserved construct never refuses the save by itself: the writer returns a
 * typed refusal only when the edit intersects a construct whose meaning or
 * anchors depend on the changed content. A refusal writes nothing and records
 * no fact; the final `replaceIfVersion` write refuses a change that happened
 * during serialization.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param path - absolute or workspace-relative path; it must resolve inside the workspace root.
 * @param snapshot - snapshot returned by {@link read}, optionally modified.
 * @param format - target editable format; it must match the snapshot kind.
 * @param expectedVersion - version returned by {@link read}; a changed file is refused.
 * @param attribution - who caused the save and the agent turn, when one caused it.
 * @param signal - caller cancellation.
 * @returns the version the guarded write published and the writer's disposition.
 * @throws RemoteError with `office-document/unsupported-feature` when the edit
 * intersects a coupled construct, naming that construct in the message.
 * @throws RemoteError with `office-document/conflict` when the file changed after the read.
 */
@Remote async save( workspaceFileScope: WorkspaceFileScope, path: string, snapshot: OfficeDocumentSnapshot, format: OfficeFormat, expectedVersion: string, attribution: DocumentSaveAttribution, signal: AbortSignal, ): Promise<OfficeDocumentSaveResult>

/**
 * Record that an open document left the editor.
 *
 * The fact is appended even when this process no longer tracks the document,
 * so a tab closed after a Host restart still leaves a durable close. History
 * is never deleted; the open set is derived by folding the log.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param documentId - Session-scoped identity of the closing document.
 * @param reason - producer-supplied closure reason; the projection stores it unread.
 * @param signal - caller cancellation.
 * @throws RemoteError when the identity belongs to another Session or no live Session exists.
 */
@Remote closeDocument( workspaceFileScope: WorkspaceFileScope, documentId: string, reason: string, signal: AbortSignal, ): Promise<void>

/**
 * Record that an open document's file moved or was renamed.
 *
 * The identity is preserved: only the locator changes, and the new path must
 * resolve inside the Session workspace root.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param documentId - Session-scoped identity retained across the move.
 * @param toPath - new path; it must resolve inside the workspace root.
 * @param signal - caller cancellation.
 * @throws RemoteError when the identity belongs to another Session, either path is out of scope, or no live Session exists.
 */
@Remote async rebindDocument( workspaceFileScope: WorkspaceFileScope, documentId: string, toPath: string, signal: AbortSignal, ): Promise<void>

/**
 * Record that an open document's file no longer resolves.
 *
 * The document stays open with its identity and history; a later save, edit,
 * or rebind marks it available again.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param documentId - Session-scoped identity that stays open.
 * @param signal - caller cancellation.
 * @throws RemoteError when the identity belongs to another Session or no live Session exists.
 */
@Remote markDocumentMissing( workspaceFileScope: WorkspaceFileScope, documentId: string, signal: AbortSignal, ): Promise<void>

/**
 * Report the features a whole-file regeneration would drop without returning content.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param path - absolute or workspace-relative path; it must resolve inside the workspace root.
 * @param signal - caller cancellation.
 * @returns every detected unsupported feature; empty when regeneration loses no mapped content.
 * @throws RemoteError when the path is out of scope or its format cannot be inspected.
 */
@Remote async detectUnsupported( workspaceFileScope: WorkspaceFileScope, path: string, signal: AbortSignal, ): Promise<readonly UnsupportedFeature[]>

/**
 * Convert a legacy binary Office file to an editable one.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param path - absolute or workspace-relative legacy path; it must resolve inside the workspace root.
 * @param signal - caller cancellation.
 * @returns the path of the converted editable file inside the workspace root.
 * @throws RemoteError when the path is out of scope or conversion fails.
 */
@Remote async convertLegacy(workspaceFileScope: WorkspaceFileScope, path: string, signal: AbortSignal): Promise<string>

/**
 * Persist one recovery draft for an open document.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param documentId - caller-owned identity of the open document; it selects the draft file.
 * @param documentPath - document path the draft is captured for; it is recorded, never written.
 * @param contentVersion - file version the snapshot was captured from.
 * @param snapshot - wire snapshot of the editable content.
 * @param signal - caller cancellation.
 * @returns the time the draft was written.
 * @throws RemoteError when the draft store is out of scope or cannot be written.
 */
@Remote async writeDraft( workspaceFileScope: WorkspaceFileScope, documentId: string, documentPath: string, contentVersion: string, snapshot: OfficeDocumentSnapshot, signal: AbortSignal, ): Promise<OfficeDocumentDraftWriteResult>

/**
 * Read the recovery draft persisted for one open document.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param documentId - caller-owned identity of the open document; it selects the draft file.
 * @param signal - caller cancellation.
 * @returns the stored draft, or `null` when the document has no draft.
 * @throws RemoteError when the draft store is out of scope or its content is malformed.
 */
@Remote async readDraft( workspaceFileScope: WorkspaceFileScope, documentId: string, signal: AbortSignal, ): Promise<OfficeDocumentDraft | null>

/**
 * Remove the recovery draft persisted for one open document.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param documentId - caller-owned identity of the open document; it selects the draft file.
 * @param signal - caller cancellation.
 * @throws RemoteError when the draft store is out of scope or its entry cannot be removed.
 */
@Remote async clearDraft(workspaceFileScope: WorkspaceFileScope, documentId: string, signal: AbortSignal): Promise<void>

/**
 * Answer one pending document command with the live editor's result.
 *
 * A result for an unknown request id is ignored and logged by the broker;
 * the call itself succeeds so a late or duplicate answer never becomes a
 * transport error.
 * @param workspaceFileScope - Session identity that ran the command on the wire.
 * @param requestId - correlation id copied from the `document/command` event.
 * @param result - the live editor's status and, for read or applied commands, content.
 * @param signal - caller cancellation; an aborted call drops the answer.
 * @throws RemoteError with `gateway/bad-request` when a reported version is not a non-negative safe integer.
 */
@Remote completeCommand( workspaceFileScope: WorkspaceFileScope, requestId: string, result: OfficeCommandResult, signal: AbortSignal, ): Promise<void>

/**
 * Register one live editor target for a document.
 *
 * The renderer calls this when its command bridge attaches and releases the
 * matching registration when that bridge detaches. A command dispatched while
 * no registration exists is answered as `document-not-open` instead of
 * waiting for the timeout.
 * @param workspaceFileScope - Session identity that owns the document.
 * @param documentId - Session-scoped identity of the document.
 * @param registrationId - caller-minted identity of this attachment.
 * @param signal - caller cancellation; an aborted call registers nothing.
 * @throws RemoteError when the identity belongs to another Session or no live Session exists.
 */
@Remote registerCommandTarget( workspaceFileScope: WorkspaceFileScope, documentId: string, registrationId: string, signal: AbortSignal, ): Promise<void>

/**
 * Release one live editor target registration.
 * @param workspaceFileScope - Session identity that owns the document.
 * @param documentId - Session-scoped identity of the document.
 * @param registrationId - identity passed to {@link registerCommandTarget}.
 * @param signal - caller cancellation; an aborted call releases nothing.
 * @throws RemoteError when the identity belongs to another Session or no live Session exists.
 */
@Remote unregisterCommandTarget( workspaceFileScope: WorkspaceFileScope, documentId: string, registrationId: string, signal: AbortSignal, ): Promise<void>

/**
 * Dispatch one correlated command from a Host document tool.
 *
 * The broker appends `document/command` and waits for the live editor; this
 * method is the Host-only entry point behind the tool's command executor and
 * is not a Remote endpoint.
 * @param request - session, document, operation, and expected version.
 * @param signal - caller cancellation; aborting withdraws the pending request.
 * @returns the editor's result, `document-not-open` without a registered target, or
 * `timeout` when a registered target did not answer in time.
 * @throws RemoteError with `gateway/internal` when the broker is not mounted.
 */
async dispatchCommand(request: DocumentCommandRequest, signal: AbortSignal): Promise<OfficeCommandOutcome>
```

Source: [`packages/document/office-exchange/src/remote.ts`](../../packages/document/office-exchange/src/remote.ts)

<a id="ctxofficeexchange--officeexchange-abstract-seam"></a>

### `ctx.officeExchange` — `OfficeExchange` (abstract seam)

Service Definition for Office file interchange. A provider implements these operations for the deployment's conversion libraries; consumers depend on the seam only.

```ts cordis-catalog
/**
 * Read an Office file into an editable unit snapshot.
 * @param path - Absolute or workspace-relative file path.
 * @param signal - Optional cancellation.
 * @returns The detected format, snapshot, and unsupported features.
 */
abstract read(path: string, signal?: AbortSignal): Promise<ExchangeReadResult>

/**
 * Serialize a unit snapshot back to Office file bytes.
 *
 * `source` carries the bytes the caller read at its expected version so an
 * in-place writer can splice them. A snapshot without the source sidecar, or
 * a caller that supplies no source, uses the regeneration writer.
 * @param snapshot - Editable content to serialize.
 * @param format - Target Office format.
 * @param source - Source package bytes read at the caller's expected version;
 * omitted for a snapshot that has no source file.
 * @param signal - Optional cancellation.
 * @returns The encoded file bytes and what the save carried or changed.
 */
abstract write( snapshot: UnitSnapshot, format: OfficeFormat, source?: SourcePackage, signal?: AbortSignal, ): Promise<OfficeWriteResult>

/**
 * Report features a whole-file regeneration would drop.
 * @param path - File path to inspect.
 * @param signal - Optional cancellation.
 * @returns Detected unsupported features; empty when regeneration loses no mapped content.
 */
abstract detectUnsupported(path: string, signal?: AbortSignal): Promise<readonly UnsupportedFeature[]>

/**
 * Convert a legacy binary Office file to an editable format.
 * @param path - Legacy file path.
 * @param signal - Optional cancellation.
 * @param outputDir - Directory receiving the converted file; omitted uses a
 * fresh temporary directory outside any workspace.
 * @returns Path of the converted editable file.
 */
abstract convertLegacy(path: string, signal?: AbortSignal, outputDir?: string): Promise<string>
```

Source: [`packages/document/office-exchange/src/service.ts`](../../packages/document/office-exchange/src/service.ts)
<!-- END GENERATED cordis-surface -->
