# Office 交换

[English](office-exchange.md) | 中文

[document 包族](../../packages/document/README.zh.md) 在磁盘字节与 Office 编辑器及智能体文档工具使用的可编辑单元快照之间移动 Office 文件。[office-exchange](../../packages/document/office-exchange) 声明 `ctx.officeExchange` 能力 seam，提供 `LocalOfficeExchange` 提供方，并承载会话级的 `officeDocument` Remote。该包不注册面向模型的工具，也不追加 Session 事件。预留的 [Office 编辑器 Client](../../packages/client/ui-office-editor) 与 [Office 文档工具](../../packages/document/tool-office) 目前不注册标签页、面板、slot 或工具；落地后，编辑器通过 `remote.officeDocument` 读写，面向模型的工具使用 `ctx.officeExchange`。

## 所有权

| 所有者 | 职责 |
|---|---|
| [office-exchange](../../packages/document/office-exchange) | `ctx.officeExchange`：读取、写入、不支持特性检测与旧格式转换；`officeDocument` Remote |
| [api/remotes](../../packages/api/remotes/README.zh.md) | 挂载生成的 `remote.officeDocument` Client 命名空间 |
| [base bundle](../../packages/bundle/base/README.zh.md) | 以部署上限挂载 `office-exchange` |

## 格式与单元快照

`OfficeFormat` 为 `docx`、`xlsx`、`pptx`、`doc`、`xls` 或 `ppt`。可编辑集合是 `EditableOfficeFormat`（`docx`、`xlsx`）；旧格式集合是 `LegacyOfficeFormat`（`doc`、`xls`、`ppt`）。

| 格式 | 操作 |
|---|---|
| `.docx`、`.xlsx` | 可编辑：`read` 将文件字节映射为 `UnitSnapshot`；`write` 将快照序列化回字节 |
| `.doc`、`.xls`、`.ppt` | 旧格式：`convertLegacy` 生成可编辑文件；对原文件调用 `read` 以 `unsupported-format` 拒绝 |
| `.pptx` | 仅预览：`detectUnsupported` 可检查；`read` 与 `write` 以 `unsupported-format` 拒绝 |

`UnitSnapshot` 是编辑器原生的可序列化快照，是以 `{ kind: 'workbook', data: IWorkbookData }` 与 `{ kind: 'document', data: IDocumentData }` 构成的可辨识联合。`read` 返回 `ExchangeReadResult`，包含检测到的 `format`、解析出的 `snapshot`，以及一次编辑并保存会丢失的全部 `UnsupportedFeature`。每个 `UnsupportedFeature` 携带 `kind`——`macro`、`tracked-changes`、`comment`、`chart`、`drawing`、`shape`、`header-footer`、`formula-recalc`、`embedded-object` 或 `other`——以及可选的、面向人类的 `detail`。非空列表意味着写入路径必须拒绝，而不是保存有损文件。

`formatOf(path)` 按扩展名分类路径，`isEditableFormat` 收窄到 `docx`/`xlsx`，`isLegacyFormat` 收窄到 `doc`/`xls`/`ppt`。

## Seam 操作与失败

`ctx.officeExchange.read(path, signal?)` 返回 `ExchangeReadResult`。`ctx.officeExchange.write(snapshot, format, signal?)` 返回编码后的字节，但不写文件。`detectUnsupported(path, signal?)` 报告易丢失特性而不返回内容。`convertLegacy(path, signal?, outputDir?)` 返回可编辑文件的路径；省略 `outputDir` 时写入任意工作区之外的临时目录。

`OfficeExchangeConfig` 约束每个提供方操作：`maxFileBytes`（默认 25 MiB）、`maxCells`（2,000,000）和 `maxConversionMs`（60,000 毫秒）。提供方在挂载时一次性解析这些上限，并以 `RangeError` 拒绝非正数或非整数值，而不是截断工作。每次操作失败都是 `OfficeExchangeError`，其稳定 `code` 为 `unsupported-format`、`unsupported-feature`、`too-large`、`conversion-failed`、`not-implemented` 或 `out-of-scope`。

## officeDocument Remote

Remote 携带 `OfficeDocumentSnapshot`，即单元快照的 JSON 投影：`{ kind: 'workbook' | 'document', data: JsonValue }`。Univer 原生接口允许 JSON 无法承载的任意自定义字段，因此 Host 在序列化前校验该投影，并在保存时恢复原生快照；非 JSON 值以 `conversion-failed` 拒绝。

| 方法 | 行为 |
|---|---|
| `read(workspaceFileScope, path)` | 从 Session 工作区根解析 `path`，将其限制在该根内，通过文件系统 seam 在 `maxBytes` 限制下读取，并返回 `OfficeDocumentReadResult`（`format`、`snapshot`、`unsupported`、`version`） |
| `save(workspaceFileScope, path, snapshot, format, expectedVersion)` | 序列化快照并通过文件系统的 `replaceIfVersion` 守卫发布；返回 `OfficeDocumentSaveResult`（`version`） |
| `detectUnsupported(workspaceFileScope, path)` | 报告易丢失特性而不返回内容 |
| `convertLegacy(workspaceFileScope, path)` | 返回位于全新工作区作用域目录中的转换后可编辑文件路径 |

`read` 记录文件的不透明内容版本，并在解析后重新 stat；读取期间发生变化的文件会被拒绝，而不是返回绑定到过期版本的快照。`save` 绝不回退为无条件写入：当文件在读取后发生变化时，文件系统报告 `FS_STALE_VERSION`，Remote 抛出 `office-document/conflict`。与快照种类不匹配的 `format`，或非可编辑格式，同样会被拒绝，而不是写入部分文件。

Remote 将 seam 与文件系统失败映射为稳定的 Remote 代码：`unsupported-format`、`unsupported-feature`、`out-of-scope`、`bound-exceeded`、`conversion-failed` 和 `conflict`，每种都携带出错的 `path`。调用方信号被中止时以 `gateway/cancelled` 拒绝。`convertLegacy` 创建 `<workspaceRoot>/.dsh/office-tmp/convert-<hash>-<random>/` 并转换到其中，因此同一 Remote 可以在不变的根检查下读取结果。每次调用获得独立目录；转换失败会删除该目录，成功则保留到用户或 Session 删除为止。

## 挂载与组合

`LocalOfficeExchange` 校验其 `Config`，并以 `maxBytes` 等于 `maxFileBytes` 挂载 `OfficeDocumentRemote`。[base bundle](../../packages/bundle/base/README.zh.md) 在 `office-exchange` 行中挂载该提供方，配置为 `maxFileBytes: 26214400`、`maxCells: 2000000` 与 `maxConversionMs: 60000`。Client 组合通过 [api/remotes](../../packages/api/remotes/README.zh.md) 挂载来自 `@deepseek-ai/dsh-office-exchange/remote` 的生成描述符，并将其暴露为 `remote.officeDocument`。

<!-- BEGIN GENERATED cordis-surface (gen-cordis-catalog.ts) — do not edit between markers -->

<a id="cordis-surface"></a>

## Cordis API

Generated from source by `scripts/gen-cordis-catalog.ts` (verified fresh by `pnpm run verify-cordis-catalog` in doc-sync; regenerate with `pnpm run gen-cordis-catalog`) — the language sides differ only in locale-specific paired document paths. Signature blocks use a `ts cordis-catalog` fence and keep the original source JSDoc; dispatch modes are defined in the [primer](../cordis-primer.zh.md#dispatch-modes), and the framework-inherited `ctx` API lives in [cordis-api/inherited.md](../cordis-api/inherited.md).

<a id="ctxofficedocument--officedocumentremote"></a>

### `ctx.officeDocument` — `OfficeDocumentRemote`

Host Remote service backing the generated `remote.officeDocument` namespace. Reads return a JSON snapshot bound to the file's observed version; saves replace only that version, so a concurrent external edit is reported as a conflict rather than overwritten.

```ts cordis-catalog
/**
 * Read an editable Office file into a JSON snapshot bound to its version.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param path - absolute or workspace-relative path; it must resolve inside the workspace root.
 * @param signal - caller cancellation.
 * @returns the detected format, snapshot, unsupported features, and observed version.
 * @throws RemoteError when the path is out of scope, the format is not editable, or the read fails.
 */
@Remote async read( workspaceFileScope: WorkspaceFileScope, path: string, signal: AbortSignal, ): Promise<OfficeDocumentReadResult>

/**
 * Serialize a snapshot and replace the file only at the caller's version.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param path - absolute or workspace-relative path; it must resolve inside the workspace root.
 * @param snapshot - snapshot returned by {@link read}, optionally modified.
 * @param format - target editable format; it must match the snapshot kind.
 * @param expectedVersion - version returned by {@link read}; a changed file is refused.
 * @param signal - caller cancellation.
 * @returns the version the guarded write published.
 * @throws RemoteError with `office-document/conflict` when the file changed after the read.
 */
@Remote async save( workspaceFileScope: WorkspaceFileScope, path: string, snapshot: OfficeDocumentSnapshot, format: OfficeFormat, expectedVersion: string, signal: AbortSignal, ): Promise<OfficeDocumentSaveResult>

/**
 * Report the features an edit-and-save would drop without returning content.
 * @param workspaceFileScope - header-derived workspace root for the Session identity on the wire.
 * @param path - absolute or workspace-relative path; it must resolve inside the workspace root.
 * @param signal - caller cancellation.
 * @returns every detected unsupported feature, empty when the file round-trips cleanly.
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
 * @param snapshot - Editable content to serialize.
 * @param format - Target Office format.
 * @param signal - Optional cancellation.
 * @returns The encoded file bytes.
 */
abstract write(snapshot: UnitSnapshot, format: OfficeFormat, signal?: AbortSignal): Promise<Uint8Array>

/**
 * Report features that would be lost by an edit-and-save.
 * @param path - File path to inspect.
 * @param signal - Optional cancellation.
 * @returns Detected unsupported features; empty when the file round-trips cleanly.
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
