# Office 文档智能体框架 — Fork + Univer 内核 实现计划

## 目标

把 ForWork **fork 成自有产品**（改 scope/品牌/签名，不蹭 deepseek），并在其核心内置"智能体操作办公文档"能力：文档成为**会话一等实体**，侧栏内嵌**可编辑的 Word/Excel**（Univer 内核），Agent 通过**共享内容模型**高保真读写同一文档。

## 已确认决策

| 决策点 | 选择 |
|---|---|
| 交付模型 | Fork deepseek-harness → 独立仓库，改 scope/品牌/签名，可改核心逻辑 |
| 编辑器内核 | Univer Office SDK（Apache-2.0 核心：Sheets/Docs 编辑 + Facade + Headless），React 直嵌、无 iframe、无 Linux 服务端 |
| 转换层 | 自建"Exchange"：xlsx↔Univer snapshot 用 ExcelJS/SheetJS，docx↔snapshot 用 docx 库，旧格式转换用 LibreOffice-kit |
| 文件类型 | `.docx`/`.xlsx` 可编辑（首期）；`.pptx` 首期只读预览（Univer Slides 是商业，deferred） |
| 集成深度 | 文档为会话一等实体：新增 `document/*` 会话事件 + 文档实体投影 + 跨 turn 上下文 + GUI 常驻面板 |
| Agent 协同 | 内容模型收敛：Agent 用 headless Univer（Facade/命令）与编辑器共享同一 Unit snapshot；应用自有持久化经 Exchange 写回文件 |
| 平台 | Windows 桌面（Electron，独立代码签名） |
| 密钥 | 随引擎切换消解（无 WOPI/JWT）；剩余密钥 = 自有模型 API Key / 签名证书（TBD 输入） |
| 许可 | harness MIT + Univer Apache-2.0 + ExcelJS MIT + SheetJS Apache-2.0，均可商用/独立签名 |

## 交付模型与数据流

### 内容模型收敛（取代之前的"文件收敛"）

```
用户编辑(Univer 内嵌)   ──┐
Agent 编辑(headless Univer 命令) ──┤→ 同一 Unit Snapshot
                                │   (IWorkbookData / IDocumentData，可序列化 JSON)
                                └→ 应用持久化: Exchange 序列化 → .xlsx/.docx
                                       → ctx.fs 版本守卫写 → fs/observed
```

- Agent 与编辑器**共享同一个内容模型**（不再是"各写各的文件"），Undo/命令/公式计算一致。
- 持久化由应用负责（Univer 开放源码不含 Exchange，应用自己 serialize→文件）；写路径走 `ctx.fs` + `fs/observed`，天然驱动 `workspace-files.changes` 观察链（预览/面板自动刷新）。
- 打开流程：文件 → Exchange 解析 → Unit snapshot → 编辑器挂载；保存流程：snapshot → Exchange 序列化 → `ctx.fs` 版本守卫写。

### 会话一等实体

- 新增 `document/*` 会话事件族（`document/open`、`document/edit`、`document/close`、`document/state`），文档打开状态进入 session 日志，可跨 turn 引用（"第三个表格的 B 列"）。
- 文档实体投影（同 workspace/文件投影的机制）：`ctx.documentProjections` 折叠文档事件，Agent 系统提示注入"当前打开的文档 + 结构摘要"，GUI 面板从投影渲染。
- 这是本次"可改核心"的落点：`SessionEventMap` 增补 + 投影 + 系统提示节 + 迁移。

## Fork / 改品牌边界

**保留（借用的框架）**：Cordis DI、agent-loop、session 日志、fs/sandbox、shell/subprocess、GUI（slots/dockkit/sidebar/files/preview）、typert、构建链。

**核心改动**：新增 `document/*` 事件与投影（见上）；office 能力**烤进 base 组合，非可逆、非市场插件**（内部仍是 Cordis 服务——框架现实，但对外不暴露为可卸载行）。

**剥离 / 重命名**（不"蹭 deepseek"的完整性清单）：
- `@deepseek-ai/dsh-*` → `<org>/<pkg>*`（全部包 scope + 依赖引用）
- 品牌：fish 标记/名称槽位 → 自有品牌（`sidebar.brand.*`、桌面菜单、窗口标题）
- `forwork` CLI 名 → 自有命令名
- 遥测：`product-analytics` → 移除或替换为自有（默认关闭）
- 模型默认适配器：`web-api` / `FORWORK_API_KEY` → 自有模型配置
- 文档/artwork/官网 → 自有；桌面 Windows/macOS 签名身份 → 自有证书
- 保留根 MIT 版权声明；Univer 等第三方保留其 NOTICE

## 新组件与核心改动

1. **核心会话模型**：`SessionEventMap` 增补 `document/*` 事件（需 `SESSION_FORMAT_VERSION` 决策与迁移）；`ctx.documentProjections` 投影 + 系统提示注入"打开文档结构"。
2. **Exchange 层**（新 `packages/document/office-exchange`）：xlsx↔`IWorkbookData`、docx↔`IDocumentData` 双向映射 + 旧格式（`.doc/.xls/.ppt`）→ 新格式转换。xlsx 用已锁定 ExcelJS/SheetJS；docx 用 run 级 OOXML 映射（保真边界见风险）。
3. **Univer 编辑器 tab**（新 `packages/client/ui-office-editor`）：复用 sidebar-right 体系，`ctx.sidebarRightTabs.register({ kind:'office', patterns:['*.docx','*.xlsx'] })`，React 直嵌 Univer Sheets/Docs preset；订阅 `workspace-files.changes` 做外部变更重载。
4. **Agent 文档工具**（新 `packages/document/tool-office`）：注册 `ctx.tools` 结构化文档读写工具（读结构、改单元格/行列/公式、改段落），底层走 headless Univer Facade/命令 + Exchange 写回；写经 `ctx.fs`。
5. **常驻文档面板**：GUI 常驻"打开的文档"列表（一等实体），从投影渲染。

## 实现任务（有序，每阶段可独立验证）

### 阶段 0 — Fork 与改品牌（无功能变化，先打通可发布底座）
1. Fork 仓库到自有组织；`@deepseek-ai/dsh-*` → `<org>/*`（含 `package.json`、tsconfig references、import、cordis.yml 行、bundle 依赖）。
2. 改品牌槽位、`forwork` CLI 名、移除/替换 `product-analytics`、替换默认模型适配器与 `FORWORK_API_KEY` 引用、删 deepseek 专属文档/artwork。
3. 保留 MIT 声明；验证：`pnpm install && pnpm run build && pnpm run typecheck && pnpm run lint` 全绿，桌面可 `package`。

### 阶段 1 — Exchange 转换层（无 UI，先 xlsx 后 docx）
4. 新建 `office-exchange`：xlsx↔`IWorkbookData`（ExcelJS/SheetJS）+ 往返单测（字节级保真）；docx↔`IDocumentData`（docx 库）+ 往返单测。
5. 旧格式转换：核实 `libreoffice-kit` 往返目标格式（非 PDF），不足则扩展独立 kit。
6. 验证：单测覆盖"未触及部分字节级保留"；`test:coverage` 每文件 100%。

### 阶段 2 — 核心文档实体（改核心）
7. `SessionEventMap` 增补 `document/*` 事件 + `SESSION_FORMAT_VERSION` 决策与相邻迁移。
8. `ctx.documentProjections` 投影 + 系统提示注入"打开文档结构摘要" + 模型可见即日志。
9. 验证：会话迁移测试 + 快照（`test:snapshot`）+ SDK 投影输出同步更新。

### 阶段 3 — Univer 编辑器 tab（GUI）
10. 新建 `ui-office-editor`：注册 `office` tab、React 直嵌 Univer（Sheets/Docs preset）、打开经 Exchange 解析、保存经 Exchange 写回、外部变更重载。
11. 常驻文档面板（一等实体列表）。
12. 验证：`pnpm run test:gui`；可见装配输出变更则 `DSH_SNAPSHOT=replay pnpm run test:web`。

### 阶段 4 — Agent 文档工具
13. 新建 `tool-office`：结构化文档读写工具（headless Univer + Exchange），model-visible schema 与结果。
14. 验证：快照覆盖工具输出；写路径走 `ctx.fs` 版本守卫。

### 阶段 5 — 治理与发布
15. 包 README（Model Experience 段）、子系统文档、config-catalog、升级指南（session 格式破坏面）。
16. 桌面独立签名（Windows/macOS 证书）+ 打包验证。

## 风险与失败模式

- **docx↔snapshot 保真度（最难）**：文本框/页眉页脚/批注/修订/公式可能不覆盖 → 明确"支持操作集"，超出走 LibreOffice 或明确报错，绝不静默丢内容。
- **LibreOffice 往返**：现有 `office-to-pdf` 只到 PDF；往返 docx/xlsx/pptx 是新能力，须在阶段 1 核实，不足则扩展独立 kit 或改用其他转换。
- **Univer 商业边界**：Slides/文件转换/协作/快公式引擎是商业 SDK——已用"自建 Exchange + pptx 首期只读"规避，但 Exchange 的 docx 保真是自己承担的风险。
- **session 格式迁移**：`document/*` 入 `SessionEventMap` 需 `SESSION_FORMAT_VERSION` 决策，破坏面走升级指南。
- **fork 完整性**：漏改任何 `forwork` 引用/遥测/品牌残留都会"仍蹭 deepseek"，需阶段 0 的引用扫描门禁（`verify-*` 脚本改造成自有断言）。

## 回滚 / 迁移（fork 顺序）

- 阶段 0 是纯 fork+改品牌，不改变运行时行为，可独立发布自有版本基线。
- 阶段 2 的 session 格式迁移遵循"仅追加相邻版本，不覆盖已发布代"；`SESSION_FORMAT_VERSION` 单调递增。
- 阶段 3/4 的 office 能力在自有产品里常驻（非可逆）；回退以"禁用 feature 开关"实现（若需）。

## 验证总览

- 每阶段：`test:coverage`（每文件 100%）+ 相关快照 + `typecheck` + `lint`。
- GUI 变更：`test:gui` + 浏览器回放快照 `test:web`。
- 发布：`build:check`、桌面 `package` + 签名验证、`doc-sync`。
- deepseek 专属 gate（`verify-*`）在阶段 0 改造成自有断言，避免对自有 scope 误报。

## 待定输入（TBD，填空即用，不阻塞）

- 自有产品名、npm org scope（贯穿阶段 0）。
- 自有代码签名身份（Windows 证书 / Apple Developer ID，阶段 5）。
- 自有模型 API 配置（阶段 0 替换默认适配器时填）。

## 延期（deferred / out of scope）

- `.pptx` 编辑（Univer Slides 商业；首期只读预览）。
- 实时多用户协作（Agent 作为 OT 第二主体）。
- 旧格式（`.doc/.xls/.ppt`）编辑后原地回写二进制格式。
- 独立代码/文本编辑器（用户确认不需要）。
- 自动托管任何服务端（Univer 内嵌后无此需求）。
