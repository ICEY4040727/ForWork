# Agent Note: ForWork 身份例外：协议令牌与外部 Office kit

Status: implemented

[English](2026-11-17-forwork-identity-exceptions.md) | 中文

## 问题

[verify-forwork-identity](../../../../scripts/verify-forwork-identity.ts) 拒绝出现在已发布面（shipped surfaces）中的 DeepSeek 身份令牌：`@deepseek-ai` 包作用域、DeepSeek Harness 品牌、遗留的 `DEEPSEEK_*` 环境变量名、遗留的 `dsh` 命令名，以及遗留的产物前缀。有两组名称无法靠字符串替换改成 ForWork 名称，而对它们不加说明地放行，与漏掉违规无法区分。

第一组是已发布的 profile、已发布的第三方插件清单和现有用户 home 已经在使用的协议词汇：`dsh.bundle`、`dsh.client`、`dsh.profile` 三个包清单键，包含 `$DSH_HOME` 的 `DSH_*` 环境变量族，以及 `dsh-app://` Electron 协议。

第二组是 `@deepseek-ai/libreoffice-kit` 及其 `-darwin-*`、`-win32-*`、`-wasm` 引擎包，即遗留 `.doc`、`.xls`、`.ppt` 转换背后的外部引擎。

## 决策

两组名称都保持现状，每项例外都写明名称搬走之前必须先改变什么。协议例外是不新增一条指名这些令牌的规则；Office 例外是现有 allowlist 中的一条记录。两者都不放宽其余规则。

### 协议令牌

身份扫描没有任何规则覆盖这些令牌，因此它们在构造上就被放行。每一项都是已发布的协议事实：

- `dsh.bundle` 和 `dsh.profile` 是包清单键。profile 组合器按 `dsh.profile.bundles` 的顺序解析 patch 层（[profile-boot.ts](../../../../apps/cli/src/profile-boot.ts)），插件管理器读取注册表发布元数据的 `dsh.bundle` 来判断某个已安装版本是否为 bundle，并通过 `dsh.bundle.patch` 定位其 patch 文件（[operations.ts](../../../../packages/boot/plugin-manager/src/operations.ts)）。
- `dsh.client` 为客户端 bundle 声明包的浏览器入口、外部运行时说明符和平台（[verify-client-packages.ts](../../../../scripts/verify-client-packages.ts)）。
- `$DSH_HOME` 及更广的 `DSH_*` 环境变量族指定桌面应用、Python SDK、e2e harness 和用户配置所使用的 harness home。`$DSH_HOME` 默认为 `~/.dsh`，现有会话、设置、凭据、profile 和附件都位于该路径下。
- `dsh-app://` 是打包后的桌面构建为应用文档及其本地资源注册的 Electron 自定义协议（[web-document.ts](../../../../apps/desktop/src/web-document.ts)）。

改这些清单键会同时牵动 profile 组合器、插件管理器、客户端构建，以及每一个声明它们的已发布第三方插件；只改 Harness 一侧，会让这些插件被当作普通依赖安装，而不是 profile 层。没有就地进行迁移的改名会让这些 home 不再被读取，升级后的安装从空状态开始，旧数据仍留在磁盘上。

这些令牌是被发布的协议词汇，不构成身份违规。只有在同时存在定义后继名称的 `SESSION_FORMAT_VERSION` 或持久化修订，以及针对现有 home 的就地迁移时，才值得迁移它们；本版本两者都不具备。

### 外部 Office kit

`@deepseek-ai/libreoffice-kit` 及其引擎包是保持外部依赖的第三方身份：

- ForWork 既不发布也不 vendor 该引擎，`THIRD_PARTY_NOTICES.md` 以 MPL-2.0 精确列出这些包身份，并附指向 kit 仓库的对应源码指针。
- [独立的 kit 归属](../architecture/2026-09-14-independent-libreoffice-kit.zh.md)决定把引擎资格认证与发布留在该仓库。
- 由 ForWork 自行构建引擎，意味着发布数 GB 的原生产物、为每个平台做资格认证，并接手 MPL-2.0 义务以及打包产物 notice 检查所核验的来源记录。

身份扫描的 `ALLOWED_DEEPSEEK_PACKAGES` 精确包含 `@deepseek-ai/libreoffice-kit`。allowlist 匹配该字面量及其 `-<suffix>` 引擎包，`@deepseek-ai` 作用域下的其他包一律不匹配。它保持为精确的包 allowlist；放行整个作用域会让任何上游包保留原名。

## 考虑过的替代方案

**在本版本内借助会话格式修订迁移协议令牌。** 本版本没有这样的修订，也没有 home 迁移，改名会分裂存储、破坏已安装的插件，而协议并不要求改名，收益只是命名。

**用 ForWork 发布或 vendor 的 kit 替换 Office 引擎。** 自建意味着为每个平台发布并通过资格认证，在本仓库承担其源码义务，同时放弃 notice 检查所核验的上游来源。

**在身份扫描中放行整个 `@deepseek-ai` 作用域。** 作用域级放行会不再报告任何新增上游依赖，扫描也就无法再区分经过评审的第三方引擎与泄漏的上游包。

## 后果

已发布的 home、已发布的插件和打包后的桌面构建继续工作，固定版本号的 kit 保持转换路径及其 notice 完整。

代价是：仓库在呈现 ForWork 身份的同时，保留带 `dsh` 前缀的协议名称和一个 `@deepseek-ai` 包。扫描的输出本身无法表明这些剩余名称是有意保留的，因此 allowlist 注释指向本 Note，作为评审者查阅的记录。

`verify-forwork-identity` 对任何其他 DeepSeek 身份令牌仍然失败，每新增一个外部 DeepSeek 包都需要修改 allowlist 并更新本 Note。协议令牌不在任何规则之下，将来若有规则覆盖它们，必须引用本 Note 或取代它。

只有在出现 `SESSION_FORMAT_VERSION` 或持久化修订并完成 home 迁移时，才重新审视协议例外；只有在 ForWork 发布自己的 kit 或替换转换引擎时，才重新审视 Office 例外。
