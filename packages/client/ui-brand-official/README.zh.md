---
description: "official 客户端构建的 ForWork 官方侧栏品牌。"
kind: "package-reference"
---

# @forwork/client-ui-brand-official

[English](README.md) | 中文

## 概述

本包为 `official` 客户端构建填充官方侧栏品牌 slot。随产品提供的 Web 与 Desktop 组合默认加载本包；其他构建 profile 保留原有品牌。空白首页的标记及悬停门扇来自 `dsh-client-ui-conversation` 回退，而不是本包。插件不增加模型请求，也不持有临时背景资源。

## 目录

- [使用本包](#use-this-package)
- [理解实现](#understand-the-implementation)
- [进一步探索](#further-exploration)
- [模型体验](#model-experience)
- [已知限制与延期工作](#known-limitations-and-deferred-work)
- [开发备注](#dev-note)

-----

<a id="use-this-package"></a>
## 使用本包

随产品提供的 Web 与 Desktop 组合已经挂载本插件。`official` 客户端构建会注册 ForWork 侧栏图形；不需要额外安装或启用。

### 选择 profile

`DSH_CLIENT_BUILD_PROFILE` 决定渲染哪个品牌。`official` 构建显示官方侧栏标志和名称；任何其他取值都让外壳回退——门形标志与本地构建标签——保持原样。会话首屏无论 profile 如何都显示来自 `dsh-client-ui-conversation` 的门形标记，因为这个回退本身就是官方标志。两种情况下插件都会照常加载并通过校验；只有侧栏图形注册受 profile 门控。

### 首页 frame 层

本包不占用 `conversation.hero.frame`。该 slot 仍由 `dsh-client-ui-conversation` 声明，可供其他 bundle 使用；空白首页保留静态门形标记和悬停门扇。移除此前点击触发的 Canvas 装饰时，也移除了上游鲸鱼图集及其 asset 路由。

### 替换品牌

自有身份的部署不组合本包，而是组合另一个占据侧栏 slot 和首屏标志 slot 的包；本包让首屏标志保持回退。占据 slot 是唯一的组合路径；品牌图形身份不可配置。

-----

<a id="understand-the-implementation"></a>
## 理解实现

<details>
<summary>实现细节——点击展开</summary>

两个侧栏填充作为一组声明感知的注册安装：嵌套的 `ctx.slots.inject()` 调用等待侧栏声明，因此无论本行在声明者之前还是之后激活，这组注册都能工作；声明消失时两个填充一并撤回，HMR 期间也不会留下残缺的品牌混合。浏览器半部是 [`src/client/index.ts`](src/client/index.ts)；node 半部是空的 Loader 席位。浏览器标题是构建环境的事（`DSH_CLIENT_TITLE`），不在 slot 系统之内。

</details>

-----

<a id="further-exploration"></a>
## 进一步探索

当品牌面不够用时阅读以下页面。它们从本包占据的 slot 进入渲染这些 slot 的外壳。

- [ui-sidebar](../ui-sidebar/README.zh.md)——声明 `sidebar.brand.mark` 与 `sidebar.brand.name` 并渲染其回退。
- [ui-conversation](../ui-conversation/README.zh.md)——声明 `conversation.hero.brand.mark` 与可用的 `conversation.hero.frame` 层。
- [Web 客户端架构](../../../docs/subsystems/web-client.zh.md)——浏览器插件行如何加载并注册 slot。

-----

<a id="model-experience"></a>
## 模型体验

无。本包只提供浏览器呈现，不触及模型请求。

#### KV Cache 影响

无；本包既不组装也不发送 provider 请求。

## 已知限制与延期工作

<a id="known-limitations-and-deferred-work"></a>

这些限制界定了品牌呈现的供给方式。它们是当前包约束，不是品牌设计对比或任务积压。

- **只有一组填充**——替代呈现属于另一个占据相同 slot 的 Cordis 包。
- **没有首页装饰**——本包让 `conversation.hero.frame` 保持为空；空白首页显示静态门形标记回退。
- **浏览器标题独立**——`DSH_CLIENT_TITLE` 在构建时选择标题文本，而不是通过 UI slot。

<a id="dev-note"></a>
### 开发备注

<details>
<summary>维护者工作上下文——点击展开</summary>

无。

</details>
