---
description: "在仓库测试和基准测试中定义测试工具，并调用工具流水线或选定的执行器。"
kind: "package-library"
---

# @deepseek-ai/dsh-tool-testkit

[English](README.md) | 中文

## 概述

使用本包定义生成内容的测试工具，并在不运行 Agent 回合的情况下测试工具执行。测试和基准测试将这些辅助函数作为开发依赖导入。调用方负责 Cordis context、插件组合和资源清理。

## 目录

- [使用本包](#use-this-package)
- [理解实现](#understand-the-implementation)
- [进一步探索](#further-exploration)
- [模型体验](#model-experience)
- [已知限制与后续工作](#known-limitations-and-deferred-work)
- [开发备注](#dev-note)

-----

<a id="use-this-package"></a>
## 使用本包

将 `@deepseek-ai/dsh-tool-testkit` 加入消费包的 `devDependencies`。测试只需要渲染内容而不需要领域结果值时，可以定义以下 fixture：

```ts
import { defineContentToolFixture } from '@deepseek-ai/dsh-tool-testkit'

const tool = defineContentToolFixture({
  name: 'echo',
  description: 'Return a test message.',
  parameters: {},
  execute: async () => [{ type: 'text', text: 'hello' }],
})
```

通过调用方的工具注册表注册返回的定义。使用 `executeToolFixture` 执行已解析的调用，或使用 `executeModelToolFixture` 将一次 JSON 编码的调用交给选定的执行器。后者要求调用方挂载该执行器，并在参数无法编码为 JSON 或执行器响应缺少结果时拒绝。调用字段见 [`src/index.ts`](src/index.ts)。

### 观察网络出口

`recordEgress(fn)` 在一个异步执行体运行期间包装 `fetch`、`node:http.request` 和 `node:https.request`，结束后恢复原函数并返回 `{ result, urls }`。它只观察调用而不替换传输，因此调用前安装的 stub 仍是实际执行者，测试无需真实网络 I/O。记录器是进程级且不可嵌套：第二个并发记录器会抛出错误。

`assertNoEgress(urls, { allow })` 在任一记录目标的主机名不在大小写不敏感的允许列表内时失败；空列表通过。必须调用模型时，将配置的模型端点主机放入 `allow`；连模型调用都违反隐私边界时省略该列表。

-----

<a id="understand-the-implementation"></a>
## 理解实现

<details>
<summary>实现细节——点击展开</summary>

[`src/index.ts`](src/index.ts) 将内容块适配为工具的规范 JSON 值，并将执行委托给生产流水线或注册表。它不创建 context，也不持有插件注册。

</details>

-----

<a id="further-exploration"></a>
## 进一步探索

- [工具包](../../core/tools/README.zh.md)——工具定义、注册表和执行策略。
- [测试策略](../../../docs/testing.zh.md)——为生产行为选择验证方式。
- [test-support 分组地图](../README.zh.md)——同组测试工具。

-----

<a id="model-experience"></a>
## 模型体验

无，这些测试辅助函数不组装或发送模型请求。

#### KV Cache 影响

无；本包自身不发送 provider 请求。

## 已知限制与后续工作

<a id="known-limitations-and-deferred-work"></a>

- 内容数组值仅用于测试 fixture；产品工具声明领域自己的输出 schema。
- 直接执行流水线不包含模型批次调度。两种执行辅助函数都不记录根 Session 日志；验证会话记录时应运行 Agent 回合。
- 插件组合和 context 释放仍由调用方负责。
- `recordEgress` 观察全局 `fetch` 绑定以及 `node:http.request` 和 `node:https.request`；`http.get`/`https.get`、直接导入的 `undici`，以及在记录开始前捕获的引用都不会经过它。

<a id="dev-note"></a>
### 开发备注

<details>
<summary>维护者工作上下文——点击展开</summary>

无。

</details>
