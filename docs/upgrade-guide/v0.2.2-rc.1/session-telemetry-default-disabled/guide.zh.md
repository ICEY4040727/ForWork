---
kind: upgrade-guide
description: "省略 session-telemetry 的 `mode`（包括基础 profile）现在会让 OpenTelemetry 采集保持禁用，而不是仅反馈模式。"
---

# 会话遥测默认禁用

[English](guide.md) | 中文

## 变更

`@deepseek-ai/dsh-session-telemetry-otel` 的 `mode` 默认值现在是 `DISABLED`。未提供 `mode` 就挂载后端时，不会构造协调器、处理器或导出器，因此活跃反馈和已存储历史都不会产生上传。随附的基础 profile 同样回退到 `DISABLED`；`DSH_TELEMETRY_MODE` 仍可覆盖该行。此前省略该模式会选择 `FEEDBACK_ONLY`，在新反馈后上传权威 Session 前缀。已显式设置 `mode: FEEDBACK_ONLY` 的 profile 或条目保持原有行为。

## 迁移

1. 在每个仍希望在反馈后上传的 profile 或 `cordis.yml` 条目中显式设置 `mode: FEEDBACK_ONLY`：

   ```yaml
   - id: sessionTelemetry-otel
     name: '@deepseek-ai/dsh-session-telemetry-otel'
     config:
       mode: FEEDBACK_ONLY
       exporter:
         url: https://collector.example.com/v1/logs
   ```

2. 在遥测必须保持关闭的位置移除该设置，或使用 `mode: DISABLED`。
3. 启动该 profile，提交文本反馈，并确认 `ctx.sessionTelemetry.sharing` 在需要上传时报告 `feedback-only`，否则报告 `disabled`。
