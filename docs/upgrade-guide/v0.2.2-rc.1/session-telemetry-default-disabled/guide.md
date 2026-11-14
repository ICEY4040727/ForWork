---
kind: upgrade-guide
description: "Omitting session-telemetry `mode`, including in the base profile, now leaves OpenTelemetry capture disabled instead of feedback-only."
---

# Session telemetry defaults to disabled

English | [中文](guide.zh.md)

## Change

`@deepseek-ai/dsh-session-telemetry-otel` now defaults `mode` to `DISABLED`. Mounting the backend without `mode` constructs no coordinator, processor, or exporter, so neither live feedback nor stored history produces an upload. The shipped base profile also falls back to `DISABLED`; `DSH_TELEMETRY_MODE` still overrides the row. Previously the omitted mode selected `FEEDBACK_ONLY`, which uploaded the canonical Session prefix after new feedback. A profile or entry that already sets `mode: FEEDBACK_ONLY` keeps its behavior.

## Migration

1. In every profile or `cordis.yml` entry that should keep uploading after feedback, set `mode: FEEDBACK_ONLY` explicitly:

   ```yaml
   - id: sessionTelemetry-otel
     name: '@deepseek-ai/dsh-session-telemetry-otel'
     config:
       mode: FEEDBACK_ONLY
       exporter:
         url: https://collector.example.com/v1/logs
   ```

2. Remove the setting or use `mode: DISABLED` where telemetry must stay off.
3. Start the profile, submit text feedback, and confirm `ctx.sessionTelemetry.sharing` reports `feedback-only` when uploads are intended, or `disabled` when they are not.
