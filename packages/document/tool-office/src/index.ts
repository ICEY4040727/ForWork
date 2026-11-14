/**
 * Office document tools plugin. The model-facing tools that read and edit Office
 * files through `ctx.officeExchange` are registered here.
 * @module @deepseek-ai/dsh-tool-office
 */

import type { Context } from '@deepseek-ai/cordis'

/**
 * Register the Office document tools on the Host context.
 * @param _ctx - owning Host context; no tool, prompt section, or session event is registered yet.
 */
export default function apply(_ctx: Context): void {}
