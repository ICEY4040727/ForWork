# Loadable Harness plugin packages

This file is GENERATED from workspace manifests (`scripts/gen-plugin-packages.ts`) and verified fresh by `pnpm run verify-plugin-packages` (part of `doc-sync`); do not edit it by hand.

Every package below exports a Cordis plugin that a bundle patch can name in a Loader row. `Config` marks packages whose row accepts a `config` mapping; query `Config.listConfigs` through `cordis_inspect_query` (filter by `name`, then query the `entry` id) for the mounted schema. Packages under `experimental` are pre-stable.

## acp

| Package | Config | Description |
|---|---|---|
| `@forwork/acp` | yes | Automation-only Agent Client Protocol server for driving DeepSeek Harness agents over JSON-RPC stdio |

## api

| Package | Config | Description |
|---|---|---|
| `@forwork/api-account-controller` | no | Expose safe account operations over authenticated Remote |
| `@forwork/api-gateway` | yes | Typert Remote Host dispatcher and Client API endpoint |
| `@forwork/api-job-controller` | yes | Job Remote observation stream and the reference-counted client job-output service |
| `@forwork/api-remote-access` | no | Method-level Remote access control for non-owner Peers and the per-Peer access snapshot |
| `@forwork/api-remotes` | no | Remote BFF assembly for application-selected Host capabilities |
| `@forwork/api-session-controller` | yes | Session Remote commands, cold reads, and live control transport |
| `@forwork/api-settings-controller` | yes | Remote owner for the configuration surfaces over the settings-domain seams |
| `@forwork/api-terminal-controller` | yes | Session-owned interactive terminals with shell discovery, screen recovery and typed Remote control |
| `@forwork/api-workspace-controller` | yes | Workspace Remote commands and reconnect-safe state transport |
| `@forwork/api-workspace-files` | yes | Workspace file service and Client resource provider: bounded reads, directory listing, and live metadata over the workspaceFiles Remote namespace |

## attachment

| Package | Config | Description |
|---|---|---|
| `@forwork/attachment-local` | yes | Private content-addressed DSH_HOME attachment storage |

## boot

| Package | Config | Description |
|---|---|---|
| `@forwork/config-editor` | no | Persist plugin configuration through profile patches and Loader reconciliation |
| `@forwork/hmr` | yes | Coordinated module and profile configuration hot reload |
| `@forwork/plugin-manager` | yes | Current-profile plugin and bundle management shared by dsh CLI, Web and agent tools |
| `@forwork/plugin-market` | yes | Configured plugin market indexes for the Web Discover page and the plugin_market agent tool |

## browser-use

| Package | Config | Description |
|---|---|---|
| `@forwork/browser-use` | no | Exclusive named browser-use provider registration |

## bundle

| Package | Config | Description |
|---|---|---|
| `@forwork/acp-app` | no | The dsh ACP profile bundle: automation-only JSON-RPC stdio and process lifecycle over dsh-base |
| `@forwork/headless` | yes | The dsh one-shot bundle: a direct core Agent/Session runner over dsh-base with no Host, HTTP, or browser layer |
| `@forwork/sdk-app` | yes | The dsh SDK profile bundle: stdio JSON-RPC serving and process lifecycle over dsh-base |
| `@forwork/web-app` | yes | The dsh browser-surface bundle: the web patch layer over dsh-base plus the runtime glue plugin (frontend dist serving, bash runtime variables, URL line) |

## client

| Package | Config | Description |
|---|---|---|
| `@forwork/client-connection` | yes | Authenticated RPC transport and generation lifecycle |
| `@forwork/client-file-upload` | no | Agent-scoped browser file upload, streaming intake, and staged receipt service |
| `@forwork/client-hmr` | yes | Web client graph synchronization and rebuilt-bundle reload transport |
| `@forwork/client-locale` | no | Locale plugin: Host-backed preference, extensible language catalog, browser fallback, and typed built-in dictionaries |
| `@forwork/client-modules` | no | Client module system, dual-face: node half composes the __DSH_BOOT__ entry graph (incremental dsh.client scan, bundle route, index tap, webPlugins service); browser half is the lazy-CJS module table the vendored cordis Loader consumes as its internal seam |
| `@forwork/client-product-analytics` | yes | Desktop product event collection and authenticated Host reporting |
| `@forwork/client-resources` | no | Unified client resource model: protocol-registered providers turn URL addresses into live values, consumed through the useResource global standard hook |
| `@forwork/client-shortcuts` | yes | Application keyboard command registry and physical-key routing |
| `@forwork/client-ui-agent-preset` | no | Agent-preset surfaces: the default for later sessions, this session's seat, and the composition editor |
| `@forwork/client-ui-agent-team` | no | Web shared task board for a subagent tree |
| `@forwork/client-ui-approval` | no | Approval composer takeover over the scoped Remote Event waterfall |
| `@forwork/client-ui-attachment` | no | Dynamic attachment presentation plugin for conversation input, message-image, and trajectory image slots |
| `@forwork/client-ui-brand-official` | no | Official DeepSeek Harness brand artwork and click-triggered home background |
| `@forwork/client-ui-chat` | no | Chat Conversation target, node definitions, renderers, and details surface |
| `@forwork/client-ui-commands` | no | Client command surface: global directory cache, '/' source, three command UI kinds, popupSelect registry |
| `@forwork/client-ui-context-menu` | no | Typed, plugin-extensible context menus with captured targets and shared Web presentation |
| `@forwork/client-ui-conversation` | no | Target-neutral Conversation assembly, shell, composer, queue, and view navigation |
| `@forwork/client-ui-deliverables` | no | Changed-files card with per-file comparison tabs, delivery cards, and clickable final-response file references for Web |
| `@forwork/client-ui-directory-picker-browse` | no | In-app directory browsing surface: the workspace directory-flow owner rendering the host's listing and creation primitives |
| `@forwork/client-ui-directory-picker-native` | no | Native directory-picker surface: the renderless workspace directory-flow occupant driving the local Desktop or Host OS chooser |
| `@forwork/client-ui-entry-customization` | no | Customize plugin entries in the left sidebar, conversation views, and right sidebar, and quickly create new plugins. |
| `@forwork/client-ui-file-actions` | no | Original-file saving and image clipboard actions for Web and Desktop |
| `@forwork/client-ui-file-drag` | no | Application-owned file drag transfers and localized failure feedback |
| `@forwork/client-ui-goal` | no | Session goal surface: GoalBar docked above the composer, read from the goal session projection |
| `@forwork/client-ui-input-trigger` | no | Input trigger pipeline: '/' and '@' detection, candidate menu, pick routing to registered sources |
| `@forwork/client-ui-jobs` | no | Session-header background-job list with on-demand streaming record panels |
| `@forwork/client-ui-layout` | no | Shell plugin: three-column AppFrame with drag handles, ctx.layout viewing-state service (navigation + panels) |
| `@forwork/client-ui-message-feedback` | no | The Web feedback surface: per-message Like/Dislike in the assistant-message action strip and the feedback dialog behind both ratings and /feedback, backed by the messageFeedback and sessionFeedback Host Remotes |
| `@forwork/client-ui-model-selection` | no | Model selection over the shared model catalog, Session projection, and session.selectModel |
| `@forwork/client-ui-office-editor` | no | Embedded Office document editor tab for the right Sidebar |
| `@forwork/client-ui-open-in-app` | no | Session application submenus and native opening controls for directories and previewed files |
| `@forwork/client-ui-permission-presets` | no | Permission surfaces: a new-session default in General settings and a current-session /permission popup over the permissions projection |
| `@forwork/client-ui-plan` | no | Plan mode controls, persistent transcript plan cards, and sidebar Markdown previews |
| `@forwork/client-ui-plugin-manager` | yes | Plugin management for the dsh web client: the sidebar Plugins panel installs, enables, disables, retries, and composes installed plugin packages |
| `@forwork/client-ui-plugin-market` | no | Discover tab of the dsh web Plugins page: browse configured plugin market indexes and install their bundles |
| `@forwork/client-ui-reference` | no | Unified Web @file and @session reference source |
| `@forwork/client-ui-renderer` | no | Browser UI renderer: React slot bindings, ctx.uiRenderer, and the assembled application root |
| `@forwork/client-ui-schedule` | no | Host task management page and Session reminder catalog |
| `@forwork/client-ui-secret-mask` | no | Web display masking for recognized credentials in sent user text |
| `@forwork/client-ui-session` | no | Session Controller adapter for React and session-scoped slots |
| `@forwork/client-ui-settings` | no | Settings domain base plugin: shared configuration forms and the canonical settings slot-type contract |
| `@forwork/client-ui-settings-account` | yes | Manage DeepSeek login and open Platform billing pages |
| `@forwork/client-ui-settings-agent-loop` | no | Settings page of the agent loop on the dsh web client's Plugins page: the parallel tool-call cap of the agent-loop namespace |
| `@forwork/client-ui-settings-general` | no | Settings ownerless-copy and product onboarding plugin: the General section, shell trigger/header chrome content, settings dictionaries, and the versioned welcome notice |
| `@forwork/client-ui-settings-models` | yes | Models settings and shared product-onboarding dialogs over existing settings and credential joins |
| `@forwork/client-ui-settings-plugin-inventory` | no | Read-only Cordis Loader inventory tab in Web Plugins settings |
| `@forwork/client-ui-settings-plugins` | no | Built-in plugins settings section for the dsh web client: the Settings navigation entry and the tab chrome feature-owned tabs register into |
| `@forwork/client-ui-settings-session-log` | no | General settings control for Session-log upload with DeepSeek API requests |
| `@forwork/client-ui-settings-shell` | no | Settings page of the shell executor on the dsh web client's Plugins page: the command timeout and the per-stream output cap of the shell namespace |
| `@forwork/client-ui-settings-subagent` | no | Settings page of Subagent delegation on the dsh web client's Plugins page: recursion depth, parallel capacity, and the models agents may choose for subagents |
| `@forwork/client-ui-settings-web-search` | no | Settings page for web search and fetch on the dsh web client's Plugins page: the Search and Fetch provider menus, the DeepSeek Search API result count, and legacy Messages search settings |
| `@forwork/client-ui-shortcuts` | no | Keyboard shortcut reference, recording, and local preference editing |
| `@forwork/client-ui-sidebar` | no | Sidebar plugin: session multi-level tree, search, grouping, state dots |
| `@forwork/client-ui-sidebar-browser` | no | Sandboxed Web browser tabs for the right Sidebar |
| `@forwork/client-ui-sidebar-documentpreview` | yes | Extensible Sidebar previews for Office documents, spreadsheets, Markdown, code, images, PDF, HTML, and plain text |
| `@forwork/client-ui-sidebar-files` | no | The workspace files page and one workspace directory tree per pane beside docked and floating file previews, listed lazily over the workspaceFiles Remote namespace |
| `@forwork/client-ui-sidebar-right` | no | Right Sidebar: the docking surface's session-bound state, its panel and header expand control, and the navigation service over it |
| `@forwork/client-ui-sidebar-terminal` | no | Interactive shell tabs for the right Sidebar |
| `@forwork/client-ui-skill` | no | Web skill references and the dedicated skill tool row |
| `@forwork/client-ui-subagent` | no | Subagent conversation catalog, continuation routing UI, and '@' reference source |
| `@forwork/client-ui-theme` | yes | Theme plugin: Host bootstrap for the pre-plugin palette; ThemeRuntime for light/dark/system state and icon replacements; --dsw-* token styles and Appearance settings row |
| `@forwork/client-ui-tool` | no | Client Tool call-tree renderer and keyed per-tool presentation slot |
| `@forwork/client-ui-trajectory` | yes | Trajectory event ledger and timing overview with live tool-result inspection |
| `@forwork/client-ui-user-questions` | no | Web ask_user_question composer takeover and plan-review presentation UI |
| `@forwork/client-ui-workflow-run` | no | Durable workflow-run Conversation Node and nested member disclosure for dsh web |
| `@forwork/client-ui-workspace` | no | Workspace picker plugin: one WorkspacePicker registered into the sidebar and empty-state workspace slots |

## compaction

| Package | Config | Description |
|---|---|---|
| `@forwork/command-compact` | no | Human-facing slash command for explicit session compaction |
| `@forwork/compaction-basic` | yes | Token-meter-driven compaction policy and LLM summarization backend for the DeepSeek Harness |
| `@forwork/compaction-image-offload` | no | Durable image offload for image-capable routes: replace over-budget request images with placeholders and retry |
| `@forwork/compaction-tool-result-pruner` | yes | Replay-safe model-free head/middle/tail pruning for tool-result surface nodes |

## computer-use

| Package | Config | Description |
|---|---|---|
| `@forwork/computer-use` | no | Exclusive named computer-use provider registration |

## context

| Package | Config | Description |
|---|---|---|
| `@forwork/agent-instructions` | yes | Workspace context loader for AGENTS.md/CLAUDE.md instruction files |
| `@forwork/file-reference-local` | yes | Local-filesystem ctx.fileReferences provider with bounded fuzzy indexes |
| `@forwork/harness-dev-context` | no | Model context naming the DeepSeek Harness checkout and Web GUI URL without letting either change the system prompt |
| `@forwork/session-reference` | yes | Cross-session snapshot references and durable untrusted model context (ctx.sessionReferenceResolver) |
| `@forwork/time-context` | yes | Durable per-step context with the current time and elapsed time |
| `@forwork/tmux-context` | yes | Opt-in durable per-step context with this agent's tmux pane and window location |

## core

| Package | Config | Description |
|---|---|---|
| `@forwork/agent` | no | Agent interface, registry, initiator scope, and event vocabulary for the DeepSeek Harness |
| `@forwork/agent-default-model` | yes | Default model selection shared by Agent entry points |
| `@forwork/agent-loop` | yes | The concrete agent loop plugin for the DeepSeek Harness |
| `@forwork/agent-tool-presentation` | yes | Deprecated tool presentation selector; use dsh-tool-executor-native or dsh-tool-executor-ptc directly |
| `@forwork/session` | no | Event-sourced session store for the DeepSeek Harness |
| `@forwork/system-prompt` | yes | System prompt assembly registry for the DeepSeek Harness |
| `@forwork/tool-executor-native` | no | Native tool execution plugin |
| `@forwork/tool-executor-ptc` | yes | PTC program execution, tool bindings, and generated SDK prompts |
| `@forwork/tools` | no | Tool registry and execution pipeline for the DeepSeek Harness |

## credentials

| Package | Config | Description |
|---|---|---|
| `@forwork/authorization` | no | Authorization seam (ctx.authorization): plugin-owned flows that obtain a credential through a conversation with the human |
| `@forwork/credentials-local` | yes | File-backed credentials provider ($DSH_HOME/.env under the live process environment) for the DeepSeek Harness |
| `@forwork/deepseek-account-platform` | yes | Authorize DeepSeek accounts through browser PKCE |

## deliverables

| Package | Config | Description |
|---|---|---|
| `@forwork/deliverable-file-references` | no | Shared model guidance for file links and deliverables across DeepSeek Harness hosts |
| `@forwork/tool-present` | yes | Explicit workspace file delivery declarations for the DeepSeek Harness |
| `@forwork/workspace-changes` | yes | Per-turn workspace file changes recorded from git working-tree snapshots and whole-file captures, with per-file comparisons, for the DeepSeek Harness |

## document

| Package | Config | Description |
|---|---|---|
| `@forwork/office-to-pdf` | yes | Shared Office-to-PDF conversion with bounded queues and caching |
| `@forwork/tool-office` | yes | Model-facing Office document tools backed by Office Exchange |

## experimental

| Package | Config | Description |
|---|---|---|
| `@forwork/action-cards` | yes | Optional action cards below answers: one shared proposal tool, card state, user actions and status notices |
| `@forwork/client-ui-action-cards` | no | Action cards below completed answers: the shared card frame, kind bodies and the compact tool row |
| `@forwork/client-ui-creation-recommendations` | no | Body of the personalized page card: location choice, creation progress, Undo and Restore |
| `@forwork/client-ui-plugin-recommendations` | no | Official plugin card body with catalog text and enabling |
| `@forwork/creation-recommendations` | no | Create persistent removable interface plugins from explicit conversation needs |
| `@forwork/cua` | no | Shared Computer Use application objects, authorization, and logged tool execution |
| `@forwork/experimental-api-speech-to-text` | yes | Authenticated experimental speech transcription for browser clients |
| `@forwork/experimental-auto-review` | no | Per-tool LLM authorization review for the DeepSeek Harness Auto permission preset |
| `@forwork/experimental-browser-use-chrome-devtools-mcp` | yes | Experimental per-Session Chromium browser tools through chrome-devtools-mcp |
| `@forwork/experimental-browser-use-playwright-mcp` | yes | Experimental per-Session Chromium browser tools through @playwright/mcp |
| `@forwork/experimental-browser-use-stagehand-native` | yes | Experimental Stagehand browser tools with separately configured native models |
| `@forwork/experimental-claude-code-mods` | yes | Experimental bridge: load Claude Code mods (hooks modules) and run their hook chains on DeepSeek Harness extension points |
| `@forwork/experimental-client-ui-claude-code-mods` | no | Web band above the prompt for Claude Code mods: draws each session's mod tree and sends button presses back to the bridge |
| `@forwork/experimental-client-ui-cot-translation` | yes | Optional machine translation of expanded reasoning, preserving original Session text |
| `@forwork/experimental-client-ui-forge` | yes | Sidebar panel for a Session's change requests, with a worktree option for new Sessions and its Host Remote |
| `@forwork/experimental-client-ui-stats` | no | Per-turn and per-step context, timing, cache, and tool-cost table contributed as an experimental conversation View |
| `@forwork/experimental-client-ui-voice-input` | no | Record speech and insert editable text into the conversation draft |
| `@forwork/experimental-computer-use-cua-driver-mcp` | yes | Experimental computer use through an installed Cua Driver MCP executable |
| `@forwork/experimental-computer-use-cua-driver-native` | no | Experimental computer-use provider embedding the Cua Driver native npm SDK |
| `@forwork/experimental-cua-macos` | yes | macOS native provider for shared Computer Use application and window objects |
| `@forwork/experimental-forge` | yes | Provider-neutral change request registry with shared status polling |
| `@forwork/experimental-forge-github` | yes | GitHub provider for the experimental forge registry, read through the GitHub CLI on the host |
| `@forwork/experimental-forge-session` | yes | Binds change requests to the Session that created, pushed, or tracked them |
| `@forwork/experimental-inspector` | yes | Experimental cross-realm CDP hub for Host debugging and Client Runtime inspection |
| `@forwork/experimental-ptc-runtime-python` | yes | CPython subprocess implementation of the DeepSeek Harness PTC execution seam |
| `@forwork/experimental-session-inspector` | no | Experimental virtualized Session log and live Chat group/node inspectors |
| `@forwork/experimental-speech-to-text` | yes | Experimental speech recognition with independently selectable providers |
| `@forwork/experimental-speech-to-text-sensevoice` | yes | Local SenseVoice ONNX transcription with a managed sherpa-onnx process |
| `@forwork/experimental-subagent-oar` | yes | Experimental one-shot subagents through OAR runtimes |
| `@forwork/experimental-tool-subagent-oar` | yes | Delegate standalone tasks to configured external subagent providers |
| `@forwork/experimental-tool-worktree` | no | Model tool that creates and enters a new Git worktree |
| `@forwork/experimental-translator` | yes | Machine translation with reusable Session results |
| `@forwork/experimental-worktree` | yes | Creates a named Git worktree and changes the calling Session working directory |
| `@forwork/hooks-claude-code` | yes | Bridge plugin: run a Claude Code hooks.json / settings hook config on the DeepSeek Harness interception seams |
| `@forwork/hooks-codex` | yes | Bridge plugin: run a Codex hooks.json hook config on the DeepSeek Harness interception seams |
| `@forwork/plugin-recommendations` | no | Official plugin suggestions with explicit user activation and persistent conversation decisions |
| `@forwork/session-title-all-prompts-llm` | yes | All-user-messages LLM provider plugin for DeepSeek Harness session titles |
| `@forwork/skill-badge` | no | Bundled dsh badge skill provider for DeepSeek Harness |
| `@forwork/tool-ralph` | yes | Model-facing fresh-agent Ralph loop over the workflow and subagent seams |
| `@forwork/tool-session-query` | yes | Workspace-authorized model-facing session history search, trace, and event read tools |
| `@forwork/tool-terminal` | yes | Six model-facing persistent PTY tools with owner isolation and generic background-job integration |
| `@forwork/webhook` | no | Fire-and-forget webhook rule runtime that creates Workspace-backed DeepSeek Harness Sessions |
| `@forwork/webhook-github` | yes | Signed GitHub HTTP webhook adapter for the DeepSeek Harness webhook runtime |

## extensions

| Package | Config | Description |
|---|---|---|
| `@forwork/client-ui-cordis` | no | Cordis dynamic-plugin definition card: the keyed cordis_define tool row with its run/stop switch |
| `@forwork/cordis-client-runner` | no | Browser half of dynamic dual-half plugin packages: event subscription, closure evaluation, guard facade, and loader entries |
| `@forwork/cordis-host-runner` | yes | Dynamic package definition registry, host-half sandbox lifecycle, and invoke handler table for model-mounted dual-half packages |
| `@forwork/tool-cordis` | no | Read-only runtime API inspection for Harness plugin development |

## feedback

| Package | Config | Description |
|---|---|---|
| `@forwork/command-feedback` | no | Log-only session feedback: the record event, the sessionFeedback Host Remote, and the human-facing slash command |
| `@forwork/message-feedback` | yes | Canonical Session-log ratings and notes for finalized assistant messages |

## fs

| Package | Config | Description |
|---|---|---|
| `@forwork/fs-local` | yes | Local-filesystem implementation of the DeepSeek Harness filesystem seam (ctx.fs) |
| `@forwork/fs-observation-policy` | no | File-context policy plugin for the DeepSeek Harness — observed-state, read-before-edit, and version-guarded write/edit added over the ctx.fs provider seam through the fs/* event gate (no service API) |
| `@forwork/fs-sandbox` | yes | Sandbox-enforcing implementation of the DeepSeek Harness filesystem seam: fences write/edit by the per-call sandbox mode (read-only denies mutation, workspace-write contains it to the workspace + temp roots) while reads pass through |
| `@forwork/tool-fs` | yes | Model-facing filesystem tools (read, write, edit) over the DeepSeek Harness filesystem seam (ctx.fs) |
| `@forwork/tool-fs-search` | yes | Model-facing filesystem discovery tools (glob, grep) backed by the packaged ripgrep binary (@vscode/ripgrep) |
| `@forwork/tool-str-replace-editor` | yes | Model-facing view, create, literal replace, and line insert tool over the Harness filesystem service |

## goal

| Package | Config | Description |
|---|---|---|
| `@forwork/command-goal` | no | Human-facing slash command for persisted same-session goals |
| `@forwork/goal` | yes | Event-sourced same-session goal state and lifecycle service for the DeepSeek Harness |
| `@forwork/goal-round-driver` | no | Race-fenced same-session goal-round driver |
| `@forwork/tool-goal` | yes | Model-facing same-session goal tools with execution-time authority checks |

## guard

| Package | Config | Description |
|---|---|---|
| `@forwork/repeat-tool-reminder` | yes | Repeat-tool-call guard plugin: advisory reminders when an agent loops on identical tool calls |
| `@forwork/tool-call-timeout-policy` | no | Tool-call timeout policy: a tools/execute wrapper that arms a per-tool deadline on exec.signal and returns TOOL_TIMEOUT when it wins |

## host

| Package | Config | Description |
|---|---|---|
| `@forwork/host-directory-picker-auto` | no | Adaptive chooser of the directory-picker seam: resolves the host situation at boot and mounts the native or browse backend for the DeepSeek Harness web GUI host |
| `@forwork/host-directory-picker-browse` | yes | In-app browsing backend of the directory-picker seam (listing/creation primitives over the host filesystem) |
| `@forwork/host-directory-picker-native` | no | Native-OS-chooser backend of the directory-picker seam for the DeepSeek Harness web GUI host |
| `@forwork/host-frontend-static` | yes | SPA dist server for the Web shell: owns the webserver fallback seat, serving explicit index entries and static assets with traversal rejection and 404 misses |
| `@forwork/host-open-in-app` | yes | Host half of open-in-app: resolved application catalog, icons, and the launch endpoint as three webServer routes |
| `@forwork/host-plugin-inventory` | no | Read-only Remote projection of current Cordis Loader plugin state |
| `@forwork/host-product-telemetry-otel` | yes | Explicit product usage events exported through OpenTelemetry HTTP logs |
| `@forwork/host-webserver` | yes | Web route-registration plugin: HTTP and upgrade routes, per-Peer index transform taps, and static dist fallback |

## interaction

| Package | Config | Description |
|---|---|---|
| `@forwork/commands` | no | Plugin-owned human command registry for DeepSeek Harness UIs |
| `@forwork/permission-presets` | yes | User-facing permission presets (ctx.permissionPresets) for the DeepSeek Harness: one product-level Permissions select bundling the sandbox-mode and approval-policy knobs, written through to their own session events |
| `@forwork/tool-ask-user` | yes | Model-facing ask_user_question tool over the ctx.userQuestions seam |
| `@forwork/user-approval` | yes | User-approval seam (ctx.approval) for the DeepSeek Harness: one-shot permission decisions dispatched to composed answerers over the approval/request waterfall, fail-closed by default |
| `@forwork/user-questions` | no | Abstract user-questions seam (ctx.userQuestions) for asking the human during agent runs |

## jobs

| Package | Config | Description |
|---|---|---|
| `@forwork/jobs-local` | yes | Process-local implementation of the DeepSeek Harness background job registry seam |
| `@forwork/tool-jobs` | yes | Model-facing background job control tools (job_output, job_list, job_kill) over the ctx.jobs registry |

## llm

| Package | Config | Description |
|---|---|---|
| `@forwork/deepseek-llm-api-extensions` | no | Additive request-field registry for the official DeepSeek LLM API adapter |
| `@forwork/llm` | no | Provider-neutral LLM service interface for the DeepSeek Harness |
| `@forwork/llm-deepseek-account` | yes | DeepSeek account provider authentication and discovery |
| `@forwork/llm-deepseek-api-key` | yes | DeepSeek api-key provider authentication and discovery |
| `@forwork/llm-pi-ai` | yes | pi-ai-backed DeepSeek adapter for the DeepSeek Harness LLM seam (design-verification twin of dsh-llm-deepseek) |
| `@forwork/llm-retry` | yes | Provider-routed LLM request retry policy for the DeepSeek Harness |
| `@forwork/plugin-package-inventory-deepseek` | yes | Active Loader-backed plugin package inventory for official DeepSeek LLM API requests |
| `@forwork/token-meter` | yes | Replay-aware token measurement service (ctx.tokenMeter) for the DeepSeek Harness |

## lsp

| Package | Config | Description |
|---|---|---|
| `@forwork/lsp` | no | Abstract LSP capability seam (ctx.lsp) for the DeepSeek Harness — language-server provider registry keyed by branded id and extension mapping, order-independent per-query selection, normalized definition/references/implementation/hover requests and results, and the LspError taxonomy |
| `@forwork/lsp-stdio` | yes | Generic stdio language-server provider for the DeepSeek Harness LSP capability seam (ctx.lsp) — spawns configured servers, translates JSON-RPC, and serves transient-open goToDefinition/findReferences/goToImplementation/hover queries in the host filesystem namespace |
| `@forwork/tool-lsp` | yes | Model-facing lsp tool over the DeepSeek Harness LSP capability seam (ctx.lsp) — one read-only tool with goToDefinition/findReferences/goToImplementation/hover operations, one-based UTF-16 cursor coordinates, bounded location rendering, and hover normalization |

## mcp

| Package | Config | Description |
|---|---|---|
| `@forwork/mcp-client` | yes | MCP client bridge: connects to MCP servers and registers their tools on ctx.tools |
| `@forwork/mcp-manager` | yes | MCP server manager: a settings-driven fleet of mcp-client instances with runtime status, restart, and probe |
| `@forwork/mcp-resources` | no | Scoped MCP resource discovery and reading through shared model tools |

## plan

| Package | Config | Description |
|---|---|---|
| `@forwork/plan-mode` | yes | Logged per-agent plan mode with deployment guidance, a direct slash command, and a user-reviewed exit |

## preset

| Package | Config | Description |
|---|---|---|
| `@forwork/agent-preset` | yes | Declare an Agent capability composition in Cordis YAML |
| `@forwork/agent-preset-registry` | yes | Declarative Agent preset registry and profile-backed editing |
| `@forwork/persona` | yes | Composition-authored deployment persona section for the DeepSeek Harness |

## ptc-runtime

| Package | Config | Description |
|---|---|---|
| `@forwork/ptc-runtime-node` | yes | Node PTC execution with async function, REPL and VM Script evaluators and optional retained state |

## sandbox

| Package | Config | Description |
|---|---|---|
| `@forwork/sandbox-local` | yes | Local process-sandbox backends for the DeepSeek Harness sandbox seam: bwrap, the npm-distributed landlock-run launcher, macOS Seatbelt, or the Windows ACL restricted-token runner — functionally probed, fail-closed |
| `@forwork/sandbox-policy` | yes | Per-call sandbox policy resolver and current model context: deployment fallbacks plus each session's mode and workspace root, shared by every enforcing capability family |

## schedule

| Package | Config | Description |
|---|---|---|
| `@forwork/schedule` | yes | Host-wide durable reminders with shared management and delivery to existing or new Sessions |
| `@forwork/tool-schedule` | no | Model-facing scheduled task tools (schedule_create, schedule_list, schedule_view, schedule_edit) over the Host ctx.schedule service |

## sdk

| Package | Config | Description |
|---|---|---|
| `@forwork/sdk-jsonrpc-server` | yes | Stdio JSON-RPC server plugin for out-of-process DeepSeek Harness SDK clients |

## session

| Package | Config | Description |
|---|---|---|
| `@forwork/message-quotes` | no | Selected-text quotation and Markdown blockquote presentation |
| `@forwork/session-checkpoint-policy` | no | Semantic session durability checkpoints before model requests and tool side effects |
| `@forwork/session-log-deepseek` | yes | Incremental lossless session-log request extension for the official DeepSeek LLM API |
| `@forwork/session-persistence-jsonl` | yes | JSONL durable session persistence backend for the DeepSeek Harness |
| `@forwork/session-projection` | no | Session-projection seam: the merge-extensible projection type table, the provider contract, and the ctx.sessionProjections registry serving whole current values of log-derived per-session state |
| `@forwork/session-projection-cache` | yes | Persisted projection cache (ctx.sessionProjectionCache): durable per-session checkpoint records on the session_projcache storage domain (per-record layout), throttled write-behind, and the cached listing read |
| `@forwork/session-stats` | no | Whole-log conversation counts and wall times projection (sessionStats) for the DeepSeek Harness |
| `@forwork/session-telemetry-otel` | yes | Feedback-authorized Session logs over byte-bounded OpenTelemetry HTTP requests |
| `@forwork/session-title` | yes | Log-backed session title service and provider registry for the DeepSeek Harness |
| `@forwork/session-title-first-prompt-llm` | yes | First-message LLM provider plugin for DeepSeek Harness session titles |
| `@forwork/session-turn-outline` | no | Whole-log turn outline projection (turnOutline) for the DeepSeek Harness |
| `@forwork/tool-working-directory` | no | Read and change the active Session working directory |
| `@forwork/working-directory` | yes | Session working directories with durable changes and model context |

## session-query

| Package | Config | Description |
|---|---|---|
| `@forwork/session-log-export` | yes | Web Session-log ZIP export and Session Markdown copy |
| `@forwork/session-query-sqlite` | yes | Concrete ctx.sessionQuery backend with SQLite FTS5 search |

## settings

| Package | Config | Description |
|---|---|---|
| `@forwork/settings` | no | Abstract user-settings seam (ctx.settings) for the DeepSeek Harness |

## shell

| Package | Config | Description |
|---|---|---|
| `@forwork/bash-local` | yes | Local-subprocess implementation of the DeepSeek Harness bash executor seam |
| `@forwork/bash-sandbox` | yes | Sandbox-consuming implementation of the DeepSeek Harness bash executor seam (confines every command via ctx.sandbox, reports denial/enforcement result facts) |
| `@forwork/pwsh-local` | yes | Local PowerShell implementation of the DeepSeek Harness bash executor seam |
| `@forwork/pwsh-sandbox` | yes | Sandbox-consuming implementation of the DeepSeek Harness PowerShell executor seam (confines every command via ctx.sandbox, reports denial/enforcement result facts) |
| `@forwork/shell-env` | yes | Tool-independent managed DSH_* shell environment registry |
| `@forwork/tool-bash` | yes | Model-facing bash tool with optional generic background-job and sandbox-escalation support |
| `@forwork/tool-bash-persistent` | yes | Model-facing owner-scoped persistent Bash tool backed by the Harness PTY service |
| `@forwork/tool-pwsh` | yes | Model-facing pwsh tool over the bash executor seam |
| `@forwork/tool-pwsh-persistent` | yes | Model-facing owner-scoped persistent PowerShell tool backed by the Harness PTY service |

## skill

| Package | Config | Description |
|---|---|---|
| `@forwork/skill` | yes | Agent skill provider registry for the DeepSeek Harness |
| `@forwork/skill-filesystem` | yes | Local filesystem skill provider for the DeepSeek Harness |
| `@forwork/skill-office` | yes | Bundled Word, PowerPoint, and Excel workflows and LibreOffice Kit operations |
| `@forwork/tool-skill` | yes | Model-facing skill loading tool for the DeepSeek Harness |
| `@forwork/tool-workspace-dependencies` | yes | The load_workspace_dependencies tool: absolute paths into a bundled Python, Node.js, and pnpm payload |

## spill

| Package | Config | Description |
|---|---|---|
| `@forwork/spill-local` | yes | Local-filesystem implementation of the DeepSeek Harness spill storage seam (private session-scoped files) |
| `@forwork/spill-policy` | yes | Token-budgeted tool-result retention with recoverable text and image paths |

## ssh

| Package | Config | Description |
|---|---|---|
| `@forwork/fs-ssh` | no | Filesystem provider over the shared POSIX SSH helper |
| `@forwork/sandbox-ssh` | no | Remote POSIX sandbox argv provider over the shared SSH helper |
| `@forwork/ssh` | yes | Shared OpenSSH connection and versioned POSIX remote helper |
| `@forwork/subprocess-ssh` | no | Subprocess and terminal provider over the shared POSIX SSH helper |

## storage

| Package | Config | Description |
|---|---|---|
| `@forwork/storage` | no | Storage hub (ctx.storage): named backend registry plus mounted data-form facilities for the DeepSeek Harness |
| `@forwork/storage-domain` | yes | Domain data form (ctx.storage.domain): schema-validated, event-emitting KV domains over storage backends for the DeepSeek Harness |
| `@forwork/storage-json` | yes | JSON file KV storage backend for the DeepSeek Harness storage hub |
| `@forwork/storage-sqlite` | yes | SQLite storage backend (kv facet) for the DeepSeek Harness storage hub |

## subagent

| Package | Config | Description |
|---|---|---|
| `@forwork/subagent` | yes | Abstract subagent seam (ctx.subagents): named-provider registry for delegating to child agents |
| `@forwork/subagent-acp` | yes | Out-of-process ACP subagent backend: drives a child agent in a spawned subprocess over the Agent Client Protocol |
| `@forwork/subagent-claude-code` | yes | One-shot Claude Code subagent provider over the official Agent SDK |
| `@forwork/subagent-codex` | yes | One-shot Codex subagent provider over the official app-server protocol |
| `@forwork/subagent-dsh-sdk` | yes | Out-of-process SDK subagent backend: drives a child DeepSeek Harness runtime subprocess over stdio JSON-RPC through the TypeScript SDK client |
| `@forwork/subagent-fork-in-process` | yes | In-process fork subagent backend: runs a child agent seeded with a prefix of the parent's log |
| `@forwork/subagent-spawn-in-process` | yes | In-process spawn subagent backend: runs a fresh child agent on ctx.agents |
| `@forwork/tool-subagent` | yes | Model-facing subagent delegation tool over the ctx.subagents seam |
| `@forwork/tool-subagent-control` | yes | Shared messaging, interruption, and child discovery tools for subagents |

## subprocess

| Package | Config | Description |
|---|---|---|
| `@forwork/subprocess-local` | no | Local-subprocess implementation of the DeepSeek Harness subprocess seam |

## team

| Package | Config | Description |
|---|---|---|
| `@forwork/agent-team` | yes | Shared task DAG and change waits for subagent trees |
| `@forwork/tool-agent-team` | yes | Scoped model-facing Agent Teams tools over ctx.agentTeams |

## telemetry

| Package | Config | Description |
|---|---|---|
| `@forwork/otel` | no | Cordis service for independent ordinary-event and byte-bounded Session-log OTLP channels |

## terminal

| Package | Config | Description |
|---|---|---|
| `@forwork/terminal` | no | Persistent PTY session seam for the DeepSeek Harness — owner-scoped ids, backend registry, interactive sends, reads, signals, and awaited cleanup |
| `@forwork/terminal-bash` | yes | Persistent shell PTY backend over the DeepSeek Harness subprocess terminal primitive |

## test-support

| Package | Config | Description |
|---|---|---|
| `@forwork/llm-replay` | yes | Replay LLM plugin: short-circuits llm/stream with model chunks reconstructed from a recorded session JSONL (keyless snapshot tests) |

## todo

| Package | Config | Description |
|---|---|---|
| `@forwork/tool-todo` | yes | Model-facing todo_write tool over the DeepSeek Harness event-sourced session log |

## typert

| Package | Config | Description |
|---|---|---|
| `@forwork/typert-loader` | yes | Loader integration for generated Typert package contributions |

## web

| Package | Config | Description |
|---|---|---|
| `@forwork/tool-web` | yes | Model-facing web tools (web_search, web_fetch) over the DeepSeek Harness web capability seam (ctx.web) |
| `@forwork/web` | yes | Abstract web access capability seam (ctx.web) for the DeepSeek Harness — search/fetch provider registry, registration-order-independent selection, request/result vocabulary, and the WebError taxonomy |
| `@forwork/web-deepseek-api` | yes | DeepSeek Search and Fetch API providers for the DeepSeek Harness web capability seam (ctx.web) |
| `@forwork/web-fetch-http` | yes | Anonymous HTTP(S) fetch provider for the DeepSeek Harness web capability seam (ctx.web) |
| `@forwork/web-search-deepseek` | yes | DeepSeek-backed search provider (native web_search via the Anthropic-compatible API) for the DeepSeek Harness web capability seam (ctx.web) |

## workflow

| Package | Config | Description |
|---|---|---|
| `@forwork/tool-workflow` | yes | Model-facing workflow tool: run a JavaScript orchestration script over ctx.workflowEngine |
| `@forwork/workflow-ptc` | yes | Workflow orchestration in the shared sandboxed Node PTC runtime |

## workspace

| Package | Config | Description |
|---|---|---|
| `@forwork/workspace` | no | Workspace entity registry (ctx.workspaceRegistry): durable workspace records with validated session attachment over the domain data form for the DeepSeek Harness |
