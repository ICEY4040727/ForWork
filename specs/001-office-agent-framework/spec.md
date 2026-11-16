# Feature Specification: Office Agent Framework

**Feature Branch**: `001-office-agent-framework`

**Created**: 2026-11-13

**Status**: Implemented

**Input**: User description: "把作为智能体框架的 harness 改造成智能体操作办公文档的框架，在侧栏挂载完整的 Word、Excel、编辑器（这里指 Word/Excel 编辑器本身）；后续要求：不是可逆插件、可改 harness 核心、密钥自动配置、借用框架而非插件作者、独立签名、Windows/macOS 桌面。"

## Clarifications

### Session 2026-11-14

- Q: Should the editable snapshot use a project-owned neutral model or the editor core's native snapshot model? → A: Use the editor core's native snapshot model directly (workbook and document snapshot types); persist attributable change operations, not raw snapshots.
- Q: Should the editor and the agent act on one live document instance or on separate instances? → A: One live instance: the editor owns the unit and agent document tools dispatch operations to it through a renderer-side command bridge, sharing one undo stack and state.
- Q: How do agent document operations reach the live instance? → A: Host-owned tools dispatch correlated commands over the gateway; the renderer executes them against the live unit and returns a correlated result.
- Q: What is the source of truth for document undo and redo? → A: The editor's native undo stack, shared by user and agent commands; change records retain inverses for audit and recovery, not for runtime undo.
- Q: What is the default save behavior? → A: User edits save explicitly, with periodic local drafts for crash recovery; an accepted agent tool call persists immediately. The original is never written silently.
- Q: How are browser-only editor dependencies tested? → A: The editor adapter is mocked in component tests, and real behavior is verified in a browser smoke.
- Q: How is the change-operation vocabulary extended? → A: Operation kinds are added additively (rows, columns, formulas, styles) with a same-version persistence acknowledgement.
- Q: How are legacy `.doc/.xls/.ppt` files opened? → A: Convert into the workspace and open the converted copy in a new tab with a visible notice; the source stays unchanged.
- Q: When does an agent document edit become durable? → A: Each accepted agent tool call applies to the live unit and persists with one immediate version-guarded save; if unsaved user edits exist, the call is refused (`pending-edits`) so attribution stays exact.
- Q: What does "the same file" mean for a legacy source? → A: The file currently open in the editor; for `.doc/.xls/.ppt` that is the converted copy, and the legacy source is never modified.

## User Scenarios & Testing *(mandatory)*

### User Story 1 - Edit a Word or Excel file in the app (Priority: P1)

A user opens a local `.xlsx` or `.docx` file from the product's sidebar. A full editor appears in the sidebar. The user changes a cell value, or edits a paragraph, and saves. Reopening the file preserves everything they did not touch: styles, formulas, images, comments, and structure.

**Why this priority**: Without reliable in-app editing there is no product; it is the smallest independently demonstrable slice.

**Independent Test**: Open a fixture set of `.xlsx`/`.docx` files, perform one targeted change in each, save, reopen, and diff against the expected result; untouched regions must be byte/feature equivalent.

**Acceptance Scenarios**:

1. **Given** a `.xlsx` with formulas and styles, **When** the user changes one cell and saves, **Then** only that cell changes and all formulas/styles remain.
2. **Given** a `.docx` with headings, tables, and images, **When** the user edits one paragraph and saves, **Then** all other paragraphs, tables, and images are preserved.
3. **Given** an unsupported operation (for example editing a macro), **When** the user attempts it, **Then** the product refuses and explains the limitation instead of saving a lossy file.

### User Story 2 - Agent operates the same document (Priority: P1)

The user asks the agent to change something in the currently open document ("把第三个表格的 B 列改成货币格式并算总和"). The agent reads the document structure, performs the edit, and the user sees the result in the open editor without manually reloading.

**Why this priority**: This is the defining capability of an "office agent framework"; it is also independently testable.

**Independent Test**: With a document open, run an agent instruction that changes a cell or paragraph; verify the editor reflects the change and the saved file contains it.

**Acceptance Scenarios**:

1. **Given** an open spreadsheet, **When** the agent sets a cell value, **Then** the editor shows the new value and the file on disk contains it.
2. **Given** an open document, **When** the agent replaces a paragraph's text, **Then** the change appears in the editor and untouched content is preserved.
3. **Given** a concurrent external change, **When** the agent saves, **Then** the product detects the version conflict and reports it instead of silently overwriting.

### User Story 3 - Documents persist as session entities (Priority: P2)

The user opens several documents across a session. The agent can refer to them across turns, and the product restores the open-document set after a restart.

**Why this priority**: Turns a one-off editor into a durable workspace; required for cross-turn references but not for the first useful slice.

**Independent Test**: Open two documents, restart the app, and confirm both reopen; ask the agent a follow-up question that references one of them and confirm it resolves.

**Acceptance Scenarios**:

1. **Given** two open documents, **When** the app restarts, **Then** both reopen in the same tab set.
2. **Given** a prior turn that opened a document, **When** the user refers to "that table" in a later turn, **Then** the agent resolves it to the opened document.

### Edge Cases

- A file is deleted or replaced on disk while open: the editor must report the external change and offer reload, not corrupt the file.
- Two writers (agent and user) save at nearly the same time: the version guard must reject the stale write and report a conflict.
- A very large workbook or document exceeds the conversion bounds: the product must fail with a bounded, actionable message.
- A legacy `.doc/.xls/.ppt` is opened: it is converted to a modern format; the original is not modified unless the user explicitly saves a converted copy.
- A document contains a feature the engine cannot edit (macros, tracked changes, complex shapes): the product names the unsupported feature and refuses the lossy save.
- An open file is renamed or moved outside the app: the document keeps its identity and rebinds to the new location, or reports the file missing without silently closing it.
- The app quits or crashes with unsaved edits: the user is warned before closing, and an interrupted session can recover its draft rather than losing work silently.
- Conversion of a corrupt or password-protected file fails: the product reports the failure and leaves the source file untouched.
- The agent attempts a write outside the document's permitted roots: the write is refused with an explanation.
- The user edits one document while the agent edits a different document: both proceed independently without cross-document conflicts.

## Requirements *(mandatory)*

### Functional Requirements

- **FR-001**: The product MUST open `.xlsx` and `.docx` files in an editable sidebar surface.
- **FR-002**: The product MUST save edits back to the file currently open in the editor (for a legacy source, the converted copy) without modifying content the user did not change.
- **FR-003**: The product MUST convert legacy `.doc/.xls/.ppt` to a modern editable format on open without altering the source.
- **FR-004**: The product MUST expose document structure and content to the agent as structured operations, not raw bytes.
- **FR-005**: Agent edits and user edits MUST converge on one document state, with version-guarded writes and a visible conflict report.
- **FR-006**: The product MUST record document open/edit/rebind/missing/close as durable session facts so the agent can reference documents across turns.
- **FR-007**: The product MUST restore the open-document set after restart.
- **FR-008**: The product MUST refuse and explain any save that would drop unsupported content.
- **FR-009**: The product MUST be delivered as the ForWork desktop application under the `@forwork/*` npm scope, the `forwork-` artifact prefix, and the `forwork` CLI, with no remaining DeepSeek brand, telemetry, default model adapter, or upstream npm scope in product surfaces. Legally-required provenance notices, protocol-internal `dsh.*`/`DSH_*` tokens, and the external LibreOffice conversion dependency are the only accepted references, each enforced by `verify-forwork-identity`.
- **FR-010**: The product MUST operate fully offline for editing and saving, with no required external document server or cloud account.
- **FR-011**: `.pptx` MUST be viewable but not editable in v1.
- **FR-012**: The product MUST preserve MIT and Apache-2.0 notices for the harness base and the editor core.
- **FR-013**: Document content and edits MUST stay on the user's device; the product MUST NOT upload documents or edit history, and telemetry MUST be off by default and opt-in only.
- **FR-014**: Every change MUST be attributable to its actor, and the product MUST provide document-level undo and redo, including reversal of agent-initiated edits. Undo and redo MUST be available from the editor's standard undo surface and MUST apply to agent and user changes alike.
- **FR-015**: A document's identity MUST survive the file being renamed or moved; restore MUST rebind to the new location or report the document as missing without data loss.
- **FR-016**: The product MUST NOT discard unsaved edits silently on close, crash, or external file change; it MUST warn or recover.
- **FR-017**: Agent initiated writes MUST stay within the filesystem roots permitted for the session; out-of-scope writes MUST be refused.
- **FR-018**: User-supplied model credentials MUST be stored in the operating system's secure store and MUST NOT be committed to the repository or written to logs.
- **FR-019**: The product MUST remain updatable from its upstream framework base without losing local product changes; upstream updates MUST preserve local work and record conflict resolution.
- **FR-020**: Agent document operations MUST apply to the same live document instance the editor displays, through one command path; an operation on a document that is not open in a live instance MUST be refused rather than applied to a detached copy. (Convergence and version-guarded writes are FR-005.)

### Key Entities *(include if feature involves data)*

- **Document**: an opened office file; identity is a stable document id scoped to the authorizing session, with the current absolute path as a rebindable locator; carries a content version and a rendered/serialized state.
- **Document Session Fact**: a logged open/edit/rebind/missing/close fact tied to a document, turn, and content version.
- **Change Record**: an attributable, reversible edit applied by the user or the agent, tied to a document version.
- **Document Draft**: a local, version-bound snapshot of unsaved edits used for crash recovery; it never overwrites the open file and is cleared after a successful save.
- **Unit Snapshot**: the editable state of a workbook or text document, held directly in the editor core's native workbook/document snapshot model so the editor and the agent share one model; the session log persists attributable change operations rather than raw snapshots.
- **Conversion Record**: the mapping result between an OOXML file and a unit snapshot, including any unsupported features detected.

## Success Criteria *(mandatory)*

### Measurable Outcomes

- **SC-001**: A user can open, edit one region, and save an `.xlsx` or `.docx` fixture of at most 5 MB and 10,000 cells in under 30 seconds.
- **SC-002**: Across the fixture set, 100% of untouched regions are preserved on save (no style, formula, image, or comment loss).
- **SC-003**: An agent edit is visible in the open editor in under 2 seconds without a manual reload.
- **SC-004**: Every attempted save that would drop unsupported content is refused with a named reason; zero silent lossy saves in the test suite.
- **SC-005**: The open-document set is fully restored after restart in 100% of tested sessions.
- **SC-006**: A brand and telemetry scan of the packaged product reports zero DeepSeek-referencing product surfaces outside the declared provenance/protocol/external-dependency allowlist.
- **SC-007**: Signed installers launch on Windows and macOS with the product's own identity and no unsigned-publisher warning.
- **SC-008**: No document content or edit history leaves the device during editing, saving, or agent operation, verified by an egress check of tested flows.
- **SC-009**: Renaming or moving an open file keeps it open in 100% of tested cases, or reports it missing without losing its identity or data.
- **SC-010**: In tested abnormal-exit scenarios, 0 unsaved edits are lost without a warning or a recoverable draft.
- **SC-011**: 100% of agent write attempts outside the permitted roots are refused in the test suite.
- **SC-012**: An agent-initiated edit can be undone at the document level in 100% of tested cases.

## Assumptions

- Documents are local files opened through the product's workspace; cloud storage is out of scope for v1.
- The user supplies their own model credentials for the agent; the product does not ship a bundled model key, and credentials live in the OS secure store.
- "编辑器" means the Word and Excel editors themselves, not a separate code editor.
- Real-time multi-user collaboration (beyond the single local user plus agent) is out of scope for v1.
- PowerPoint editing, legacy in-place binary write-back, and automatic document-server provisioning are out of scope for v1.
- Creating blank documents from scratch is out of scope for v1; the product operates on existing files.
- Telemetry is opt-in; the default build sends nothing.
- The product is a fork that keeps tracking its upstream framework base; upstream updates are applied as merges that preserve local work.
- The editor core's native snapshot types are the interchange format for the unit snapshot; no project-owned neutral model is maintained.
- Agent document operations require the document to be open in a live editor instance; headless runs without an editor refuse document operations instead of editing a detached copy.
