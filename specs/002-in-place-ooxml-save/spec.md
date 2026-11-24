# Feature Specification: In-Place OOXML Save

**Feature Branch**: `protective-talk` (feature directory `002-in-place-ooxml-save`; this checkout has no branch-creation hook installed)

**Created**: 2026-11-18

**Status**: Draft

**Input**: User description: "原地保留式 OOXML 写入：.docx/.xlsx 保存时只重写被编辑的部分，未触碰字节按原样保留，使"拒绝打开"的构造集趋近于零"

## Clarifications

### Session 2026-11-18

- Q: When an edit forces a part to change (for example `word/document.xml` after a paragraph edit), must every byte outside the edited XML span also stay identical, or is byte identity only required for parts the edit does not touch at all? → A: Span-level splice — only the edited unit's XML span is rewritten; every byte outside it is identical.
- Q: Should macro-enabled `.docm` and `.xlsm` files become editable under the same preservation rules, or stay outside the accepted editable formats? → A: Include `.docm` and `.xlsm` as editable; macro parts are carried untouched and the macro-enabled content type is preserved.
- Q: Which independent application should prove that a saved file opens without a repair prompt? → A: Headless LibreOffice plus structural validation on every change, and a fixed fixture subset opened in desktop Word and Excel before release. Superseded by Session 2026-11-24: the release host has no desktop Word or Excel, so the subset is validated in the system LibreOffice on Windows, and no no-repair claim is made.
- Q: Where should the per-save record of which package entries changed live? → A: A durable session save fact carrying exact counts plus changed, added, and removed entry names, capped for large saves.
- Q: When the edited unit's own XML carries a property or construct the editable model does not represent, should the save refuse or carry it through the splice? → A: Carry it verbatim when it is independent of the edited content; refuse only when its semantics or anchors depend on the edited content.

### Session 2026-11-24

- Q: Which hosts and which independent application supply the release validation evidence for this feature? → A: Windows only. The system LibreOffice installation on the host (build 26.8.3.1) is the independent engine for the import/render/read-back matrix, and macOS is out of scope. Desktop Word and Excel are not installed on this host, so their no-repair check cannot be run; it is recorded as a residual risk disclosed in the PR and release notes, never claimed as passed. LibreOffice can silently repair a package and reports no repair diagnostics, so no matrix result is described as a no-repair proof.

## User Scenarios & Testing *(mandatory)*

### User Story 1 - Save one edited region without disturbing the rest of the package (Priority: P1)

A user opens a real-world `.docx` or `.xlsx` that carries content the editor does not model: a custom theme, a chart, a pivot table, tracked changes in another clause, form controls, an embedded object, a macro. They change one paragraph or one cell and save. The saved file differs from the source only in the edited unit spans; every other byte in the package is identical, and the file opens in the target Office application without a repair prompt.

**Why this priority**: This is the feature. Today a save rebuilds the package from the editable model, so anything outside that model is either silently dropped or refused outright before the user can save.

**Independent Test**: Take the documented construct corpus (at least one fixture per construct class), perform one targeted edit in each, save, and compare source and output at package level: only the edited unit spans may differ, every entry and every byte outside those spans must be identical, and an independent OOXML consumer must open the output without repair.

**Acceptance Scenarios**:

1. **Given** a `.docx` whose package contains a chart part, a theme, custom style definitions, and a tracked change in an untouched paragraph, **When** the user edits one other paragraph and saves, **Then** the chart, theme, style definitions, tracked change, media, relationship and content-type parts are byte-identical to the source, and the document opens without repair.
2. **Given** an `.xlsx` containing a pivot table, an Excel table, defined names, a calc chain, a custom theme, and a drawing, **When** the user changes one cell and saves, **Then** all of those entries are byte-identical and the workbook opens without repair.
3. **Given** any supported file, **When** the user saves without making an edit, **Then** the saved file is byte-identical to the source.
4. **Given** a construct whose meaning or anchors depend on the content being edited, **When** the user saves, **Then** the save is refused with that construct named and the source file is unchanged.

---

### User Story 2 - Complex files stop being refused (Priority: P1)

Documents that today cannot be saved at all, because the source contains a macro, a chart, a drawing or shape, a comment, an embedded object, a pivot table, or a tracked change anywhere in the file, become saveable as long as the edit does not intersect that construct. The unsupported-feature surface no longer describes the whole file; it describes only the constructs a specific edit would damage.

**Why this priority**: The whole-file refusal is the user-visible cost of not preserving bytes; removing it is the measurable outcome the feature exists for.

**Independent Test**: For every construct class in the refusal corpus, open, make a non-intersecting edit, save, reopen, and confirm both that the save was accepted and that the construct survived byte-for-byte.

**Acceptance Scenarios**:

1. **Given** a macro-bearing `.docm`, **When** the user edits a paragraph that does not contain or anchor the macro and saves, **Then** the save succeeds, the VBA project part is byte-identical, the file stays macro-enabled, and the macro is still present when reopened.
2. **Given** an `.xlsx` with a pivot table and a chart on another sheet, **When** the user edits an ordinary cell and saves, **Then** both survive byte-for-byte.
3. **Given** a `.docx` with a tracked change inside the edited paragraph, **When** the user saves, **Then** the save is refused and the tracked change is named as the reason.
4. **Given** a file whose unsupported construct is untouched, **When** the editor renders the save control, **Then** it is enabled rather than disabled because of that construct.

---

### User Story 3 - Every save is attributable and auditable at entry level (Priority: P2)

A reviewer or support engineer can see exactly which package entries a save changed and why. When a saved file fails to open, the difference between source and output is enough to identify the offending entry; no hidden global rewrite is involved.

**Why this priority**: Auditability turns "we only rewrite the edited part" into a verifiable claim and gives a repair path a concrete artifact.

**Independent Test**: Perform a sequence of saves, inspect the per-save disposition report, and confirm that the reported changed entries exactly equal the byte-level difference between source and output.

**Acceptance Scenarios**:

1. **Given** a completed save, **When** the product reports the result, **Then** it can list every entry carried verbatim and every entry spliced, added, or removed.
2. **Given** a save that changed exactly one span, **When** the report is compared with a byte-level difference, **Then** the two agree exactly.
3. **Given** a refused save, **When** it fails, **Then** no entry was written and the source digest is unchanged.

---

### Edge Cases

- The source stores an entry under an unusual name, uses a non-default compression method, or stores it uncompressed: entry identity follows the container record, not an assumed canonical name.
- A `.docm` or `.xlsm` is saved: the macro-enabled content type and the VBA project entry survive, and the output is never converted into the macro-free variant of the format.
- Two entries whose names differ only by case appear in the source: the save fails loudly instead of silently merging or rewriting either entry.
- The edit adds content that requires a new entry (a copied image, a new hyperlink, a new header): only the new entry and the entries that must reference it may change; existing entries keep their bytes.
- The edit removes content that owned an entry (the only image in a header): the removal set is explicit and limited to the orphaned entry, its relationship, and its content-type declaration.
- A saved file is opened and saved again with no further edit: the second output is byte-identical to the first, not merely equivalent.
- The file changed on disk between open and save: the save is refused as a conflict and the source is untouched.
- The file is password-protected, encrypted, or not a valid OOXML package: the product fails loudly and writes nothing.
- The file exceeds the configured size or cell bounds: the product fails loudly with a bounded message and writes nothing.
- A construct whose meaning or anchors depend on the edited content intersects the edited range: the save is refused, names the construct, and leaves the file byte-identical.
- A construct is only partially inside the edited range (a tracked change spanning the edited paragraph boundary): the save is refused instead of partially rewriting it.
- An edited cell or paragraph carries an unmapped property that is independent of its text (fit-to-text, a width variant, a row-level flag): the property is re-emitted inside the spliced span and the save succeeds.
- A changed entry cannot be re-encoded without reinterpretation: the save is refused instead of emitting a best-effort entry.
- The source span of an edited unit cannot be located unambiguously (duplicate markup, repeated identical units, non-standard namespace prefixes): the save is refused instead of splicing at a guessed offset.
- The supplied relationship graph resolves case-insensitively or through relative targets: preservation neither duplicates nor drops the referenced entry.
- A large package within bounds contains thousands of entries: preservation stays bounded and does not load unrelated entries into the editor model.
- A save changes more entries than the report cap: the durable fact keeps exact counts and a deterministic prefix of names, and the save succeeds regardless of report size.
- The same content is saved from two independently opened instances: the second save is rejected by the version guard.

## Requirements *(mandatory)*

### Functional Requirements

**Preservation**

- **FR-001**: On save, every source package entry whose content the edit does not change MUST be reproduced with byte-identical content.
- **FR-002**: A save that applies no edit MUST produce a file byte-identical to the source, including container-level entry order, compression, timestamps, and metadata.
- **FR-003**: Only entries whose content the edit changes, plus entries the edit necessarily adds, removes, or re-references, MAY differ from the source. Inside an entry the edit changes, only the edited unit's XML span MAY differ; every byte outside the union of edited spans MUST be identical.
- **FR-004**: Unmodeled constructs MUST be carried verbatim rather than interpreted: macros, charts, drawings and shapes, comments, embedded objects, pivot tables and caches, tracked changes, Excel tables, defined names, calc chains, custom XML, external links, threaded comments, themes, font tables, web settings, document properties, fields, content controls, bookmarks, and package entries not otherwise listed.
- **FR-005**: A carried-over entry MUST keep its container metadata (name, order, compression method, CRC, timestamps, extra fields) and MUST NOT be re-encoded.
- **FR-006**: When the edit requires a new entry or a removal, the affected relationship graph and content-type declarations MUST be updated in the same save, and no unrelated declaration MAY change.

**Refusal**

- **FR-007**: A save MUST NOT be refused solely because the source contains a construct outside the editable model.
- **FR-008**: A save MAY be refused only when the edit intersects a construct whose meaning or anchors depend on the content being changed, or when preservation of a changed entry cannot be guaranteed. An unmapped property inside the edited span that is independent of the edited content MUST be carried verbatim instead of causing a refusal. The refusal MUST name the construct, and the save MUST leave the source byte-identical.
- **FR-009**: The product MUST NOT silently drop, normalize, or reinterpret content it does not explicitly change. Inside an edited span, unmapped properties independent of the edited content are re-emitted verbatim; only the content the edit actually changes is replaced.
- **FR-010**: Every difference between source and saved file MUST be attributable to the edit or to the minimal reference updates in FR-006.

**Validity**

- **FR-011**: The saved package MUST remain structurally valid OOXML: every part referenced by a relationship exists, every part has a content-type declaration, and no declaration points at a missing part.
- **FR-012**: Every saved corpus fixture MUST pass structural package validation on every change. The fixed release subset MUST import, render, and read back the saved edit in the system LibreOffice installation on Windows, with the engine build and host recorded for each fixture. LibreOffice can silently repair a package and reports no repair diagnostics, so these results MUST NOT be described as a no-repair proof; desktop Word/Excel no-repair validation is not a release gate and MUST be disclosed as an unverified residual risk.

**Contract and persistence**

- **FR-013**: The save operation MUST use the source package (its bytes or an equivalent entry inventory) as the base for the output; preservation MUST NOT depend on reconstructing source bytes from the editable model.
- **FR-014**: Preservation state MUST survive the full read → wire → editor → save path, including editor snapshot serialization; a construct preserved on read MUST still be preserved by a save that follows an editor round trip.
- **FR-015**: Version guards MUST remain: a save against a stale source version is refused, and no preservation work is written.
- **FR-016**: Saving MUST remain deterministic: the same source and the same edit produce identical output.
- **FR-017**: The product MUST record, per save, which entries were carried verbatim and which were spliced, added, or removed, in a durable save fact that stays readable after restart. The fact carries exact counts plus the affected entry names, capped at 256 names per list for large saves, and MUST NOT fail or block the save because the report is large.
- **FR-018**: Package bounds (entry count, expanded size, file size, cell count, elapsed time) MUST continue to apply, with a named bounded failure and no write on failure.

**User and model surfaces**

- **FR-019**: The refusal surface MUST distinguish "the file contains X, so saving is blocked" from "this edit touches X, so it cannot be saved"; only the second form may block a save.
- **FR-020**: Preserved-but-uneditable content MUST be identified to the user with a visible, localized marker for every preserved construct class in the corpus, without implying that the whole file is unsavable; the save control MUST stay enabled.
- **FR-021**: Agent-initiated saves MUST follow the same preservation and refusal rules, with the same attribution and version guards as user saves.

**Format scope**

- **FR-022**: The product MUST accept `.docm` and `.xlsm` as editable under the same preservation and refusal rules as `.docx` and `.xlsx`, and MUST preserve the content type that distinguishes a macro-enabled container; a save MUST NOT silently produce a macro-free file or change the format.

### Key Entities

- **Source Package**: the file as opened — container bytes plus the entry inventory (name, compressed and uncompressed bytes, CRC, method, order, timestamps, extra fields) and the relationship/content-type graph.
- **Editable Part Set**: the package entries that map into the editable document or workbook model; an edit may splice spans only inside these entries, and every other entry is carried verbatim.
- **Edited Range**: the content locations the edit actually changes, expressed in the editable model and resolved to the exact XML source span of each edited unit (paragraph, table cell, or sheet cell); every byte outside the union of those spans is preserved.
- **Construct Intersection**: the set of unmodeled constructs overlapping the edited range whose meaning or anchors depend on the edited content; empty for every save that must succeed. Independent unmapped properties inside the span are carried, not counted as an intersection.
- **Part Disposition**: one record per source entry describing it as carried verbatim, spliced (with the edited spans named), added, or removed, with the reason that justifies any non-verbatim disposition.
- **Preservation Evidence**: the per-save disposition report, persisted with the durable save fact and paired with the resulting byte-level difference from the source, used to prove FR-010 after a restart and to diagnose a file that fails to open.
- **Conflicting Constraint**: a construct whose meaning or anchors depend on the edited content; the only remaining legitimate reason to refuse a save.

## Success Criteria *(mandatory)*

### Measurable Outcomes

- **SC-001**: In a corpus of at least 30 fixtures covering the construct classes in FR-004, 100% of no-edit saves produce a byte-identical file.
- **SC-002**: Across the same corpus, 100% of saves leave every untouched entry byte-identical and every byte outside the edited unit spans exact, including unmapped properties and constructs in untouched content of an edited part.
- **SC-003**: Across the same corpus, 100% of saves pass structural package validation. In the Windows system-LibreOffice matrix, 100% of the fixed release subset imports, renders, and reads back the edit as intended with no engine error, and each result records the engine build and host. Because LibreOffice can repair silently, no result is reported as a no-repair proof; desktop Word/Excel no-repair behavior is a disclosed residual risk, not a success criterion.
- **SC-004**: Across the same corpus, 100% of non-intersecting edits save successfully; saves refused for a construct located outside the edited range number zero.
- **SC-005**: Zero silent losses: for every save in the corpus, the byte-level difference contains no entry the disposition report does not account for.
- **SC-006**: 100% of saves intersecting a coupled construct are refused with that construct named, and the source digest is unchanged.
- **SC-007**: A user can open, edit one region of a package of at most 5 MB and 10,000 cells, and save in under 30 seconds on the reference desktop.
- **SC-008**: Saving the same source with the same edit twice produces identical bytes in 100% of tested cases.
- **SC-009**: Every save yields a disposition report that, in 100% of tested cases, exactly matches an independent byte-level difference, and the report remains readable from the session history after a restart.
- **SC-010**: No document bytes leave the device during preservation, saving, or reporting (the offline guarantee is retained).
- **SC-011**: Across the macro-enabled fixtures, 100% of saves leave the VBA project byte-identical and keep the macro-enabled content type; zero saves silently convert a `.docm`/`.xlsm` to a macro-free file.
- **SC-012**: Across the fixtures where an edited cell or paragraph carries only unmapped properties independent of its text, 100% of saves succeed and re-emit those properties; the number refused for such independent properties is zero.
- **SC-013**: For every preserved construct class in the corpus, the editor shows a visible localized preserved-not-editable marker and keeps the save control enabled in 100% of tested cases; no preserved construct disables saving.

## Assumptions

- "In-place" means the saved file is assembled from the source package so that only edited entries change. The visible file is still published through the existing atomic, version-guarded replace, not patched while another process may read it.
- Preservation is span-level, not merely entry-level: the edited unit's XML is spliced into the source text, so every byte outside the edited spans is retained even when its part changes. Edited spans are located in the original XML, never re-derived by re-serializing the part and diffing it.
- A part that must be wholly re-encoded to apply an edit (for example a style table when a style changes) is a preservation failure for that part unless the edit's spans can be applied surgically; the save refuses rather than regenerating the part wholesale.
- Scope is `.docx`, `.docm`, `.xlsx`, and `.xlsm`. `.pptx` remains preview-only; legacy `.doc/.xls/.ppt` conversion and legacy binary write-back stay out of scope. The constitution records the widened v1 editable set and the package-level write stack in amendment 3.2.0, and the plan cites that amendment.
- Carried-over constructs remain uneditable: preservation removes the loss and refusal risk, it does not add editor support. Editing a macro or a pivot table is out of scope.
- Macros are carried as opaque bytes and are never executed.
- "The refusal set approaches zero" means constructs outside the edited range never refuse, and independent unmapped properties inside it are carried; the residual set is exactly the coupled constructs intersected by an edit, plus genuine preservation failures, not a fixed file-level list.
- A no-op save is expected to be byte-identical; a save that changes only an entry timestamp is a defect.
- The existing editable mapping (text, formatting, tables, cells, styles, formulas, conditional formatting, data validation, hyperlinks, images) is retained; this feature changes how bytes are assembled, not what the editor can edit.
- The existing file-size, cell-count, conversion-time, and package-entry bounds remain; preservation must stay bounded and fail loudly rather than degrade silently.
- A document remains a first-class session entity with version guards; conflict, draft, and attribution behavior is unchanged. The durable save fact gains the disposition summary, so the session-event schema changes and both SDK projections must acknowledge it.
- Correctness is judged by byte-level comparison, structural package validation, and the system LibreOffice on Windows for the fixed release subset, not by the product's own reader. Desktop Word/Excel no-repair behavior is unverified on this host and disclosed as a residual risk.
- A property is "independent" of the edited content when its meaning does not reference the changed text or the identity of the changed node (for example a cell's fit-to-text or a table cell's width rule). A construct is "coupled" when the edit would change its meaning or invalidate its anchor (a tracked change wrapping the edited text, a comment anchor inside the edited range, a content control whose text is being changed).
- A save request with no pending edit still completes as a save: it records the save fact and advances the version, and its output is byte-identical to the source.
- The documented corpus is the evidence base: a construct absent from the corpus may still be preserved, but the claim cannot be made without a fixture.
