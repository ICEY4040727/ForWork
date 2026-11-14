<!--
Sync Impact Report
==================
Version change: 1.0.0 -> 1.1.0 (MINOR)
Rationale: no principle removed; one principle added and four materially expanded. The core intent
of all seven original principles is preserved. If the V rename is judged a redefinition rather
than an expansion, treat this as 2.0.0.

Modified principles:
- II. Offline, Self-Contained Desktop First -> II. Offline, Local-First Desktop
  (adds data ownership/privacy; scopes "offline" to editing and saving; names paid cloud SDKs)
- IV. Lossless by Default; Never Silently Drop Content -> IV. Lossless, Non-Destructive Editing
  (adds agent write authority, attribution, and conflict handling)
- V. Additive Increments; Rename Last -> V. Incremental, Reversible Delivery
  (generalizes a one-time sequencing rule into durable delivery governance)
- VII. License and Attribution Hygiene -> VII. License and Dependency Hygiene
  (adds a dependency policy and license gate)

Added sections:
- VIII. Fork Stewardship

Removed sections: none

Follow-up TODOs: none (ratification date is known).
-->

# Office Agent Harness Constitution

## Core Principles

### I. Independent Identity (fork, not plugin)
The product is a fork of the DeepSeek Harness used as a base framework, not a plugin authored for DeepSeek. Every package scope, brand mark, CLI name, model default, telemetry path, and distribution signature belongs to this product. A `deepseek` reference that remains reachable in a shipped surface is a defect, detected by an executed scan rather than by inspection. Provenance notices for the MIT base and Apache-2.0 editor remain, per Principle VII.

### II. Offline, Local-First Desktop
Targets are Windows and macOS Electron desktop, independently signed. Document editing and saving MUST run locally with no required Linux server, cloud account, container runtime, or network round-trip. The model call is the only expected network dependency. User documents remain local unless the user explicitly exports or shares them; telemetry is off by default. AGPL engines, cloud-account document services, and paid editor SDK modules are excluded from v1 unless an amendment records the trade-off and its license.

### III. Documents Are First-Class Session Entities
An opened document is durable session state, not an ad-hoc tool argument. Its open/edit/close facts are logged so the model request is reconstructable from the log (model-visible ⟺ logged), and the agent can reference a document across turns. The session log is the source of truth; the UI is a projection of it, and the open-document set MUST survive restart.

### IV. Lossless, Non-Destructive Editing
An edit preserves every untouched byte, style, and feature it does not explicitly change. Where the engine cannot represent an operation, the operation fails loudly and names the unsupported feature; it never rewrites a file in a lossy way without saying so. Every write is attributable to its actor (user or agent), version-guarded, and rejected on a stale version with a visible conflict; no path silently overwrites another actor's change. Agent-initiated edits that alter user content MUST be observable and reversible at the document level.

### V. Incremental, Reversible Delivery
New capability lands as additive packages and new session events before any destructive operation. A destructive step (repository-wide rename, scope change, telemetry removal, signing-identity change) MUST wait until the additive capability is working and verified, so it can be reviewed, reverted, and landed in pieces. "Rename last" is the current instance of this rule, not its permanent limit.

### VI. Evidence Over Claims
A change is complete only after the narrowest relevant check passes: typecheck, lint, unit or behavior tests, coverage, session snapshot, or a real packaged smoke. A passing build of one path does not prove another. Report only checks actually run, and never present an unverified path as working.

### VII. License and Dependency Hygiene
The harness base is MIT; the Univer core is Apache-2.0. Forking and commercial distribution retain both notices. Third-party conversion libraries are used under compatible licenses. A dependency that imposes copyleft on the product, or that requires a cloud account for core editing, MUST NOT be a hard dependency without an explicit amendment. Dependency additions are checked for license compatibility before adoption, and the shipped package retains the required notices.

### VIII. Fork Stewardship
The fork MUST stay mergeable with its upstream base. Product changes extend shared seams or add packages rather than rewriting upstream modules in place when an additive option exists. Upstream merges are performed with a preserved-work recovery copy, a clean merge tree, and recorded conflict intent, so borrowed framework fixes keep arriving. Divergence that would make a future upstream merge conflict-heavy requires a recorded rationale and a planned reconciliation.

## Technology Constraints

- Node >= 24.9, TypeScript strict, ESM everywhere, pnpm workspaces.
- Editor core: Univer open-source (Apache-2.0) Sheets/Docs presets, embedded in React; no iframe and no server.
- Conversion: a self-owned Exchange layer (ExcelJS, SheetJS CE, a docx library, LibreOffice kit) mapping OOXML to and from Univer unit snapshots.
- v1 edits `.docx` and `.xlsx`; `.pptx` is read-only. Legacy `.doc/.xls/.ppt` are converted on open.
- Persistence writes through the filesystem seam with version guards so observers converge.

## Development Workflow

1. Constitution -> specification -> plan -> tasks -> implementation.
2. Each phase is independently verifiable and independently deliverable.
3. Tests and snapshots accompany behavior; a model-visible or user-visible change without coverage is incomplete.
4. Destructive repository operations (rename, scope, signature) are last and gated on a green additive baseline.
5. Upstream base merges run on a clean tree with a verified recovery copy and no shared stash; conflicts preserve both sides' intent.

## Governance

This constitution supersedes ad-hoc preference. Amendments require a recorded rationale, the affected requirement, and a migration note. Every change review verifies the principles above; a violation MUST be justified in the review or removed. Versioning follows semantic versioning: MAJOR for principle removal or incompatible redefinition, MINOR for a new principle or material expansion, PATCH for clarification. Ratified and amended dates use ISO `YYYY-MM-DD`.

**Version**: 1.1.0 | **Ratified**: 2026-11-13 | **Last Amended**: 2026-11-13
