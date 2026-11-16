<!--
Sync Impact Report
==================
Version change: 1.1.0 -> 2.0.0 (MAJOR)
Rationale: Principle VIII is redefined from "the fork must stay mergeable" (wholesale
merge-forward) to "selective upstream follow" with an executable decision table. The obligation
changes, not only its wording; the other principles are unchanged.

Modified principles:
- VIII. Fork Stewardship -> VIII. Selective Upstream Follow

Added sections: none
Removed sections: none
Follow-up TODOs: none
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

### VIII. Selective Upstream Follow
The fork follows its upstream base selectively. Routine synchronization is not a wholesale merge: take the changes that keep security and framework correctness current, and deliberately drop changes that do not serve this product. Every sync records what was taken, what was dropped, and why in a dated Agent Note; `docs/development.md#upstream-sync` owns the operational procedure.

Decision order: a security fix is taken; a change to a file this fork has diverged on is skipped, or taken and re-applied on top of the local change; DeepSeek product, brand, account, telemetry, API adapter, and default search are dropped; an AGPL, cloud-account, or paid-SDK dependency is dropped; a breaking `SESSION_FORMAT_VERSION` change or one requiring migration is deferred and evaluated with an upgrade guide.

| Change | Action | Reason |
| --- | --- | --- |
| Security fix | Take | Risk otherwise accumulates |
| Core framework bug fix (Cordis, agent loop, session, fs/sandbox, shell, terminal, client rendering, Typert, dockkit, remotes) that does not touch a diverged file | Take | Direct benefit |
| Dependency or lockfile security/compatibility upgrade | Take, then rerun the full relevant checks | Prevent breakage |
| Fix on a file this fork has diverged on | Skip by default; if critical, take and re-apply the local change | Conflict cost |
| DeepSeek product, brand, account, telemetry, `web-deepseek-api`, default search | Drop | Unrelated and conflicts with Principle I |
| Experimental or optional package this product does not ship (liveshare, computer-use, ...) | Drop unless adopted | No product value |
| Dependency imposing AGPL, a cloud account, or a paid SDK | Drop | Principles II and VII |
| Breaking `SESSION_FORMAT_VERSION` change or required migration | Defer; evaluate each with an upgrade guide | Highest migration risk |

Product changes still extend shared seams or add packages rather than rewriting upstream modules in place when an additive option exists. A cherry-pick that touches a diverged file uses a preserved-work recovery copy, a clean tree, and no shared stash, and records the conflict resolution.

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
5. Upstream synchronization cherry-picks per Principle VIII; a cherry-pick that touches a diverged file uses a verified recovery copy and a clean tree (no shared stash) and re-applies the local change.

## Governance

This constitution supersedes ad-hoc preference. Amendments require a recorded rationale, the affected requirement, and a migration note. Every change review verifies the principles above; a violation MUST be justified in the review or removed. Versioning follows semantic versioning: MAJOR for principle removal or incompatible redefinition, MINOR for a new principle or material expansion, PATCH for clarification. Ratified and amended dates use ISO `YYYY-MM-DD`.

**Version**: 2.0.0 | **Ratified**: 2026-11-13 | **Last Amended**: 2026-11-16
