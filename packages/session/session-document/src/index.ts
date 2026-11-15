/**
 * Host document session primitives shared by the `document` projection, the
 * open-document prompt section, restore, and the document tools: stable
 * identity bindings across rename and missing paths, version-guarded
 * attributable change records with their inverses, the durable `document/*`
 * fact writers, the ordered restore set, and the recovery-draft decision.
 *
 * The package is mounted as the `session-document` Cordis plugin in base
 * profiles; its projection and prompt section are owned by the mounting fiber.
 *
 * @module @deepseek-ai/dsh-session-document
 */

export * from './document-changes.ts'
export * from './document-facts.ts'
export * from './document-identity.ts'
export * from './document-projection.ts'
export * from './document-recovery.ts'
export * from './document-restore.ts'
export { apply, inject, name } from './plugin.ts'
