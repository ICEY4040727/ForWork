/**
 * Host-only document session primitives shared by the `document` projection,
 * restore, and the document tools: stable identity bindings across rename and
 * missing paths, and version-guarded attributable change records with their
 * inverses.
 *
 * The package registers no Cordis service and emits no Session event; callers
 * project the returned values into `document/*` facts.
 *
 * @module @deepseek-ai/dsh-session-document
 */

export * from './document-changes.ts'
export * from './document-identity.ts'
