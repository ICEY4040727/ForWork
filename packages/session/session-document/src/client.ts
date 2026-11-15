/**
 * Client-namespace projection of the document domain: a pure re-export of the
 * package's types outlet.
 *
 * Client code imports only this namespace (repo discipline), so `./client`
 * projects the same single-source content `./types` serves to Host consumers
 * with zero duplication and no Host value import.
 *
 * @module @deepseek-ai/dsh-session-document/client
 */

export type * from './types.ts'
