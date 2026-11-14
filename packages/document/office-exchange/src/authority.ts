/** Write-root confinement for Office paths supplied by a model or a tool. */

import { realpath } from 'node:fs/promises'
import { basename, dirname, join, resolve, sep } from 'node:path'
import { OfficeExchangeError } from './errors.ts'

/** Errno codes that mean a path component does not exist. */
const MISSING_CODES: ReadonlySet<string> = new Set(['ENOENT', 'ENOTDIR'])

/**
 * Directories a session may write Office documents into. Roots must exist;
 * callers resolve them once per session and reuse the list across operations.
 */
export interface WriteRoots {
  /** Absolute directory paths; a resolved target must equal one or live inside it. */
  readonly roots: readonly string[]
}

/**
 * Whether a filesystem failure reports an absent path. `fs.realpath` rejects
 * only with a Node errno error.
 * @param error - Rejection from `fs.realpath`.
 * @returns True for a missing path or a path component that is not a directory.
 */
function isMissing(error: unknown): boolean {
  const code = String((error as NodeJS.ErrnoException).code)
  return MISSING_CODES.has(code)
}

/**
 * Resolve a path through `fs.realpath`, mapping an absent path to `undefined`.
 * @param targetPath - Absolute path to resolve.
 * @returns The real path, or undefined when the path does not exist.
 * @throws When resolution fails for a reason other than absence.
 */
async function realpathIfPresent(targetPath: string): Promise<string | undefined> {
  try {
    return await realpath(targetPath)
  } catch (error) {
    /* v8 ignore next -- Node fs rejects only with ErrnoException; a portable test cannot produce another failure here. */
    if (!isMissing(error)) throw error
    return undefined
  }
}

/**
 * Resolve a path to its real location, allowing components that do not exist
 * yet by resolving the nearest existing ancestor and appending the remainder.
 * @param targetPath - Path to resolve; a relative path resolves against the process working directory.
 * @returns The resolved path, and whether trailing components were missing.
 * @throws {OfficeExchangeError} `out-of-scope` when no existing ancestor can be resolved.
 */
async function resolveNearestExisting(targetPath: string): Promise<{ readonly resolved: string; readonly missing: boolean }> {
  const segments: string[] = []
  let current = resolve(targetPath)
  for (;;) {
    const existing = await realpathIfPresent(current)
    if (existing !== undefined) {
      return {
        resolved: segments.length === 0 ? existing : join(existing, ...segments),
        missing: segments.length > 0,
      }
    }
    const parent = dirname(current)
    /* v8 ignore next 3 -- Every absolute path reaches its filesystem root before the loop can exhaust ancestors. */
    if (parent === current) {
      throw new OfficeExchangeError('out-of-scope', `Cannot resolve ${targetPath}: no existing ancestor.`)
    }
    segments.unshift(basename(current))
    current = parent
  }
}

/**
 * Whether a resolved target is a root or a descendant of it. Comparing against
 * the root plus the platform separator keeps a sibling whose name shares the
 * root's prefix outside.
 * @param root - Resolved root directory.
 * @param target - Resolved target path.
 * @returns True when the target equals the root or lies below its separator boundary.
 */
function isWithin(root: string, target: string): boolean {
  return target === root || target.startsWith(join(root, sep))
}

/**
 * Assert that a target path resolves inside one of the permitted write roots.
 * Roots and target are resolved through symbolic links, so a link that points
 * outside a root is rejected. A missing target is resolved through its nearest
 * existing ancestor, which lets a new file be created under a permitted root.
 * @param targetPath - Path the caller intends to read or write.
 * @param roots - Permitted root directories, each of which must exist.
 * @param options - `allowMissing` admits a not-yet-created target; it defaults to true.
 * @returns The resolved absolute target path.
 * @throws {OfficeExchangeError} `out-of-scope` when the target is outside every root, or is missing while disallowed.
 */
export async function assertWithinRoots(
  targetPath: string,
  roots: readonly string[],
  options: { readonly allowMissing?: boolean } = {},
): Promise<string> {
  const target = await resolveNearestExisting(targetPath)
  const allowMissing = options.allowMissing ?? true
  if (target.missing && !allowMissing) {
    throw new OfficeExchangeError(
      'out-of-scope',
      `${targetPath} does not exist and missing targets are not allowed (write roots: [${roots.join(', ')}])`,
    )
  }
  const resolvedRoots = await Promise.all(roots.map(root => realpath(resolve(root))))
  if (!resolvedRoots.some(root => isWithin(root, target.resolved))) {
    throw new OfficeExchangeError(
      'out-of-scope',
      `${targetPath} is outside the permitted write roots: [${roots.join(', ')}]`,
    )
  }
  return target.resolved
}
