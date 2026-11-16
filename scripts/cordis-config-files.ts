/** Cordis Loader configuration file discovery. */

import { existsSync, globSync, readFileSync, statSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'

/**
 * Whether a discovered file is git's no-symlink representation of a link.
 *
 * A checkout without symlink support (`core.symlinks=false`, the default on
 * Windows) writes a tracked symlink as a regular file whose whole body is the
 * link target, without a trailing newline. Its target is a Loader config that
 * discovery already returned through its real path, so reading the
 * representation as YAML would only report a spurious scalar root.
 * @param root - repository root containing the file.
 * @param file - repository-relative discovered path.
 * @returns whether the file is a link target placeholder that resolves.
 */
function isSymlinkPlaceholder(root: string, file: string): boolean {
  const path = join(root, file)
  const source = readFileSync(path, 'utf8')
  if (source === '' || source.includes('\n')) return false
  const target = resolve(dirname(path), source)
  return existsSync(target) && statSync(target).isFile()
}

/**
 * Return repository-relative Cordis Loader YAML paths under `root`.
 *
 * Translation consistency records are YAML sidecars, never Loader inputs.
 * A symlink checked out without symlink support is excluded because its target
 * is discovered through its own path.
 *
 * @param root Repository root to scan.
 * @returns Sorted repository-relative Loader configuration paths.
 */
export function cordisConfigFiles(root: string): string[] {
  return globSync(['**/*cordis*.yml', '**/*cordis*.yaml'], {
    cwd: root,
    exclude: ['.claude/**', 'node_modules/**', 'vendor/**', '**/*.i18n.yaml'],
  })
    .sort()
    .filter(file => !isSymlinkPlaceholder(root, file))
}
