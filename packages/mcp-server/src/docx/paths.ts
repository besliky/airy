// Path confinement for every file the MCP tools touch: all input and output
// paths must resolve inside the workspace root, which comes from
// AIRY_WORKSPACE_ROOT (default: the process working directory). This is the
// headless twin of the renderer's allowlisted-paths IPC gate.
import { realpathSync } from 'node:fs'
import { basename, dirname, isAbsolute, join, resolve, sep } from 'node:path'

export const WORKSPACE_ROOT_ENV = 'AIRY_WORKSPACE_ROOT'

/** The confined root every docx path must live under. */
export function workspaceRoot(): string {
  const fromEnv = process.env[WORKSPACE_ROOT_ENV]
  if (fromEnv && fromEnv.trim() !== '') return resolve(fromEnv.trim())
  return resolve(process.cwd())
}

export class PathOutsideWorkspaceError extends Error {
  constructor(path: string, root: string) {
    super(
      `Path "${path}" resolves outside the workspace root "${root}". ` +
        `All paths must stay inside the root (set ${WORKSPACE_ROOT_ENV} to change it).`,
    )
    this.name = 'PathOutsideWorkspaceError'
  }
}

/**
 * The real path when it exists; for a non-existent path, the physical path of
 * its deepest EXISTING ancestor with the remaining tail components appended.
 *
 * The ancestor walk closes a save-as escape: a not-yet-existing target like
 * `root/link/new.docx` (where `link` is a directory symlink pointing outside
 * the root) makes the full realpathSync fail with ENOENT, but resolving the
 * existing `link` component pins the physical directory the missing tail
 * would be written into — so the prefix check sees the out-of-root location
 * instead of the in-root lexical spelling. Fresh in-root paths keep working:
 * their deepest existing ancestor resolves back inside the root (or is the
 * root itself), so the appended tail stays confined. A path with no existing
 * ancestor at all (up to the filesystem root) falls back to the lexical form
 * — the prefix check then applies to the spelling the caller gave.
 */
function realPathOrDeepestExisting(path: string): string {
  try {
    return realpathSync(path)
  } catch {
    const tail: string[] = []
    let current = path
    for (;;) {
      const parent = dirname(current)
      // dirname(x) === x only at the filesystem root — nothing exists to
      // resolve against, so the lexical spelling is the best answer left
      if (parent === current) return path
      tail.unshift(basename(current))
      current = parent
      try {
        return join(realpathSync(current), ...tail)
      } catch {
        // this ancestor does not exist either — keep climbing
      }
    }
  }
}

/**
 * Confinement prefix check. Windows filesystems are case-insensitive (and
 * case-preserving), so the comparison folds case there — a candidate spelled
 * `c:\users\...` must not be spuriously rejected against the root
 * `C:\Users\...`. Everywhere else the check stays case-sensitive like the
 * filesystem it guards.
 */
function isInsideRoot(candidate: string, root: string): boolean {
  const fold = process.platform === 'win32'
  const toComparable = (path: string) => (fold ? path.toLowerCase() : path)
  const prefix = toComparable(root.endsWith(sep) ? root : root + sep)
  const value = toComparable(candidate)
  return value === toComparable(root) || value.startsWith(prefix)
}

/**
 * Resolve `rawPath` against the workspace root and enforce confinement.
 *
 * Relative paths resolve against the root (not the process cwd, which may drift
 * between tool calls). Traversal (`../`) that escapes the root is rejected
 * lexically (path.resolve collapses `..` before the comparison), which also
 * covers URL-style and Windows-separator smuggling. Symlinks are then
 * resolved for BOTH the root and the candidate before the prefix check, so a
 * link that lives inside the root but points outside cannot smuggle paths out
 * of confinement (a link pointing elsewhere inside the root stays usable).
 * A non-existent candidate — a save-as to a fresh file — resolves through its
 * deepest existing ancestor, so an out-pointing symlink anywhere on the path
 * is still caught while genuinely fresh in-root paths stay usable. The
 * returned path keeps its lexical spelling — callers address files by the
 * path the agent gave, only the confinement decision uses the real paths.
 */
export function resolveConfined(rawPath: string, root = workspaceRoot()): string {
  if (typeof rawPath !== 'string' || rawPath.trim() === '') {
    throw new Error('A non-empty file path is required')
  }
  const resolved = isAbsolute(rawPath) ? resolve(rawPath) : resolve(root, rawPath)
  const realRoot = realPathOrDeepestExisting(root)
  const realCandidate = realPathOrDeepestExisting(resolved)
  if (!isInsideRoot(realCandidate, realRoot)) {
    throw new PathOutsideWorkspaceError(rawPath, root)
  }
  return resolved
}
