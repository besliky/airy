// Path confinement for every file the MCP tools touch: all input and output
// paths must resolve inside the workspace root, which comes from
// AIRY_WORKSPACE_ROOT (default: the process working directory). This is the
// headless twin of the renderer's allowlisted-paths IPC gate.
import { isAbsolute, resolve, sep } from 'node:path'

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
 * Resolve `rawPath` against the workspace root and enforce confinement.
 *
 * Relative paths resolve against the root (not the process cwd, which may drift
 * between tool calls). Traversal (`../`) that escapes the root is rejected; the
 * check is lexical (path.resolve collapses `..` before the comparison), which
 * also covers URL-style and Windows-separator smuggling. Symlinks that point
 * outside the root are not chased — v1 keeps the check cheap and predictable.
 */
export function resolveConfined(rawPath: string, root = workspaceRoot()): string {
  if (typeof rawPath !== 'string' || rawPath.trim() === '') {
    throw new Error('A non-empty file path is required')
  }
  const resolved = isAbsolute(rawPath) ? resolve(rawPath) : resolve(root, rawPath)
  if (resolved !== root && !resolved.startsWith(root + sep)) {
    throw new PathOutsideWorkspaceError(rawPath, root)
  }
  return resolved
}
