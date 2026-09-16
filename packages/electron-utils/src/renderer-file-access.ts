/**
 * Allowlist of directories whose files renderer processes may ask the main
 * process to read. A compromised renderer can otherwise name any local path
 * over the files:read* / *:open-path IPC channels and exfiltrate its bytes.
 *
 * Directories enter the list only through user-driven flows: native dialog
 * picks (dialog-memory), files the shell or an editor module opens, accepted
 * AI-panel attachments (dialogs and drag-drop), and the default save folder.
 * Enforcement happens in the per-app read handlers; the pure membership test
 * is exported for unit tests.
 */
import { dirname, posix, win32 } from 'node:path'

/** Bounded so a long session cannot grow the list without limit. */
const MAX_GRANTED_DIRS = 64

const grantedDirs = new Set<string>()

/** Normalize per the given platform (win32 paths compare case-insensitively). */
function normalizeDir(dir: string, platform: NodeJS.Platform): string {
  const normalized = platform === 'win32' ? win32.resolve(dir.trim()) : posix.resolve(dir.trim())
  return platform === 'win32' ? normalized.toLowerCase() : normalized
}

function separatorOf(platform: NodeJS.Platform): string {
  return platform === 'win32' ? '\\' : '/'
}

/** '/', 'C:\\' — granting a root would allow reading the whole disk. */
function isRootLike(normalized: string, platform: NodeJS.Platform): boolean {
  return platform === 'win32' ? /^[a-z]:\\?$/.test(normalized) : normalized === '/'
}

/** Pure test: is `path` inside `dir` (or equal to it)? Windows compares case-insensitively. */
export function isPathInsideDir(
  dir: string,
  path: string,
  platform: NodeJS.Platform = process.platform,
): boolean {
  const dirValue = normalizeDir(dir, platform)
  const pathValue = normalizeDir(path, platform)
  if (!dirValue || !pathValue) return false
  return pathValue === dirValue || pathValue.startsWith(`${dirValue}${separatorOf(platform)}`)
}

/** Pure test: is `path` inside any of `dirs`? */
export function pathIsInsideAny(dirs: readonly string[], path: string): boolean {
  return dirs.some((dir) => isPathInsideDir(dir, path))
}

/** Grant read access to every file inside `dir`. */
export function grantRendererDir(dir: string): void {
  const trimmed = dir.trim()
  // empty would resolve to the process cwd and silently grant it
  if (!trimmed) return
  const normalized = normalizeDir(trimmed, process.platform)
  if (!normalized || isRootLike(normalized, process.platform) || grantedDirs.has(normalized)) {
    return
  }
  grantedDirs.add(normalized)
  // FIFO eviction keeps the newest user-touched directories
  while (grantedDirs.size > MAX_GRANTED_DIRS) {
    const oldest = grantedDirs.keys().next().value as string | undefined
    if (oldest === undefined) break
    grantedDirs.delete(oldest)
  }
}

/** Grant read access to the directory containing `path` (the file itself included). */
export function grantRendererFileAccess(path: string): void {
  const trimmed = path.trim()
  if (trimmed) grantRendererDir(dirname(trimmed))
}

/** Whether a renderer-supplied path may be read back over IPC. */
export function rendererMayReadPath(path: string): boolean {
  const trimmed = path.trim()
  if (!trimmed) return false
  // Only absolute paths can be matched against granted directories; relative
  // ones would resolve against the main process cwd, which is never granted.
  if (!win32.isAbsolute(trimmed) && !posix.isAbsolute(trimmed)) return false
  return pathIsInsideAny([...grantedDirs], trimmed)
}

/** Test/inspection helpers */
export function grantedRendererDirs(): readonly string[] {
  return [...grantedDirs]
}

export function resetRendererFileGrants(): void {
  grantedDirs.clear()
}
