/**
 * Per-renderer allowlist of directories whose files a renderer process may
 * ask the main process to read. A compromised renderer can otherwise name any
 * local path over the files:read* / *:open-path IPC channels and exfiltrate
 * its bytes.
 *
 * The allowlist is scoped PER SENDER (webContents id): a directory granted
 * through one tab's dialog pick is NOT readable by a different tab's
 * renderer — all renderers share one main process, so a process-global list
 * would let a compromised editor read a folder some other tab happened to
 * pick. Directories enter a sender's list only through user-driven flows for
 * THAT sender: native dialog picks (dialog-memory, which threads the
 * requester id), files the shell or an editor module opens into a specific
 * tab, accepted AI-panel attachments (dialogs and witnessed drag-drops), and
 * the default save folder. A grant covers the picked DIRECTORY (dirname of
 * the file, so relative assets resolve); entries are dropped when the
 * sender's webContents is destroyed (forgetRendererFileAccess). Enforcement
 * happens in the per-app read handlers; the pure membership test is exported
 * for unit tests.
 */
import { dirname, posix, win32 } from 'node:path'

/** Bounded so a long session cannot grow one sender's list without limit. */
const MAX_GRANTED_DIRS = 64
/** Bounded so destroyed-but-unforgotten senders cannot grow the map forever. */
const MAX_SENDERS = 64

/** sender webContents id → FIFO set of granted directories */
const grantedDirsBySender = new Map<number, Set<string>>()

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

function senderDirs(senderId: number): Set<string> {
  let dirs = grantedDirsBySender.get(senderId)
  if (!dirs) {
    dirs = new Set<string>()
    grantedDirsBySender.set(senderId, dirs)
    // FIFO eviction of stale senders (teardown normally removes them via
    // forgetRendererFileAccess; this is the safety net)
    while (grantedDirsBySender.size > MAX_SENDERS) {
      const oldest = grantedDirsBySender.keys().next().value as number | undefined
      if (oldest === undefined) break
      grantedDirsBySender.delete(oldest)
    }
  }
  return dirs
}

/** Grant `senderId` read access to every file inside `dir`. */
export function grantRendererDir(dir: string, senderId: number): void {
  const trimmed = dir.trim()
  // empty would resolve to the process cwd and silently grant it
  if (!trimmed) return
  const normalized = normalizeDir(trimmed, process.platform)
  const dirs = senderDirs(senderId)
  if (!normalized || isRootLike(normalized, process.platform) || dirs.has(normalized)) {
    return
  }
  dirs.add(normalized)
  // FIFO eviction keeps the newest user-touched directories
  while (dirs.size > MAX_GRANTED_DIRS) {
    const oldest = dirs.keys().next().value as string | undefined
    if (oldest === undefined) break
    dirs.delete(oldest)
  }
}

/** Grant `senderId` read access to the directory containing `path` (the file
 *  itself included). */
export function grantRendererFileAccess(path: string, senderId: number): void {
  const trimmed = path.trim()
  if (trimmed) grantRendererDir(dirname(trimmed), senderId)
}

/** Whether a path the renderer with id `senderId` supplied may be read back
 *  over IPC. Grants are per sender; another tab's grant does not count. */
export function rendererMayReadPath(senderId: number, path: string): boolean {
  const trimmed = path.trim()
  if (!trimmed) return false
  // Only absolute paths can be matched against granted directories; relative
  // ones would resolve against the main process cwd, which is never granted.
  if (!win32.isAbsolute(trimmed) && !posix.isAbsolute(trimmed)) return false
  const dirs = grantedDirsBySender.get(senderId)
  if (!dirs) return false
  return pathIsInsideAny([...dirs], trimmed)
}

/** Drop every grant held by a torn-down sender (call from its webContents
 *  'destroyed' handler). */
export function forgetRendererFileAccess(senderId: number): void {
  grantedDirsBySender.delete(senderId)
}

/** Test/inspection helpers */
export function grantedRendererDirs(senderId: number): readonly string[] {
  return [...(grantedDirsBySender.get(senderId) ?? [])]
}

export function resetRendererFileGrants(): void {
  grantedDirsBySender.clear()
}
