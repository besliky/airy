/**
 * Main-side record of files the user really dragged or pasted into a
 * renderer window, used to keep the files:add attachment channel from
 * widening the renderer read allowlist on its own.
 *
 * The witness is written only by the preload-world drop/paste listeners in
 * drop-open.ts: page code cannot reach ipcRenderer.send directly, so a
 * compromised renderer cannot forge entries — it can only replay paths the
 * user actually dropped into that very webContents. record/query/prune are
 * pure module state so the policy is unit-testable.
 */
import { rendererMayReadPath } from './renderer-file-access'


/** Generous window for attach-then-read chats; purely memory hygiene. */
const WITNESS_TTL_MS = 30 * 60_000
/** Bounded per sender so a drag-happy session cannot grow memory without limit. */
const MAX_WITNESSED_PER_SENDER = 128
/** Mirrors drop-open's cap on how many files one event may carry. */
const MAX_WITNESSED_PER_EVENT = 20

/** sender webContents id → absolute path → witnessed-at (epoch ms) */
const witnessedBySender = new Map<number, Map<string, number>>()

function prune(map: Map<string, number>, now: number): void {
  for (const [path, at] of map) {
    if (now - at > WITNESS_TTL_MS) map.delete(path)
  }
  while (map.size > MAX_WITNESSED_PER_SENDER) {
    const oldest = map.keys().next().value as string | undefined
    if (oldest === undefined) break
    map.delete(oldest)
  }
}

/** Record paths the given sender really dropped/pasted (invalid input is ignored). */
export function recordWitnessedDrops(
  senderId: number,
  paths: unknown,
  now: number = Date.now(),
): void {
  if (!Array.isArray(paths)) return
  const map = witnessedBySender.get(senderId) ?? new Map<string, number>()
  let recorded = 0
  for (const entry of paths) {
    if (typeof entry !== 'string') continue
    const path = entry.trim()
    if (!path) continue
    if (recorded >= MAX_WITNESSED_PER_EVENT) break
    map.set(path, now)
    recorded++
  }
  prune(map, now)
  witnessedBySender.set(senderId, map)
}

/** Was `path` dropped/pasted into the given sender within the TTL? */
export function witnessedDroppedPath(
  senderId: number,
  path: string,
  now: number = Date.now(),
): boolean {
  const map = witnessedBySender.get(senderId)
  if (!map) return false
  const trimmed = path.trim()
  const at = map.get(trimmed)
  if (at === undefined) return false
  if (now - at > WITNESS_TTL_MS) {
    map.delete(trimmed)
    return false
  }
  return true
}

/** Forget everything witnessed for a torn-down sender. */
export function forgetWitnessedDrops(senderId: number): void {
  witnessedBySender.delete(senderId)
}

/** Test helper. */
export function resetWitnessedDrops(): void {
  witnessedBySender.clear()
}

/** Policy for the files:add attachment grant: an accepted attachment may widen
 *  the read allowlist only when the user really dragged/pasted that exact file
 *  into the asking renderer (witness), or its folder is already readable (the
 *  grant widens nothing). A renderer-named path with neither origin is
 *  accepted into the attachment list but grants nothing, so the read
 *  channels still refuse it.
 */
export function mayGrantAttachmentRead(
  senderId: number,
  path: string,
  now: number = Date.now(),
): boolean {
  return witnessedDroppedPath(senderId, path, now) || rendererMayReadPath(senderId, path)
}

// ── files:add payload shape ────────────────────────────────────────────────

/** Same bounds sheets enforces with zod (z.string().min(1).max(1024), max 50). */
export const MAX_ATTACHMENT_ADD_PATHS = 50
export const MAX_ATTACHMENT_PATH_CHARS = 1024

/**
 * Shape policy for the files:add channel, shared by the apps that validate
 * payloads by hand (docs/slides/html; sheets enforces the same bounds with
 * zod): the payload must be a bounded array of non-empty, length-bounded
 * strings. Returns the trimmed list, or null when the shape is wrong — the
 * handler then processes nothing (fail closed) instead of stat-ing an
 * unbounded or non-string renderer-supplied list.
 */
export function parseAttachmentPaths(raw: unknown): string[] | null {
  if (!Array.isArray(raw) || raw.length === 0 || raw.length > MAX_ATTACHMENT_ADD_PATHS) {
    return null
  }
  const paths: string[] = []
  for (const entry of raw) {
    if (typeof entry !== 'string') return null
    const path = entry.trim()
    if (!path || path.length > MAX_ATTACHMENT_PATH_CHARS) return null
    paths.push(path)
  }
  return paths
}
