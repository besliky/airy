import { randomUUID } from 'node:crypto'
import { readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs'
import type { TabKind } from '../shared/tabs-api'

/**
 * Persisted open-tab set (userData/session.json): every file-backed tab in
 * strip order, plus which of them was active. In-memory untitled tabs (never
 * saved), the Home tab, and chrome-free present tabs have no backing file and
 * are skipped. Pure serialize/parse/prune logic lives here for unit tests;
 * the shell wires it into TabManager change notifications and launch.
 */
export interface SessionTabEntry {
  kind: TabKind
  path: string
}

export interface SessionState {
  tabs: SessionTabEntry[]
  /** backing path of the tab that was active when the session was saved */
  activePath: string | null
}

/** kinds that own files and can be restored; home is permanent, not session data */
const RESTORABLE_KINDS: ReadonlySet<string> = new Set([
  'docs',
  'sheets',
  'slides',
  'pdf',
  'markdown',
  'html',
])

/** cap against runaway/corrupt files opening thousands of tabs on launch */
const MAX_TABS = 64

export function serializeSession(
  tabs: Array<{ id: string; kind: TabKind; filePath?: string }>,
  activeId: string | null,
): SessionState {
  const entries: SessionTabEntry[] = []
  let activePath: string | null = null
  for (const tab of tabs) {
    if (!RESTORABLE_KINDS.has(tab.kind) || !tab.filePath) continue
    entries.push({ kind: tab.kind, path: tab.filePath })
    if (activeId !== null && tab.id === activeId) activePath = tab.filePath
  }
  return { tabs: entries, activePath }
}

/** Parse a raw JSON value (already JSON.parse-d); malformed yields an empty state. */
export function parseSession(raw: unknown): SessionState {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw))
    return { tabs: [], activePath: null }
  const record = raw as Record<string, unknown>
  if (!Array.isArray(record.tabs)) return { tabs: [], activePath: null }
  const tabs: SessionTabEntry[] = []
  const seen = new Set<string>()
  for (const entry of record.tabs) {
    if (tabs.length >= MAX_TABS) break
    if (entry === null || typeof entry !== 'object' || Array.isArray(entry)) continue
    const { kind, path } = entry as Record<string, unknown>
    if (typeof kind !== 'string' || !RESTORABLE_KINDS.has(kind) || typeof path !== 'string')
      continue
    if (!path || path.length > 4096) continue
    // duplicate paths would open the same file twice; the shell dedupes opens
    // anyway, so drop them here to keep the persisted set canonical
    if (seen.has(path)) continue
    seen.add(path)
    tabs.push({ kind: kind as TabKind, path })
  }
  const activePath =
    typeof record.activePath === 'string' && seen.has(record.activePath) ? record.activePath : null
  return { tabs, activePath }
}

/**
 * Drop entries whose file no longer exists (moved/deleted/disconnected drive):
 * restoring them would surface "file not found" dialogs for files the user
 * already dealt with. The active tab falling away leaves activation to the
 * caller (the last surviving tab stays active).
 */
export function pruneSession(
  state: SessionState,
  fileExists: (path: string) => boolean,
): SessionState {
  const tabs = state.tabs.filter((tab) => fileExists(tab.path))
  const activePath =
    state.activePath !== null && tabs.some((tab) => tab.path === state.activePath)
      ? state.activePath
      : null
  return { tabs, activePath }
}

/** Read and parse the session file; missing/corrupt content yields null. */
export function readSessionState(path: string): SessionState | null {
  let raw: unknown
  try {
    raw = JSON.parse(readFileSync(path, 'utf8'))
  } catch {
    return null
  }
  return parseSession(raw)
}

/** Persist atomically (temp + rename, same pattern as app-settings.json). */
export function writeSessionState(path: string, state: SessionState): void {
  const payload = { tabs: state.tabs, activePath: state.activePath }
  const tempPath = `${path}.${process.pid}.${randomUUID()}.tmp`
  try {
    writeFileSync(tempPath, JSON.stringify(payload, null, 2), {
      encoding: 'utf8',
      flag: 'wx',
      flush: true,
    })
    renameSync(tempPath, path)
  } catch (error) {
    try {
      unlinkSync(tempPath)
    } catch {
      // The write may have failed before the temporary file was created.
    }
    throw error
  }
}
