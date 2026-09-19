import { randomUUID } from 'node:crypto'
import { readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs'
import type { TabKind } from '../shared/tabs-api'

/**
 * Persisted open-tab set (userData/session.json): every file-backed tab in
 * strip order, grouped per window in window-creation order, plus which tab of
 * each window was active and which window was focused. In-memory untitled
 * tabs (never saved), the Home tab, and chrome-free present tabs have no
 * backing file and are skipped. The format is window-ordered so a session
 * with a tab moved to a second window restores both windows; the legacy
 * single-window shape ({ tabs, activePath }) is parsed as one window. Pure
 * serialize/parse/prune logic lives here for unit tests; the shell wires it
 * into TabManager change notifications and launch.
 */
export interface SessionTabEntry {
  kind: TabKind
  path: string
}

/** one restored window: its file-backed tabs in strip order + active tab */
export interface SessionWindowState {
  tabs: SessionTabEntry[]
  /** backing path of the tab that was active when the session was saved */
  activePath: string | null
}

export interface SessionState {
  windows: SessionWindowState[]
  /** index into windows of the window that was focused when the session was saved */
  focusedWindow: number
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

/** serialize one window's tab list (strip order preserved, file-backed only) */
export function serializeSessionWindow(
  tabs: Array<{ id: string; kind: TabKind; filePath?: string }>,
  activeId: string | null,
): SessionWindowState {
  const entries: SessionTabEntry[] = []
  let activePath: string | null = null
  for (const tab of tabs) {
    if (!RESTORABLE_KINDS.has(tab.kind) || !tab.filePath) continue
    entries.push({ kind: tab.kind, path: tab.filePath })
    if (activeId !== null && tab.id === activeId) activePath = tab.filePath
  }
  return { tabs: entries, activePath }
}

/** serialize every live window, in window-creation order, with the focused one */
export function serializeSession(
  windows: Array<{
    tabs: Array<{ id: string; kind: TabKind; filePath?: string }>
    activeId: string | null
  }>,
  focusedWindow: number,
): SessionState {
  return {
    windows: windows.map((w) => serializeSessionWindow(w.tabs, w.activeId)),
    focusedWindow:
      Number.isInteger(focusedWindow) && focusedWindow >= 0 && focusedWindow < windows.length
        ? focusedWindow
        : 0,
  }
}

/** Parse one window's raw entry list (shared by the v2 and legacy shapes). */
function parseWindowTabs(raw: unknown, seen: Set<string>): SessionTabEntry[] {
  if (!Array.isArray(raw)) return []
  const tabs: SessionTabEntry[] = []
  for (const entry of raw) {
    if (tabs.length >= MAX_TABS) break
    if (entry === null || typeof entry !== 'object' || Array.isArray(entry)) continue
    const { kind, path } = entry as Record<string, unknown>
    if (typeof kind !== 'string' || !RESTORABLE_KINDS.has(kind) || typeof path !== 'string')
      continue
    if (!path || path.length > 4096) continue
    // duplicate paths would open the same file twice (in one window or two);
    // the shell dedupes opens anyway, so drop them here to keep the persisted
    // set canonical
    if (seen.has(path)) continue
    seen.add(path)
    tabs.push({ kind: kind as TabKind, path })
  }
  return tabs
}

/** Parse a raw JSON value (already JSON.parse-d); malformed yields an empty state. */
export function parseSession(raw: unknown): SessionState {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw))
    return { windows: [], focusedWindow: 0 }
  const record = raw as Record<string, unknown>
  // legacy single-window shape: { tabs, activePath } → one focused window
  if (Array.isArray(record.tabs)) return { windows: [parseLegacyWindow(record)], focusedWindow: 0 }
  if (!Array.isArray(record.windows)) return { windows: [], focusedWindow: 0 }
  const seen = new Set<string>()
  const windows: SessionWindowState[] = []
  let total = 0
  for (const rawWindow of record.windows) {
    if (total >= MAX_TABS) break
    if (rawWindow === null || typeof rawWindow !== 'object' || Array.isArray(rawWindow)) continue
    const tabs = parseWindowTabs((rawWindow as Record<string, unknown>).tabs, seen)
    if (tabs.length === 0) continue
    const activePath = (rawWindow as Record<string, unknown>).activePath
    windows.push({
      tabs,
      activePath: typeof activePath === 'string' && seen.has(activePath) ? activePath : null,
    })
    total += tabs.length
  }
  const focused = record.focusedWindow
  return {
    windows,
    focusedWindow:
      typeof focused === 'number' && Number.isInteger(focused) && focused >= 0
        ? Math.min(focused, Math.max(0, windows.length - 1))
        : 0,
  }
}

function parseLegacyWindow(record: Record<string, unknown>): SessionWindowState {
  const tabs = parseWindowTabs(record.tabs, new Set())
  const activePath = record.activePath
  return {
    tabs,
    activePath:
      typeof activePath === 'string' && tabs.some((tab) => tab.path === activePath)
        ? activePath
        : null,
  }
}

/**
 * Drop entries whose file no longer exists (moved/deleted/disconnected drive):
 * restoring them would surface "file not found" dialogs for files the user
 * already dealt with. Windows left without a single tab disappear (nothing to
 * restore in them); an active tab falling away leaves activation to the
 * caller (the last surviving tab stays active), and a focused window that
 * vanished passes the focus to the last surviving one.
 */
export function pruneSession(
  state: SessionState,
  fileExists: (path: string) => boolean,
): SessionState {
  const windows = state.windows
    .map((w) => ({
      tabs: w.tabs.filter((tab) => fileExists(tab.path)),
      activePath:
        w.activePath !== null &&
        w.tabs.some((tab) => tab.path === w.activePath && fileExists(tab.path))
          ? w.activePath
          : null,
    }))
    .filter((w) => w.tabs.length > 0)
  return {
    windows,
    focusedWindow:
      state.focusedWindow < windows.length ? state.focusedWindow : Math.max(0, windows.length - 1),
  }
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
  const payload = { windows: state.windows, focusedWindow: state.focusedWindow }
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
