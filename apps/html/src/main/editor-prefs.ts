/// Remembered editor view preferences (UX-1704): the source pane's word-wrap
/// toggle and the optional split-view scroll-sync. The store is workspace-
/// level: one `htmlEditorPrefs` object inside the shared userData/
/// app-settings.json that the shell and every editor already write through
/// the single-writer queue (packages/electron-utils/src/app-settings-file.ts).
/// Writes go exclusively through queueAppSettingsUpdate (the PR #108 /
/// OBS-1532 discipline): the read-merge-write runs as one queued section and
/// re-reads the on-disk state inside it, so a toggle landing while other
/// settings writes are in flight can neither drop their keys nor be dropped
/// by them. Malformed persisted fields are individually reset to defaults —
/// one bad key never disables the other.
import { queueAppSettingsUpdate, readAppSettingsFile } from '@airy-office/electron-utils'
import { DEFAULT_EDITOR_PREFS } from '../shared/ipc'
import type { EditorPrefs, EditorPrefsPatch } from '../shared/ipc'

export type { EditorPrefs, EditorPrefsPatch }

/** where the prefs live inside app-settings.json */
export const EDITOR_PREFS_KEY = 'htmlEditorPrefs'

/** accept only real booleans; anything else falls back to the default */
function boolOr(value: unknown, fallback: boolean): boolean {
  return typeof value === 'boolean' ? value : fallback
}

/** parse the persisted object; malformed values are dropped field-by-field. Pure — unit-tested */
export function parseEditorPrefs(raw: unknown): EditorPrefs {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return { ...DEFAULT_EDITOR_PREFS }
  const o = raw as Record<string, unknown>
  return {
    wordWrap: boolOr(o.wordWrap, DEFAULT_EDITOR_PREFS.wordWrap),
    scrollSync: boolOr(o.scrollSync, DEFAULT_EDITOR_PREFS.scrollSync),
  }
}

/** apply a partial update over the current value; non-object patches are ignored. Pure — unit-tested */
export function mergeEditorPrefs(current: EditorPrefs, patch: unknown): EditorPrefs {
  if (!patch || typeof patch !== 'object' || Array.isArray(patch)) return current
  const p = patch as Record<string, unknown>
  return {
    wordWrap: boolOr(p.wordWrap, current.wordWrap),
    scrollSync: boolOr(p.scrollSync, current.scrollSync),
  }
}

/** the stored prefs for this workspace, or the defaults when nothing is stored yet */
export function readEditorPrefs(settingsPath: string): EditorPrefs {
  return parseEditorPrefs(readAppSettingsFile(settingsPath)[EDITOR_PREFS_KEY])
}

/**
 * Merge a patch through the shared single-writer queue and resolve with the
 * stored value. Rejects on I/O errors; callers treat persistence as
 * best-effort — the toggle still applies in-session.
 */
export async function writeEditorPrefs(
  settingsPath: string,
  patch: EditorPrefsPatch,
): Promise<EditorPrefs> {
  await queueAppSettingsUpdate(settingsPath, (current) => ({
    ...current,
    [EDITOR_PREFS_KEY]: mergeEditorPrefs(parseEditorPrefs(current[EDITOR_PREFS_KEY]), patch),
  }))
  return readEditorPrefs(settingsPath)
}
