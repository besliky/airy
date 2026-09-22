/// Single-writer discipline for userData/app-settings.json — a flat JSON
/// object shared by the shell and every editor main (language, theme,
/// onboarding flag, dialog-dir LRU, …). Its correctness rule is simple but
/// easy to break: every write must be a read-merge-write that keeps keys it
/// did not touch.
///
/// OBS-1532: an audit twice observed a seeded `onboardingSeen` key vanish
/// mid-run while only {starPrompt, lastRunVersion} remained (self-healed by
/// the next merge-write). A static audit found no wholesale writer, but the
/// merge-write existed as two independent copies (shell app-settings.ts and
/// dialog-memory.ts) and the one async-context writer had no serialization.
/// This module closes the whole class:
///   1. ONE merge implementation — nothing else may write the file, so a
///      wholesale JSON.stringify writer cannot (re)appear unnoticed.
///   2. Async read-modify-write goes through a per-path promise queue: one
///      queued write in flight at a time, later writes chain behind it, and
///      the re-read happens inside the queued section, so no caller can
///      persist a snapshot taken before another write landed.
///   3. Synchronous merge writes stay synchronous (Electron's single main
///      thread already serializes them) and share the same code path.
/// The residual risk is a second PROCESS sharing the file (two instances on
/// one userData); the shell's single-instance lock covers that.
import { randomUUID } from 'node:crypto'
import { readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs'

/** app-settings.json is always a plain flat object of JSON values */
export type AppSettings = Record<string, unknown>

/**
 * A queued update: either a patch object merged over the fresh on-disk
 * state, or a reducer that receives the fresh state — read INSIDE its
 * queued section, never a snapshot captured earlier — and returns the next
 * full object.
 */
export type AppSettingsUpdate = AppSettings | ((current: AppSettings) => AppSettings)

export function readAppSettingsFile(settingsPath: string): AppSettings {
  try {
    const raw: unknown = JSON.parse(readFileSync(settingsPath, 'utf8'))
    if (raw && typeof raw === 'object' && !Array.isArray(raw)) {
      return raw as AppSettings
    }
  } catch {
    // missing or corrupt file: treat as empty settings
  }
  return {}
}

/**
 * The canonical write: merge `updates` into a FRESH read and replace the
 * file atomically. The temporary file lives beside the destination so the
 * rename never crosses filesystems, and a crash mid-write can never
 * truncate settings — readers see either the old or the new object.
 * Wholesale rewrites are impossible by construction: this is the only
 * writer, and it always merges.
 */
export function writeAppSettingsFile(settingsPath: string, updates: AppSettings): void {
  const settings = { ...readAppSettingsFile(settingsPath), ...updates }
  const tempPath = `${settingsPath}.${process.pid}.${randomUUID()}.tmp`
  try {
    writeFileSync(tempPath, JSON.stringify(settings, null, 2), {
      encoding: 'utf8',
      flag: 'wx',
      flush: true,
    })
    renameSync(tempPath, settingsPath)
  } catch (error) {
    try {
      unlinkSync(tempPath)
    } catch {
      // The write may have failed before the temporary file was created.
    }
    throw error
  }
}

/** per-settings-file write queues: at most one queued write in flight each */
const writeQueues = new Map<string, Promise<void>>()

/**
 * Serialized read-modify-write for async callers. Queued writes for the
 * same path run strictly one at a time (promise chain); a reducer's
 * `current` is read inside its turn, so updates computed earlier can never
 * clobber keys written in between. A failing write rejects its own promise
 * but never poisons the chain — the next queued write still runs.
 */
export function queueAppSettingsUpdate(
  settingsPath: string,
  update: AppSettingsUpdate,
): Promise<void> {
  const previous = writeQueues.get(settingsPath) ?? Promise.resolve()
  const run = previous.then(() => {
    if (typeof update === 'function') {
      const current = readAppSettingsFile(settingsPath)
      writeAppSettingsFile(settingsPath, update(current))
    } else {
      writeAppSettingsFile(settingsPath, update)
    }
  })
  // the chain continues past failures; only `run`'s caller sees the error
  const tail = run.catch(() => {})
  writeQueues.set(settingsPath, tail)
  // drop the queue entry once idle so a long-lived main leaks nothing
  void tail.then(() => {
    if (writeQueues.get(settingsPath) === tail) writeQueues.delete(settingsPath)
  })
  return run
}
