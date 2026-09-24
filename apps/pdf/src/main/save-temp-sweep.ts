import { lstat, readdir, unlink } from 'node:fs/promises'
import { join } from 'node:path'

/**
 * Orphaned PDF-save temps (OBS-1670): atomicWriteFile commits each save with a
 * dot-prefixed temp next to the target and renames it over the target, but a
 * kill -9 mid-write bypasses every error-path unlink and the temp survives in
 * the user's directory forever. Nothing revisits a user-picked directory on
 * its own — so each fresh save sweeps the target directory's expired temps
 * first (savePdfToPath, right before the new temp is created).
 */

/**
 * One hour, mirroring the slides video-export sweep (OBS-1658). The write
 * below a live save takes seconds at most and the temp's mtime is set at
 * creation; age here means "time since the last write", so only a truly
 * orphaned temp can sit untouched that long — an in-flight save's temp (and
 * any concurrent save's temp) is never a candidate.
 */
export const SAVE_TEMP_TTL_MS = 60 * 60 * 1000

// The atomic-write temp signature in user directories: a dot-prefixed
// `<target>.<12 lowercase hex>.tmp` — the 12 hex chars are randomBytes(6),
// owned by atomicWriteFile in @airy-office/electron-utils (this app does not
// build temps itself). Kept in one place here so the sweep pattern cannot
// drift from the construction it mirrors.
const SAVE_TEMP_NAME = /^\..+\.[0-9a-f]{12}\.tmp$/

/**
 * Remove expired PDF-save temps from the save target's directory. Age alone
 * decides; non-matching files, directories, symlinks, and similarly named
 * nested paths are untouched; per-file errors are swallowed (another
 * window/process may hold a file), and a missing directory sweeps nothing —
 * the sweep never throws and never blocks the save it precedes.
 */
export async function sweepStaleSaveTemps(directory: string, now = Date.now()): Promise<string[]> {
  const cutoff = now - SAVE_TEMP_TTL_MS
  const removed: string[] = []
  let entries
  try {
    entries = await readdir(directory, { withFileTypes: true })
  } catch {
    // The directory may not exist yet (first save to a new folder) — nothing to sweep.
    return removed
  }
  for (const entry of entries) {
    if (!entry.isFile() || !SAVE_TEMP_NAME.test(entry.name)) continue
    const path = join(directory, entry.name)
    try {
      const info = await lstat(path)
      if (!info.isFile() || info.mtimeMs >= cutoff) continue
      await unlink(path)
      removed.push(path)
    } catch {
      // Best-effort sweep; the file may already be gone.
    }
  }
  return removed
}
