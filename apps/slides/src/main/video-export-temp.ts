import { randomBytes } from 'node:crypto'
import { lstat, readdir, unlink } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'

/**
 * Orphaned video-export temps (OBS-1658): the streamed export writes a
 * dot-prefixed temp next to the user-picked target and commits it with a
 * rename, but a kill -9 mid-export bypasses every cleanup hook
 * (render-process-gone / destroyed never fire) and the temp survives in the
 * user's directory. Unlike the app-owned generated-page dirs, nothing
 * revisits a user-picked directory on its own — so each fresh export into a
 * directory sweeps its expired siblings first (video-file-stream-begin).
 */

/**
 * One hour. An in-progress export's temp is appended to continuously, so its
 * mtime stays fresh: age here means "time since the last write", and only a
 * truly orphaned temp can sit untouched that long — live exports (including
 * concurrent ones from other windows) are never candidates.
 */
export const VIDEO_EXPORT_TEMP_TTL_MS = 60 * 60 * 1000

// The suite's own temp signature in user directories: a dot-prefixed
// `<target>.<12 lowercase hex>.tmp` — the 12 hex chars are randomBytes(6),
// shared with atomicWriteFile's same-shape temps. Kept next to
// videoExportTempPath so the pattern and the construction cannot drift.
const VIDEO_EXPORT_TEMP_NAME = /^\..+\.[0-9a-f]{12}\.tmp$/

/** Same-directory temp path for a streamed video export (atomic commit). */
export function videoExportTempPath(targetPath: string): string {
  return join(dirname(targetPath), `.${basename(targetPath)}.${randomBytes(6).toString('hex')}.tmp`)
}

/**
 * Remove expired video-export temps from the export target's directory.
 * Age alone decides; non-matching files, directories, symlinks, and
 * similarly named nested paths are untouched; per-file errors are swallowed
 * (another window/process may hold a file), and a missing directory sweeps
 * nothing — the sweep never throws and never blocks the export it precedes.
 */
export async function sweepStaleVideoExportTemps(
  directory: string,
  now = Date.now(),
): Promise<string[]> {
  const cutoff = now - VIDEO_EXPORT_TEMP_TTL_MS
  const removed: string[] = []
  let entries
  try {
    entries = await readdir(directory, { withFileTypes: true })
  } catch {
    // The directory may have vanished between pick and begin — nothing to sweep.
    return removed
  }
  for (const entry of entries) {
    if (!entry.isFile() || !VIDEO_EXPORT_TEMP_NAME.test(entry.name)) continue
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
