import { lstat, readdir, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

/**
 * Orphaned PDF-export temp dirs (BUG-1767, audit EXP-4): exportSlidesPdf
 * builds the print HTML in an app-owned mkdtemp directory and removes it in a
 * finally block, but a kill -9 mid-export bypasses every cleanup hook — the
 * multi-MB slides.html directory survives in the shared OS temp root forever
 * (same orphan family as the video streaming temps, OBS-1658, and the
 * generated-page dirs). Unlike the generated pages, nothing revisits this
 * prefix at startup, so each fresh PDF export sweeps the expired
 * airy-slides-pdf-* directories first.
 */

/** One hour, mirroring VIDEO_EXPORT_TEMP_TTL_MS: the print HTML is written
 * once at mkdtemp time and a live export finishes it long before the TTL, so
 * age here practically means "orphaned by a kill -9". */
export const PDF_EXPORT_TEMP_TTL_MS = 60 * 60 * 1000

/** mkdtemp prefix for PDF-export temp directories; kept next to the sweep
 * pattern so the construction and the regex cannot drift. */
export const PDF_EXPORT_TEMP_PREFIX = 'airy-slides-pdf-'

// mkdtemp appends exactly six characters from [a-zA-Z0-9]; anything else
// under a similar name is not ours.
const PDF_EXPORT_TEMP_DIR_NAME = new RegExp(`^${PDF_EXPORT_TEMP_PREFIX}[0-9a-zA-Z]{6}$`)

/**
 * Remove only expired PDF-export temp directories from the OS temp root.
 * Age alone decides; non-matching names, plain files, symlinks, and
 * similarly named nested paths are untouched; per-entry errors are swallowed
 * (another process may hold a directory), and a missing temp root sweeps
 * nothing — the sweep never throws and never blocks the export it precedes.
 */
export async function sweepStalePdfExportTempDirs(
  tempRoot: string = tmpdir(),
  now = Date.now(),
): Promise<string[]> {
  const cutoff = now - PDF_EXPORT_TEMP_TTL_MS
  const removed: string[] = []
  let entries
  try {
    entries = await readdir(tempRoot, { withFileTypes: true })
  } catch {
    // The temp root may be unwritable or gone — nothing to sweep.
    return removed
  }
  for (const entry of entries) {
    if (!entry.isDirectory() || !PDF_EXPORT_TEMP_DIR_NAME.test(entry.name)) continue
    const path = join(tempRoot, entry.name)
    try {
      const info = await lstat(path)
      if (!info.isDirectory() || info.mtimeMs >= cutoff) continue
      await rm(path, { recursive: true, force: true })
      removed.push(path)
    } catch {
      // Best-effort sweep; the directory may already be gone.
    }
  }
  return removed
}
