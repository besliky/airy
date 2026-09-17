import { readdirSync, rmSync } from 'node:fs'
import { join, sep } from 'node:path'

/**
 * Untitled staging area (userData/untitled-staging): "New spreadsheet" and
 * "New PDF" need a real file on disk (their editors have no in-memory blank
 * mode), but writing blank files straight into the default save folder
 * litters Documents/Airy with abandoned "Untitled Spreadsheet.xlsx" copies.
 * The blank file is staged here instead; the first explicit save opens the
 * Save dialog defaulting to the default save folder, and the staged file is
 * removed (the tab rebinds to the picked path). Closing without saving
 * deletes the staged file. A crash leaves survivors in the staging dir —
 * session restore reopens them; anything left after that is a crash leftover
 * from a session that never restored and gets purged on the next launch.
 */

export function untitledStagingDir(userDataDir: string): string {
  return join(userDataDir, 'untitled-staging')
}

/** path-container check that resists prefix tricks (no `dirfoo` matching `dir`) */
export function isInsideDirectory(dir: string, path: string): boolean {
  if (!path.startsWith(dir)) return false
  return path.length === dir.length || path[dir.length] === sep
}

/** staged files worth offering back: only the extensions the shell stages */
const STAGED_EXTENSIONS = new Set(['.xlsx', '.pdf'])

export function listStagedFiles(dir: string): string[] {
  try {
    return readdirSync(dir)
      .filter((name) => STAGED_EXTENSIONS.has(name.slice(name.lastIndexOf('.'))))
      .map((name) => join(dir, name))
  } catch {
    return [] // no staging dir yet — nothing staged
  }
}

/**
 * Delete a staged file. Refuses anything outside the staging dir (the tab's
 * path may have been rebound to a real save location by the time this runs),
 * which also makes this idempotent for already-moved files.
 */
export function removeStagedFile(dir: string, path: string): boolean {
  if (!isInsideDirectory(dir, path)) return false
  try {
    rmSync(path, { force: true })
    return true
  } catch {
    return false
  }
}

/**
 * Staged files no open tab owns anymore. Run once at launch after session
 * restore: a clean quit already purged its staged tabs, so survivors here are
 * crash leftovers whose session was not restored (restore disabled, or the
 * entry was pruned) — deleting them keeps the staging dir self-cleaning.
 */
export function orphanedStagedFiles(stagedPaths: string[], openPaths: string[]): string[] {
  const open = new Set(openPaths)
  return stagedPaths.filter((path) => !open.has(path))
}
