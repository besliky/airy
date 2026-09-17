/**
 * Read policy for renderer-named merge sources (sheets workbook:open-for-merge).
 *
 * The channel feeds renderer-supplied paths straight into the open pipeline,
 * which reads arbitrary .csv bytes, converts .xls, and snapshots .xlsx — so a
 * compromised renderer could otherwise exfiltrate any local spreadsheet. The
 * same grant policy as the files read handlers applies: every path (after `~`
 * expansion) must sit inside a directory the user actually granted to THIS
 * renderer (dialog picks, shell-routed opens, accepted attachments, witnessed
 * drops). Legit callers keep working: the ribbon flow picks files through the
 * main-side dialog, and the AI merge tool only names paths that arrived as
 * accepted attachments — both grant the file's folder. The decision is pure
 * (deps injected) so it is unit-testable.
 */
import { basename, isAbsolute, join } from 'node:path'

/** Extensions the merge pipeline can open (must match prepareWorkbookForOpen). */
export const MERGE_SOURCE_EXTS = new Set(['xlsx', 'xlsm', 'xls', 'csv'])

export type MergeSourceRejection =
  { kind: 'ext'; ext: string } | { kind: 'missing' } | { kind: 'not-granted' }

export interface MergeSourceCheckDeps {
  /** app.getPath('home') — for expanding `~/`-prefixed renderer paths */
  homeDir: string
  /** renderer-file-access grant check, already scoped to the asking sender */
  mayRead: (path: string) => boolean
  exists: (path: string) => boolean
}

/** Expand a leading `~/` against the given home directory (raw otherwise). */
export function expandHomePath(path: string, homeDir: string): string {
  return path.startsWith('~/') ? join(homeDir, path.slice(2)) : path
}

/**
 * Validate one renderer-named merge source. Returns the resolved path to open
 * on success, or the rejection kind the handler maps to a localized error.
 */
export function checkMergeSourcePath(
  rawPath: string,
  deps: MergeSourceCheckDeps,
): { resolved: string } | { rejection: MergeSourceRejection } {
  const resolved = expandHomePath(rawPath, deps.homeDir)
  const name = basename(resolved)
  const dot = name.lastIndexOf('.')
  const ext = dot > 0 ? name.slice(dot + 1).toLowerCase() : ''
  if (!MERGE_SOURCE_EXTS.has(ext)) return { rejection: { kind: 'ext', ext } }
  // Only absolute paths can be matched against granted directories; a
  // relative one would resolve against the main process cwd (never granted).
  if (!isAbsolute(resolved)) return { rejection: { kind: 'not-granted' } }
  // The grant check runs before any filesystem touch (existsSync included) so
  // an ungranted path cannot even be probed for existence.
  if (!deps.mayRead(resolved)) return { rejection: { kind: 'not-granted' } }
  if (!deps.exists(resolved)) return { rejection: { kind: 'missing' } }
  return { resolved }
}

/**
 * Validate a whole merge-source list; either every path resolves (order kept)
 * or the first rejection wins — the caller opens nothing on rejection.
 */
export function checkMergeSourcePaths(
  rawPaths: readonly string[],
  deps: MergeSourceCheckDeps,
): { resolved: string[] } | { rejection: MergeSourceRejection } {
  const resolved: string[] = []
  for (const rawPath of rawPaths) {
    const checked = checkMergeSourcePath(rawPath, deps)
    if ('rejection' in checked) return checked
    resolved.push(checked.resolved)
  }
  return { resolved }
}
