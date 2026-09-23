import { statSync } from 'node:fs'

/**
 * Staleness fence for in-place saves (BUG-1654): a blind write over a saved
 * path resurrects a file that was renamed/deleted externally (silent fork)
 * and lets a second window on the same file win last-writer-wins without a
 * warning. Each app captures a stamp when a document is opened and refreshes
 * it after every successful save; before overwriting the same path the save
 * handler compares the stamp against the file's current identity and refuses
 * (with an Overwrite / Save As / Cancel prompt) on any mismatch.
 */
export interface FileStamp {
  mtimeMs: number
  size: number
}

export type StalenessVerdict =
  /** file exists and its mtime+size still match the stamp */
  | 'fresh'
  /** file exists but was modified since the stamp was taken */
  | 'changed'
  /** the path no longer exists (renamed or deleted externally) */
  | 'missing'

/** Capture the current identity of a file, or null when it does not exist. */
export function statFileStamp(path: string): FileStamp | null {
  try {
    const stats = statSync(path)
    return { mtimeMs: stats.mtimeMs, size: stats.size }
  } catch {
    return null
  }
}

/**
 * Compare a file's current identity against a stamp taken at open / last
 * save. A missing stamp counts as fresh: the fence has no baseline to
 * compare against (first save, pre-existing sessions), which degrades to
 * the old unfenced behavior instead of blocking legitimate saves.
 *
 * Call this as late as possible before the write (sync stat immediately
 * followed by the atomic write) — the fence narrows the external-change
 * window to the write itself but cannot close it entirely: a writer racing
 * between the stat and the rename still wins, the same accepted residual
 * as the BUG-1305 stat→read fence.
 */
export function checkSaveStaleness(
  path: string,
  stamp: FileStamp | null | undefined,
): StalenessVerdict {
  if (!stamp) return 'fresh'
  const current = statFileStamp(path)
  if (!current) return 'missing'
  return current.mtimeMs === stamp.mtimeMs && current.size === stamp.size ? 'fresh' : 'changed'
}
