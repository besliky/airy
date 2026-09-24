import { randomBytes } from 'node:crypto'
import { lstat, open, realpath, rename, unlink } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'

/** Transient Windows codes: antivirus/indexer briefly locks the rename target. */
const RETRYABLE_RENAME_CODES = new Set(['EPERM', 'EACCES', 'EBUSY'])
const RENAME_RETRIES = 4
const RENAME_RETRY_BASE_DELAY_MS = 50

// Test seam: the exponential rename-retry backoff is honest production
// pacing for transient Windows AV/indexer locks, but its 50..800 ms sleeps
// dominate every suite that forces the retries — the retry COUNT and the
// fallback routing are the contract, the wall-clock pacing is not (PERF-1101).
// Tests collapse the base delay to ~0; production always runs the real ladder.
let renameRetryDelayMs = RENAME_RETRY_BASE_DELAY_MS

/** Test-only: collapse the rename retry backoff (null restores the 50 ms ladder). */
export function _setRenameRetryDelayForTests(ms: number | null): void {
  renameRetryDelayMs = ms ?? RENAME_RETRY_BASE_DELAY_MS
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

/**
 * A save target is a symlink whose destination does not exist and the caller
 * can only promote a finished temp file via rename (renameDurably): rename(2)
 * cannot follow a link, and with no bytes in hand there is no in-place write
 * to fall back on — refusing beats silently destroying the link (BUG-1695).
 */
export class BrokenSymlinkTargetError extends Error {
  constructor(targetPath: string) {
    super(
      `Cannot save: "${targetPath}" is a broken symlink (its destination does not exist). ` +
        'Remove the link or save to its destination path directly.',
    )
    this.name = 'BrokenSymlinkTargetError'
  }
}

/**
 * BUG-1695: rename(2) does not follow symlinks — renaming a temp onto a
 * symlink path silently REPLACES the link with a regular file, forking the
 * document in two (the destination keeps the old bytes under its own name
 * while the save reports success). Resolve an existing symlink target to its
 * destination so the temp+rename land on the real file and the link survives;
 * this follow-the-link semantics is exactly what a plain write would do, so
 * no caller can reach a location it could not reach before the fix. Returns
 * the path to rename onto — the resolved destination, or the original path
 * when it is not a symlink (the common case, one extra lstat). A DANGLING
 * link (destination missing, or unreachable so even the kernel could not
 * follow it) resolves to null for the caller to handle honestly.
 */
async function renameTargetFor(filePath: string): Promise<string | null> {
  const info = await lstat(filePath).catch(() => null)
  if (!info?.isSymbolicLink()) return filePath
  return realpath(filePath).catch(() => null)
}

/**
 * Same-dir temp file + rename, so a crash mid-write can't truncate the target.
 * Rename-over-existing fails transiently on Windows under Defender/indexer
 * locks: retry with backoff, then fall back to an in-place write — losing
 * atomicity for that one save beats failing a save a plain writeFileSync would
 * have completed. Shared by every app main (docs/sheets/pdf/markdown/html);
 * formerly one near-identical copy per app.
 *
 * Durability (power-loss, not crash): the temp file is fsynced before the
 * rename — without it a renamed entry can surface empty or truncated after
 * power loss on several filesystems. On POSIX the parent directory is fsynced
 * after the rename so the new directory entry itself is durable. Windows
 * cannot open a directory for fsync (EPERM), so there the directory flush is
 * skipped: NTFS's metadata journaling is the accepted residual risk.
 *
 * Symlinks (BUG-1695): an existing symlink target is followed — the temp is
 * created in the destination's directory and renamed over the destination, so
 * the link itself is never replaced and readers through the link see the
 * save. A dangling link cannot be resolved; the save then rides the durable
 * in-place write on the link path, where open('w') follows the link and
 * creates the destination exactly like a plain write would — the link stays
 * a link, and a missing destination directory surfaces as an honest ENOENT.
 */
export async function atomicWriteFile(filePath: string, data: Uint8Array): Promise<void> {
  const renameTarget = await renameTargetFor(filePath)
  if (renameTarget === null) {
    // Dangling symlink: there is nothing to rename over and destroying the
    // link is the one forbidden outcome — create through it instead (BUG-1695).
    await writeInPlaceDurably(filePath, data)
    return
  }
  const tmp = join(
    dirname(renameTarget),
    `.${basename(renameTarget)}.${randomBytes(6).toString('hex')}.tmp`,
  )
  try {
    const handle = await open(tmp, 'w')
    try {
      await handle.writeFile(data)
      await handle.sync()
    } finally {
      await handle.close()
    }
    await renameDurably(tmp, renameTarget)
  } catch (error) {
    const retryable = RETRYABLE_RENAME_CODES.has((error as NodeJS.ErrnoException).code ?? '')
    if (retryable) {
      // The temp file is cleaned on BOTH fallback outcomes (BUG-1222): on
      // the double-failure path — rename stayed locked AND the in-place
      // write failed — the old code skipped both unlinks and left a hidden
      // .<name>.<hex>.tmp next to the target forever, the same orphan class
      // withSaveTmpCleanup fixed for the MCP saves. The save itself already
      // failed and reported; the leftover bytes have no recovery value.
      try {
        await writeInPlaceDurably(renameTarget, data)
      } finally {
        await unlink(tmp).catch(() => {})
      }
      return
    }
    await unlink(tmp).catch(() => {})
    throw error
  }
}

/**
 * Rename a fully written (and already fsynced) temp file over its target with
 * the shared durability discipline: retry transient EPERM/EACCES/EBUSY with
 * backoff, then fsync the parent directory on POSIX so the new directory
 * entry survives power loss. Exported for the sheets xlsx save paths, which
 * carried diverging copies of this protocol (BUG-1203): the caller must have
 * fsynced the temp file's contents before calling. A retryable error that
 * survives every retry rethrows with its original code so the caller can run
 * its own non-atomic fallback.
 *
 * Symlinks (BUG-1695): an existing symlink target is followed — the rename
 * lands on the link's destination so the link survives. A dangling link
 * cannot be followed by a rename and no in-place bytes are available here,
 * so it refuses with BrokenSymlinkTargetError instead of replacing the link.
 */
export async function renameDurably(temporaryPath: string, targetPath: string): Promise<void> {
  const renameTarget = await renameTargetFor(targetPath)
  if (renameTarget === null) throw new BrokenSymlinkTargetError(targetPath)
  for (let attempt = 0; ; attempt += 1) {
    try {
      await rename(temporaryPath, renameTarget)
      // best-effort: the rename already succeeded, a failing directory
      // flush must not fail the save (unusual filesystems reject dir fsync)
      if (process.platform !== 'win32') {
        await syncDirectory(dirname(renameTarget))
      }
      return
    } catch (error) {
      const code = (error as NodeJS.ErrnoException).code ?? ''
      if (!RETRYABLE_RENAME_CODES.has(code) || attempt >= RENAME_RETRIES) throw error
      await sleep(renameRetryDelayMs * 2 ** attempt)
    }
  }
}

/**
 * The non-atomic last resort after a persistently locked rename. In-place
 * writes change no directory entry, so there is no dir-fsync to do — but the
 * bytes themselves must be flushed (BUG-1203): a bare writeFile could lose
 * the whole save on power loss even though the write returned.
 */
async function writeInPlaceDurably(filePath: string, data: Uint8Array): Promise<void> {
  const handle = await open(filePath, 'w')
  try {
    await handle.writeFile(data)
    await handle.sync()
  } finally {
    await handle.close()
  }
}

async function syncDirectory(directory: string): Promise<void> {
  const handle = await open(directory, 'r').catch(() => null)
  if (handle === null) return
  try {
    await handle.sync()
  } catch {
    // journaled/network filesystems may refuse a directory fsync — the file
    // data itself is already durable at this point
  } finally {
    await handle.close().catch(() => {})
  }
}

/** 'PK\x03\x04' local-file-header check — cheap docx/zip sanity test. */
export function looksLikeZip(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 4 &&
    bytes[0] === 0x50 && // 'P'
    bytes[1] === 0x4b && // 'K'
    bytes[2] === 0x03 &&
    bytes[3] === 0x04
  )
}
