import { randomBytes } from 'node:crypto'
import { open, rename, unlink } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'

/** Transient Windows codes: antivirus/indexer briefly locks the rename target. */
const RETRYABLE_RENAME_CODES = new Set(['EPERM', 'EACCES', 'EBUSY'])
const RENAME_RETRIES = 4

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

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
 */
export async function atomicWriteFile(filePath: string, data: Uint8Array): Promise<void> {
  const tmp = join(
    dirname(filePath),
    `.${basename(filePath)}.${randomBytes(6).toString('hex')}.tmp`,
  )
  try {
    const handle = await open(tmp, 'w')
    try {
      await handle.writeFile(data)
      await handle.sync()
    } finally {
      await handle.close()
    }
    await renameDurably(tmp, filePath)
  } catch (error) {
    const retryable = RETRYABLE_RENAME_CODES.has((error as NodeJS.ErrnoException).code ?? '')
    if (retryable) {
      // Preserve the completed temp until the non-atomic fallback succeeds.
      // If that write fails or the process exits, the new bytes still exist.
      await writeInPlaceDurably(filePath, data)
      await unlink(tmp).catch(() => {})
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
 */
export async function renameDurably(temporaryPath: string, targetPath: string): Promise<void> {
  for (let attempt = 0; ; attempt += 1) {
    try {
      await rename(temporaryPath, targetPath)
      // best-effort: the rename already succeeded, a failing directory
      // flush must not fail the save (unusual filesystems reject dir fsync)
      if (process.platform !== 'win32') {
        await syncDirectory(dirname(targetPath))
      }
      return
    } catch (error) {
      const code = (error as NodeJS.ErrnoException).code ?? ''
      if (!RETRYABLE_RENAME_CODES.has(code) || attempt >= RENAME_RETRIES) throw error
      await sleep(50 * 2 ** attempt)
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
