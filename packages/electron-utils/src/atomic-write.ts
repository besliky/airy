import { randomBytes } from 'node:crypto'
import { rename, unlink, writeFile } from 'node:fs/promises'
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
 */
export async function atomicWriteFile(filePath: string, data: Uint8Array): Promise<void> {
  const tmp = join(
    dirname(filePath),
    `.${basename(filePath)}.${randomBytes(6).toString('hex')}.tmp`,
  )
  try {
    await writeFile(tmp, data)
    for (let attempt = 0; ; attempt += 1) {
      try {
        await rename(tmp, filePath)
        return
      } catch (error) {
        const code = (error as NodeJS.ErrnoException).code ?? ''
        if (!RETRYABLE_RENAME_CODES.has(code) || attempt >= RENAME_RETRIES) throw error
        await sleep(50 * 2 ** attempt)
      }
    }
  } catch (error) {
    const retryable = RETRYABLE_RENAME_CODES.has((error as NodeJS.ErrnoException).code ?? '')
    if (retryable) {
      // Preserve the completed temp until the non-atomic fallback succeeds.
      // If that write fails or the process exits, the new bytes still exist.
      await writeFile(filePath, data)
      await unlink(tmp).catch(() => {})
      return
    }
    await unlink(tmp).catch(() => {})
    throw error
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
