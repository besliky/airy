import { createHash } from 'node:crypto'
import { existsSync, statSync, unlinkSync } from 'node:fs'
import { join } from 'node:path'

/**
 * Crash-recovery copies for the plain-text editors (markdown/html), mirroring
 * docs' docs-autosave pattern: a dirty renderer pushes its serialized text
 * every ~30s (write-recovery IPC); a normal save or an explicit discard clears
 * the copy; opening a file whose copy is newer offers Restore/Discard.
 *
 * The copy lives one-per-document, named by a hash of its path, under
 * <userData>/<editor>-autosave/. A per-path epoch guards the clear/write race:
 * a write-recovery that started before a clear must not resurrect the copy the
 * user just discarded.
 */
export type RecoveryDecision = 'restore' | 'discard'

/** a copy is offered only when it is strictly newer than the saved file */
export function shouldOfferTextRecovery(recoveryMtimeMs: number, fileMtimeMs: number): boolean {
  return recoveryMtimeMs > fileMtimeMs
}

/** stable per-document copy name; the digest keeps user paths out of userData filenames */
export function textRecoveryPathFor(dir: string, filePath: string): string {
  const digest = createHash('sha1').update(filePath).digest('hex').slice(0, 16)
  return join(dir, `${digest}.recovery`)
}

export interface TextRecoveryStoreOptions {
  /** atomic write implementation (temp + rename), shared with the editor's save path */
  write: (path: string, text: string) => Promise<void>
  /** read the on-disk original (charset decoding stays the editor's job) */
  readOriginal: (path: string) => Promise<string>
}

export class TextRecoveryStore {
  private readonly epochs = new Map<string, number>()

  constructor(
    private readonly dir: string,
    private readonly options: TextRecoveryStoreOptions,
  ) {}

  pathFor(filePath: string): string {
    return textRecoveryPathFor(this.dir, filePath)
  }

  /** bump the epoch and delete the copy; races an in-flight writeCopy */
  clear(filePath: string): void {
    this.epochs.set(filePath, (this.epochs.get(filePath) ?? 0) + 1)
    try {
      unlinkSync(this.pathFor(filePath))
    } catch {
      /* nothing to clean */
    }
  }

  /** persist a recovery copy; no-ops when a clear happened since the call started */
  async writeCopy(filePath: string, text: string): Promise<void> {
    const epochAtStart = this.epochs.get(filePath) ?? 0
    const target = this.pathFor(filePath)
    try {
      await this.options.write(target, text)
    } catch (error) {
      // best-effort by contract: recovery must never surface in the editor
      console.warn('[recovery] copy write failed:', error)
      return
    }
    if ((this.epochs.get(filePath) ?? 0) !== epochAtStart) {
      // a save or discard cleared the copy while this write was in flight
      try {
        unlinkSync(target)
      } catch {
        /* already gone */
      }
    }
  }

  /**
   * Decide whether the copy for a just-opened file should be offered, ask the
   * user through the injected prompt, and return the text to load plus
   * whether it came from the copy. Cancel/anything-but-restore discards the
   * copy (docs' semantics: Esc on the prompt equals Discard).
   */
  async maybeRecover(
    filePath: string,
    prompt: () => Promise<RecoveryDecision>,
  ): Promise<{ text: string; recovered: boolean }> {
    const original = await this.options.readOriginal(filePath)
    const copyPath = this.pathFor(filePath)
    let newer = false
    try {
      if (
        existsSync(copyPath) &&
        existsSync(filePath) &&
        shouldOfferTextRecovery(statSync(copyPath).mtimeMs, statSync(filePath).mtimeMs)
      ) {
        newer = true
      }
    } catch {
      newer = false
    }
    if (!newer) return { text: original, recovered: false }
    if ((await prompt()) !== 'restore') {
      this.clear(filePath)
      return { text: original, recovered: false }
    }
    try {
      return { text: await this.options.readOriginal(copyPath), recovered: true }
    } catch {
      return { text: original, recovered: false }
    }
  }
}
