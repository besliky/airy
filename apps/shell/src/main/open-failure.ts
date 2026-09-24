/// Why a user-intended document open failed (BUG-1655): an unreadable file
/// (EACCES/EPERM), a moved or deleted one (ENOENT) or a directory wearing a
/// document extension (EISDIR) used to end in silence — no dialog, no tab,
/// the shell just fell back to the home screen. Dependency-injected like
/// dropped-files.ts so the classification and reporting stay unit-testable
/// without Electron.
import { accessSync, constants, statSync } from 'node:fs'
import { basename } from 'node:path'

/**
 * errno-style reason the path cannot be opened as a document, or null when
 * the path exists and is readable (any failure beyond this preflight — a
 * locked file, a corrupt document — belongs to the module that loads it).
 * The returned Error always carries the errno-style `code`, so the shared
 * friendly-error mapping can localize it.
 */
export function classifyOpenFailure(filePath: string): Error | null {
  try {
    if (statSync(filePath).isDirectory()) {
      return Object.assign(
        new Error(`EISDIR: illegal operation on a directory, open ${filePath}`),
        {
          code: 'EISDIR',
        },
      )
    }
    // stat() only needs +x on the parent directories: a chmod 000 file in a
    // readable directory stats fine and would otherwise die later, unseen, in
    // the renderer. The R_OK probe catches exactly that case.
    accessSync(filePath, constants.R_OK)
    return null
  } catch (err) {
    return err instanceof Error ? err : new Error(String(err))
  }
}

/** the existing main-process error-dialog channel plus its localized title */
export interface OpenFailureDeps {
  /**
   * showErrorDialog from error-dialog.ts; the window is resolved by the caller.
   * `onClosed` fires when the raised dialog is dismissed (or immediately when
   * the channel swallows the call because another dialog owns the screen), so
   * the per-path dedupe below can release its entry.
   */
  showErrorDialog: (message: string, err: Error, onClosed?: () => void) => void
  /** localized dialog title; receives the failing file's basename */
  openFailedMessage: (name: string) => string
}

/**
 * Paths whose failure dialog is currently on screen (BUG-1678). The unpacked
 * single-instance retry loop re-broadcasts the forwarded open up to 21 times,
 * and each broadcast re-reported the failure — a cascade of identical dialogs
 * for one path. While a path's dialog is still up, repeat failures of the SAME
 * path are swallowed; different paths keep their own dialogs, and once the
 * dialog is dismissed the path may report again.
 */
const activeDialogPaths = new Set<string>()

/** forget every active-dialog marker; test teardown only */
export function resetOpenFailureDedupe(): void {
  activeDialogPaths.clear()
}

/**
 * Surface WHY an intended open (CLI argument, double-click, macOS open-file)
 * produced no tab: classify the path and raise the friendly error dialog.
 * Returns true when a dialog was raised — or is already on screen for this
 * exact path (the duplicate is swallowed, BUG-1678); a null classification
 * (the failure is invisible to this preflight, e.g. no window yet) stays
 * silent so the caller keeps its existing home-tab fallback either way.
 */
export function reportOpenFailure(filePath: string, deps: OpenFailureDeps): boolean {
  const failure = classifyOpenFailure(filePath)
  if (!failure) return false
  if (activeDialogPaths.has(filePath)) return true
  activeDialogPaths.add(filePath)
  deps.showErrorDialog(deps.openFailedMessage(basename(filePath)), failure, () => {
    activeDialogPaths.delete(filePath)
  })
  return true
}
