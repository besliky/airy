/// Batch-open handling for process launches (BUG-1737): argv, the macOS
/// open-file event and the single-instance forwarding used to open only the
/// FIRST recognized file of the batch — a 6-path launch from a file manager's
/// multi-select produced 1 tab and silently dropped the other 5, because the
/// old supportedFileIn/unsupportedFileIn/missingFileIn helpers `find`-ed a
/// single winner. The helpers here classify EVERY recognized path and open
/// them all, keeping the per-path failure dialogs (#154) and the Home
/// fallback intact. Dependency-injected like dropped-files.ts so the batch
/// semantics stay unit-testable without Electron.
import { existsSync } from 'node:fs'
import { reportOpenFailure, type OpenFailureDeps } from './open-failure'

// ---- routing regexes (the single source of truth for shell open routing) ----

export const DOCX_RE = /\.docx$/i
export const XLSX_RE = /\.(xlsx|xlsm|xls|csv)$/i
export const PPTX_RE = /\.pptx$/i
export const PDF_RE = /\.pdf$/i
export const MD_RE = /\.(md|markdown)$/i
export const HTML_RE = /\.html?$/i

/** document formats we recognize but don't open — surfaced as a dialog, not silently dropped */
export const UNSUPPORTED_DOC_RE = /\.(doc|rtf|odt|ppt|pps|odp|ods|xlsb|pages|key|numbers)$/i

/** union of the per-module open regexes above, built from their sources so the
 *  two cannot drift (unreadable-path guard in routeDocumentPath, launch-argv
 *  classification in collectLaunchFiles) */
export const SUPPORTED_DOC_RE = new RegExp(
  [DOCX_RE, XLSX_RE, PPTX_RE, PDF_RE, MD_RE, HTML_RE].map((re) => re.source).join('|'),
  'i',
)

// ---- launch classification ----

/** every recognized path in a launch, in first-seen order (BUG-1737) */
export interface LaunchFiles {
  /** existing files with a supported extension — each gets its own tab */
  supported: string[]
  /** existing files with a known-but-unsupported extension (.doc/.rtf/...) */
  unsupported: string[]
  /** supported-extension paths that cannot be stat'ed (vanished, EACCES on
   *  the path or a parent). Switch-looking arguments (leading "-") are
   *  ignored, as they always were. */
  missing: string[]
}

export function emptyLaunchFiles(): LaunchFiles {
  return { supported: [], unsupported: [], missing: [] }
}

/** total number of recognized paths — 0 means "this launch names no document" */
export function launchFileCount(files: LaunchFiles): number {
  return files.supported.length + files.unsupported.length + files.missing.length
}

/** flat list in bucket order (supported, then unsupported, then missing);
 *  used for the single-instance forward where a plain array travels well */
export function launchPathList(files: LaunchFiles): string[] {
  return [...files.supported, ...files.unsupported, ...files.missing]
}

/**
 * Classify one path into a batch, exactly as an argv entry would be. Used by
 * collectLaunchFiles and by the macOS open-file event, whose paths never
 * appear in argv (and which can fire several times before ready for a
 * Finder multi-select — they all ride in the same pending batch).
 */
export function addLaunchPath(files: LaunchFiles, filePath: string): void {
  if (SUPPORTED_DOC_RE.test(filePath) && existsSync(filePath)) files.supported.push(filePath)
  else if (UNSUPPORTED_DOC_RE.test(filePath) && existsSync(filePath))
    files.unsupported.push(filePath)
  else if (!filePath.startsWith('-') && SUPPORTED_DOC_RE.test(filePath))
    files.missing.push(filePath)
}

/**
 * Classify EVERY argv entry instead of `find`-ing the first (BUG-1737).
 * Duplicates open once, mirroring the drop-payload dedupe. Unlike the drop
 * payload there is deliberately NO cap: a file-manager multi-select is
 * explicit user intent for every entry, and a silent tail-drop would repeat
 * the exact loss this bug is about (the drop cap is the separate UX-1627
 * question, not a precedent to copy).
 */
export function collectLaunchFiles(argv: string[]): LaunchFiles {
  const files = emptyLaunchFiles()
  const seen = new Set<string>()
  for (const arg of argv) {
    if (seen.has(arg)) continue
    seen.add(arg)
    addLaunchPath(files, arg)
  }
  return files
}

// ---- batch open ----

export interface LaunchOpenDeps {
  /** the single router (openDocumentPath); false = no tab for this path */
  openDocument: (path: string) => boolean
  /** the #154 dialog channel plus its localized title (openFailureDeps) */
  openFailure: OpenFailureDeps
  /** the shared warning box; receives a localized message */
  showWarning: (message: string) => void
  /** localized template for one-or-more unsupported extensions */
  unsupportedMessage: (exts: string[]) => string
}

/**
 * Open every collected launch path (BUG-1737). Each supported path routes
 * through the normal File > Open pipeline and gets its own tab; one that
 * still produces no tab (a directory wearing a document extension, an
 * unreadable file, a path that vanished between classification and open)
 * raises the #154 dialog via reportOpenFailure, as a lone such argument
 * always did. Paths collected as `missing` never reach the router — they
 * are reported through the same channel. Known-unsupported formats get ONE
 * aggregated warning instead of a per-file dialog storm: that is the
 * established batch convention (handleDroppedFiles), and staying silent
 * would repeat the exact loss this bug exists for. Returns true when at
 * least one tab was opened; the caller keeps its Home fallback for the
 * empty result.
 */
export function openLaunchFiles(files: LaunchFiles, deps: LaunchOpenDeps): boolean {
  let opened = false
  const failed: string[] = []
  for (const path of files.supported) {
    if (deps.openDocument(path)) {
      opened = true
    } else {
      failed.push(path)
    }
  }
  failed.push(...files.missing)
  const unsupportedExts = [
    ...new Set(
      files.unsupported.map((path) => path.slice(path.lastIndexOf('.') + 1).toLowerCase()),
    ),
  ]
  reportFailuresThen(failed, unsupportedExts, deps)
  return opened
}

/**
 * Report the batch's failures ONE AT A TIME, then warn. The shell shows a
 * single error box at a time (the BUG-1678 guard in error-dialog.ts): a
 * synchronous burst of reportOpenFailure calls would surface the first
 * dialog and silently drop the rest — the exact loss BUG-1737 is about —
 * so each next path's #154 dialog is raised only after the previous one is
 * dismissed. A call the channel swallows (an unrelated dialog owns the
 * screen) fires onClosed immediately, so the chain then drains without
 * waiting — the same busy-screen semantics the single-path flow always had.
 * A path whose failure the preflight cannot see (or whose dialog the
 * open-failure dedupe swallows as a repeat) never calls the channel, so it
 * advances the chain at once. The aggregated unsupported warning, if any,
 * comes after the last dialog instead of stacking on top of it.
 */
function reportFailuresThen(
  failed: string[],
  unsupportedExts: string[],
  deps: LaunchOpenDeps,
): void {
  const next = (index: number): void => {
    if (index >= failed.length) {
      if (unsupportedExts.length > 0) deps.showWarning(deps.unsupportedMessage(unsupportedExts))
      return
    }
    let channelCalled = false
    reportOpenFailure(failed[index], {
      openFailedMessage: (name) => deps.openFailure.openFailedMessage(name),
      showErrorDialog: (message, err, onClosed) => {
        channelCalled = true
        deps.openFailure.showErrorDialog(message, err, () => {
          onClosed?.()
          next(index + 1)
        })
      },
    })
    if (!channelCalled) next(index + 1)
  }
  next(0)
}
