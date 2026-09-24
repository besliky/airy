import type { WebContents } from 'electron'
import type { TabManager } from './tab-manager'

/**
 * Per-editor "this file was renamed on disk" notifications. Each embedded
 * editor syncs its internal save path and title bar when its tab's file is
 * renamed from the Home list (pdf tabs have no entry — a pdf rename is driven
 * by the pdf module's own setPdfRenamedHook instead).
 */
export interface FileRenamedHooks {
  docs: (wc: WebContents, oldPath: string, newPath: string) => void
  sheets: (wc: WebContents, oldPath: string, newPath: string) => void
  slides: (wc: WebContents, oldPath: string, newPath: string) => void
  markdown: (wc: WebContents, oldPath: string, newPath: string) => void
  html: (wc: WebContents, oldPath: string, newPath: string) => void
}

/**
 * A Home rename moved a file on disk: re-point every matching open tab in ALL
 * shell windows, then notify each affected tab's editor. The file's identity
 * changed globally, so routing through the focused window alone would leave a
 * second window's tab on the stale title/path — its Ctrl+S would write to the
 * pre-rename name (BUG-1675).
 */
export function renameFileInAllWindows(
  managers: readonly Pick<TabManager, 'renameTabFile'>[],
  oldPath: string,
  newPath: string,
  hooks: FileRenamedHooks,
): void {
  for (const manager of managers) {
    for (const tab of manager.renameTabFile(oldPath, newPath)) {
      if (tab.kind === 'slides') hooks.slides(tab.webContents, oldPath, newPath)
      else if (tab.kind === 'docs') hooks.docs(tab.webContents, oldPath, newPath)
      else if (tab.kind === 'sheets') hooks.sheets(tab.webContents, oldPath, newPath)
      else if (tab.kind === 'markdown') hooks.markdown(tab.webContents, oldPath, newPath)
      else if (tab.kind === 'html') hooks.html(tab.webContents, oldPath, newPath)
    }
  }
}
