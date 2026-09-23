import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const read = (rel: string) => readFileSync(join(here, '..', rel), 'utf8')

const appSrc = read('src/renderer/App.tsx')
const shellSrc = read('src/renderer/ExcelShell.tsx')

/** The File-tab dropdown's items, one JSX <button> block per entry. */
const fileMenuItems = (): string[] => {
  const start = shellSrc.indexOf('<div className="file-menu">')
  const menu = shellSrc.slice(start, shellSrc.indexOf('</div>', start))
  return menu.match(/<button[\s\S]*?<\/button>/g) ?? []
}

const menuItem = (labelKey: string): string => {
  const item = fileMenuItems().find((block) => block.includes(`t('${labelKey}')`))
  expect(item, `File menu item ${labelKey}`).toBeDefined()
  return item as string
}

/**
 * OBS-1605: the File-tab dropdown gated "Print…" on `canSave` (pending edits
 * > 0), so a freshly opened, unmodified workbook had a dead Print item while
 * the native application menu's Ctrl+P worked. Printing now has its own
 * `canPrint` gate, deliberately independent of the save journal; the save,
 * save-as and export guards stay exactly as they were.
 */
describe('File menu print gate (OBS-1605)', () => {
  it('gates File › Print on canPrint, never on the save journal', () => {
    const print = menuItem('appFilePrint')
    expect(print).toContain('disabled={!canPrint}')
    expect(print).not.toContain('canSave')
    expect(print).toMatch(/onOpenPrintDialog\(\)/)
    // The shell contract keeps print availability a separate prop.
    expect(shellSrc).toMatch(/readonly canPrint: boolean/)
  })

  it('prints with no pending edits and in dirty mode alike', () => {
    // App passes the gate without any reference to pendingEdits — the same
    // ungated policy the native application menu's Ctrl+P has always had.
    expect(appSrc).toMatch(/^\s+canPrint$/m)
    // The save gate it was wrongly sharing is untouched.
    expect(appSrc).toContain('canSave={pendingEdits > 0}')
  })

  it('leaves the save guards exactly as they were', () => {
    expect(menuItem('appFileSave')).toContain('disabled={!canSave}')
    expect(menuItem('appFileSaveAs')).toContain('disabled={!canSaveAs}')
    // The QAT save button still follows the journal (save-as-toolbar pins it).
    expect(shellSrc).toMatch(/disabled=\{!canSave\}[\s\S]{0,80}<SaveIcon \/>/)
    expect(appSrc).toContain('canSaveAs={workbookFile !== null}')
  })

  it('keeps the export items on the save guard (unchanged behavior)', () => {
    expect(menuItem('appFileExportPdf')).toContain('disabled={!canSave}')
    expect(menuItem('appFileExportCsv')).toContain('disabled={!canSave}')
  })
})
