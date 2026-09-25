/**
 * BUG-1724 source-contract pins for the slides save staleness fence, in the
 * style of apps/html/tests/save-staleness-wiring.test.ts. The fence behavior
 * itself is covered end-to-end by tests/save-fence.test.ts; this file pins the
 * wiring that behavior tests cannot observe: the rename re-stamp, the renderer
 * autosave tick marking its saves as automatic, and the preload flag encoding.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const main = readFileSync(join(here, '../src/main/slides-main.ts'), 'utf8')
const renderer = readFileSync(join(here, '../src/renderer/App.tsx'), 'utf8')
const preload = readFileSync(join(here, '../src/preload/index.ts'), 'utf8')
const fileActions = readFileSync(join(here, '../src/renderer/file-actions.ts'), 'utf8')

describe('slides save staleness fence wiring (BUG-1724)', () => {
  it('the save handler checks staleness before an in-place write', () => {
    expect(main).toContain("checkSaveStaleness(session.path, session.saveStamp) !== 'fresh'")
  })

  it('automatic saves are declined without a dialog and stay dirty', () => {
    expect(main).toContain("if (auto === true) return { ok: false, reason: 'external-modified' }")
  })

  it('the external-change prompt offers Save As (default) / Overwrite / Cancel', () => {
    expect(main).toContain("buttons: [tm('btnSaveAs'), tm('btnOverwrite'), tm('btnCancel')]")
    expect(main).toContain('defaultId: 0,')
    expect(main).toContain('cancelId: 2,')
  })

  it('every successful write refreshes the baseline stamp', () => {
    // in-place save, Save As (shared by the menu and the fence), AI draft write
    expect(main).toContain('session.saveStamp = statFileStamp(session.path)')
    expect(main).toContain('session.saveStamp = statFileStamp(r.filePath)')
    expect(main).toContain('session.saveStamp = statFileStamp(draftPath)')
  })

  it('open stamps the session and a rename re-stamps at the new path', () => {
    expect(main).toContain('saveStamp: statFileStamp(path),')
    expect(main).toContain('session.saveStamp = statFileStamp(newPath)')
  })

  it('the fence runs only after the untitled drafts branch picked a fresh path', () => {
    const saveHandler = main.slice(
      main.indexOf("ipcMain.handle('slides:save'"),
      main.indexOf("ipcMain.handle('slides:save-as'"),
    )
    expect(saveHandler).toContain('pickDraftPath')
    expect(saveHandler.indexOf('pickDraftPath')).toBeLessThan(
      saveHandler.indexOf('checkSaveStaleness'),
    )
  })

  it('the autosave tick marks its saves as automatic so the fence declines without a modal', () => {
    expect(renderer).toContain('void save(true, true)')
  })

  it('a manual save (⌘S / close guard / menu) never passes the auto flag', () => {
    // the close-guard "Save" path and plain save() calls stay auto=false, so a
    // stale file still raises the Save As / Overwrite / Cancel dialog there
    expect(renderer).not.toContain('save(false, true)')
  })

  it('the preload forwards a strict boolean auto flag over the slides:save channel', () => {
    expect(preload).toContain("ipcRenderer.invoke('slides:save', auto === true)")
  })

  it('the renderer suppresses the error toast for the by-design external-modified refusal', () => {
    expect(fileActions).toContain("r.reason === 'external-modified'")
    expect(fileActions).toContain('window.slidesApi.save(auto)')
  })
})
