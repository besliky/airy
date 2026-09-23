/**
 * BUG-1654: the html save handler shares the markdown app's in-place write
 * path, so it must run the same staleness fence: before overwriting the
 * document's own path, compare the on-disk mtime+size against the stamp taken
 * at open / last save, and on a mismatch prompt Overwrite / Save As / Cancel
 * instead of silently forking a renamed file or winning last-writer-wins
 * against a second window. The fence mechanics are behavior-tested for the
 * markdown twin (apps/markdown/tests/save-staleness.test.ts) and unit-tested
 * in packages/electron-utils/tests/save-staleness.test.ts; this pins the html
 * wiring.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const source = readFileSync(join(here, '../src/main/html-main.ts'), 'utf8')
const renderer = readFileSync(join(here, '../src/renderer/App.tsx'), 'utf8')

describe('html save staleness fence wiring', () => {
  it('the save handler checks staleness before an in-place write', () => {
    expect(source).toContain('checkSaveStaleness(target, saveStampByWc.get(e.sender.id))')
    expect(source).toContain('if (request.auto === true) return done({ ok: true, canceled: true })')
  })

  it('the fence is scoped to in-place saves only', () => {
    expect(source).toContain('if (pathAtRequest && resolve(pathAtRequest) === resolve(target))')
  })

  it('a successful save refreshes the baseline stamp', () => {
    expect(source).toContain('saveStampByWc.set(e.sender.id, statFileStamp(target))')
  })

  it('open and rename keep the baseline in sync with the tracked path', () => {
    expect(source).toContain('saveStampByWc.set(wcId, statFileStamp(openPath))')
    expect(source).toContain('saveStampByWc.set(wcId, statFileStamp(newPath))')
    expect(source).toContain('saveStampByWc.delete(wcId)')
  })

  it('the external-change prompt offers Save As / Overwrite / Cancel', () => {
    expect(source).toContain("buttons: [tm('btnSaveAs'), tm('btnOverwrite'), tm('btnCancel')]")
  })

  it('autosave marks its saves as automatic so the fence declines without a modal', () => {
    expect(renderer).toContain("void doSave('save', undefined, true)")
  })
})
