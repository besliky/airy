/**
 * BUG-1107 follow-up: the five runtime openGeneratedPath wrappers (sheets,
 * slides, pdf, markdown, html exports) used to route the generated file into
 * whichever shell window held FOCUS — an export from a background tab of an
 * unfocused window "teleported" across windows. The asking view's
 * webContents id now rides along every openGeneratedPath call, and the shell
 * resolves the sender's window first (managerForSender), keeping the focused
 * window only as the sender-less fallback (menu/standalone paths).
 *
 * Source-wiring test (same style as export-atomic-write.test.ts): these are
 * Electron main modules that cannot be imported into a unit test.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const read = (app: string, rel: string): string =>
  readFileSync(join(here, '../../', app, rel), 'utf8')

const HOOK_SIGNATURE = 'openGeneratedPath?: (path: string, senderWcId?: number) => boolean'

describe('openGeneratedPath sender routing (BUG-1107 follow-up)', () => {
  it('every runtime hook type accepts the asking view webContents id', () => {
    for (const source of [
      read('sheets', 'src/main/sheets-main.ts'),
      read('slides', 'src/main/session-state.ts'),
      read('pdf', 'src/main/pdf-main.ts'),
      read('markdown', 'src/main/markdown-main.ts'),
      read('html', 'src/main/html-main.ts'),
    ]) {
      expect(source).toContain(HOOK_SIGNATURE)
    }
  })

  it('each wrapper forwards the sender id to the shell router', () => {
    for (const source of [
      read('sheets', 'src/main/sheets-main.ts'),
      read('slides', 'src/main/slides-main.ts'),
      read('pdf', 'src/main/pdf-main.ts'),
      read('markdown', 'src/main/markdown-main.ts'),
      read('html', 'src/main/html-main.ts'),
    ]) {
      expect(source).toContain('runtime.openGeneratedPath?.(path, senderWcId)')
    }
  })

  it('IPC handlers pass the asking view id, not a focused-window guess', () => {
    expect(read('markdown', 'src/main/markdown-main.ts')).toContain(
      'openExportedPdf(picked.filePath, e.sender.id)',
    )
    expect(read('html', 'src/main/html-main.ts')).toContain(
      'openExportedPdf(picked.filePath, e.sender.id)',
    )
    expect(read('sheets', 'src/main/sheets-main.ts')).toContain(
      'openGeneratedFile(result.path, event.sender.id)',
    )
    expect(read('sheets', 'src/main/sheets-main.ts')).toContain(
      'openGeneratedFile(filePath, event.sender.id)',
    )
    expect(read('pdf', 'src/main/pdf-main.ts')).toContain(
      'openGeneratedPdf(targetPath, e.sender.id)',
    )
    // the slides export goes through exportSlidesPdf's callback option
    expect(read('slides', 'src/main/slides-main.ts')).toContain(
      'openExportedPdf: (path) => openExportedPdf(path, e.sender.id)',
    )
  })

  it('the shell resolves the sender window first for all five runtimes', () => {
    const shell = read('shell', 'src/main/index.ts')
    // five configure*Runtime wirings here, plus the docs shell hook from the
    // original BUG-1107 fix (openGeneratedPath with the same signature)
    expect(shell.match(/openGeneratedPath: \(path, senderWcId\) =>/g)?.length).toBe(6)
    expect(shell).toContain('openGeneratedDocument(path, managerForSender(senderWcId))')
    expect(shell).not.toContain('openGeneratedPath: (path) => openGeneratedDocument(path)')
  })
})
