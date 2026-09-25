import { afterAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
  addLaunchPath,
  collectLaunchFiles,
  emptyLaunchFiles,
  launchFileCount,
  launchPathList,
  openLaunchFiles,
  type LaunchOpenDeps,
} from '../src/main/launch-files'
import { resetOpenFailureDedupe } from '../src/main/open-failure'

/**
 * BUG-1737: launch argv (and the single-instance forward) used to open only
 * the FIRST recognized file of a batch — `find` in supportedFileIn /
 * unsupportedFileIn / missingFileIn — so a 6-path multi-select "Open in Airy"
 * produced 1 tab and silently dropped the other 5. collectLaunchFiles /
 * openLaunchFiles are the batch replacement; everything here runs against the
 * real filesystem and the real reportOpenFailure preflight, with only the
 * router/dialog/warning channels injected.
 */

const here = dirname(fileURLToPath(import.meta.url))
// index.ts is an Electron main module that cannot be imported into a unit
// test, so the wiring itself is the contract under regression guard (same
// style as session-restore-cascade.test.ts)
const shellMainSource = readFileSync(join(here, '../src/main/index.ts'), 'utf8')

let root: string

beforeEach(() => {
  // reportOpenFailure keeps module-state dedupe markers; a leaked one from a
  // previous test must never swallow the next test's dialog (BUG-1678)
  resetOpenFailureDedupe()
  root = mkdtempSync(join(tmpdir(), 'bug1737-'))
})

afterAll(() => {
  rmSync(root, { recursive: true, force: true })
})

/** the audit's exact shape: one existing file per supported type */
function sixTypePaths(): string[] {
  return (
    [
      ['doc5.docx', 'x'],
      ['book5.xlsx', 'x'],
      ['deck5.pptx', 'x'],
      ['doc5.pdf', 'x'],
      ['note5.md', 'x'],
      ['page5.html', 'x'],
    ] as const
  ).map(([name, content]) => {
    const path = join(root, name)
    writeFileSync(path, content)
    return path
  })
}

/** two existing supported paths, fresh-named per call */
function twoPaths(): [string, string] {
  const first = join(root, `first-${Math.random().toString(36).slice(2)}.docx`)
  const second = join(root, `second-${Math.random().toString(36).slice(2)}.md`)
  writeFileSync(first, 'x')
  writeFileSync(second, 'x')
  return [first, second]
}

function deps(overrides: Partial<LaunchOpenDeps> = {}): LaunchOpenDeps & {
  opened: string[]
  dialogs: string[]
  warnings: string[]
} {
  const opened: string[] = []
  const dialogs: string[] = []
  const warnings: string[] = []
  return {
    opened,
    dialogs,
    warnings,
    openDocument: vi.fn((path: string) => {
      opened.push(path)
      return true
    }),
    openFailure: {
      // the real reportOpenFailure classifies the path itself; the channel
      // records the raised #154 dialog and auto-dismisses it (onClosed), so
      // the sequential-failure chain keeps moving
      showErrorDialog: vi.fn((message: string, _err: Error, onClosed?: () => void) => {
        dialogs.push(message)
        onClosed?.()
      }),
      openFailedMessage: vi.fn((name: string) => `Could not open “${name}”`),
    },
    showWarning: vi.fn((message: string) => {
      warnings.push(message)
    }),
    unsupportedMessage: (exts) => `unsupported: ${exts.join(', ')}`,
    ...overrides,
  }
}

describe('collectLaunchFiles', () => {
  it('a 6-type batch collects ALL paths, not just the first (BUG-1737)', () => {
    const paths = sixTypePaths()
    const files = collectLaunchFiles(paths)
    expect(files.supported).toEqual(paths)
    expect(launchFileCount(files)).toBe(6)
  })

  it('supported paths keep argv order so tabs open in launch order', () => {
    const paths = sixTypePaths()
    const files = collectLaunchFiles([paths[3], paths[0], paths[5]])
    expect(files.supported).toEqual([paths[3], paths[0], paths[5]])
  })

  it('a mixed batch routes every entry: valid, missing, legacy, junk, flags', () => {
    const [docx, md] = twoPaths()
    const missing = join(root, 'gone.docx')
    const legacy = join(root, 'old.doc')
    writeFileSync(legacy, 'x')
    const png = join(root, 'photo.png')
    writeFileSync(png, 'x')
    const files = collectLaunchFiles(['electron', docx, missing, legacy, '--some-flag', png, md])
    expect(files.supported).toEqual([docx, md])
    expect(files.unsupported).toEqual([legacy])
    expect(files.missing).toEqual([missing])
  })

  it('a vanished supported path counts as missing; a leading-dash argument never does', () => {
    const vanished = join(root, 'vanished.md')
    const files = collectLaunchFiles([vanished, '-unknown-switch.md'])
    expect(files.missing).toEqual([vanished])
  })

  it('duplicates open once', () => {
    const paths = sixTypePaths()
    const files = collectLaunchFiles([paths[0], paths[0], paths[1], paths[0]])
    expect(files.supported).toEqual([paths[0], paths[1]])
  })

  it('a directory wearing a document extension is collected as supported; the router reports it', () => {
    const dir = join(root, 'folder.docx')
    mkdirSync(dir)
    const files = collectLaunchFiles([dir])
    expect(files.supported).toEqual([dir])
  })

  it('an argv without documents collects nothing', () => {
    const files = collectLaunchFiles(['electron', '.', '--flag', '/etc'])
    expect(files).toEqual(emptyLaunchFiles())
  })

  it('addLaunchPath accumulates several macOS open-file events into one batch', () => {
    const paths = sixTypePaths()
    const legacy = join(root, 'legacy.pages')
    writeFileSync(legacy, 'x')
    const files = emptyLaunchFiles()
    for (const p of paths.slice(0, 3)) addLaunchPath(files, p)
    addLaunchPath(files, legacy)
    expect(files.supported).toEqual(paths.slice(0, 3))
    expect(files.unsupported).toEqual([legacy])
    expect(launchPathList(files)).toHaveLength(4)
  })
})

describe('openLaunchFiles', () => {
  it('opens EVERY valid path of the 6-type batch, in order, with no dialogs', () => {
    const paths = sixTypePaths()
    const d = deps()
    expect(openLaunchFiles(collectLaunchFiles(paths), d)).toBe(true)
    expect(d.opened).toEqual(paths)
    expect(d.dialogs).toEqual([])
    expect(d.warnings).toEqual([])
  })

  it('the mixed batch opens all valid paths AND reports each invalid one', () => {
    const [docx, md] = twoPaths()
    const missingXlsx = join(root, 'gone.xlsx')
    const missingPdf = join(root, 'nope.pdf')
    const legacy = join(root, 'ancient.doc')
    writeFileSync(legacy, 'x')
    const d = deps()
    const files = collectLaunchFiles([missingXlsx, docx, legacy, missingPdf, md])
    expect(openLaunchFiles(files, d)).toBe(true)
    expect(d.opened).toEqual([docx, md])
    // one #154 dialog per unopenable path — not just the first
    expect(d.dialogs).toEqual(['Could not open “gone.xlsx”', 'Could not open “nope.pdf”'])
    // known-unsupported formats share ONE aggregated warning (drop convention)
    expect(d.warnings).toEqual(['unsupported: doc'])
  })

  it('a supported path the router cannot open (directory) raises its own #154 dialog', () => {
    const dir = join(root, 'folder.pptx')
    mkdirSync(dir)
    const d = deps({ openDocument: vi.fn(() => false) })
    expect(openLaunchFiles(collectLaunchFiles([dir]), d)).toBe(false)
    expect(d.dialogs).toEqual(['Could not open “folder.pptx”'])
    expect(d.warnings).toEqual([])
  })

  it('a router failure invisible to the preflight stays silent; Home fallback covers it', () => {
    // a readable file whose open failed for another reason (no window yet):
    // the real reportOpenFailure classification passes → no dialog
    const file = join(root, 'fine.md')
    writeFileSync(file, 'x')
    const d = deps({ openDocument: vi.fn(() => false) })
    expect(openLaunchFiles(collectLaunchFiles([file]), d)).toBe(false)
    expect(d.dialogs).toEqual([])
  })

  it('an all-invalid batch explains every path and opens nothing', () => {
    // legacy files must EXIST to classify as unsupported: a vanished .doc is
    // not collectable at all (same as the pre-batch single-path behavior)
    const legacyDoc = join(root, 'a.doc')
    const legacyRtf = join(root, 'b.rtf')
    writeFileSync(legacyDoc, 'x')
    writeFileSync(legacyRtf, 'x')
    const d = deps()
    const files = collectLaunchFiles([legacyDoc, legacyRtf, join(root, 'gone.docx')])
    expect(openLaunchFiles(files, d)).toBe(false)
    expect(d.opened).toEqual([])
    expect(d.dialogs).toEqual(['Could not open “gone.docx”'])
    expect(d.warnings).toEqual(['unsupported: doc, rtf'])
  })

  it('several bad paths report ONE AT A TIME: no dialog before the previous closes', () => {
    // the shell's error channel is single-dialog (BUG-1678): a synchronous
    // burst would surface the first failure and silently drop the rest, so
    // openLaunchFiles chains each dialog on the previous one's dismissal
    const first = join(root, 'first-gone.docx')
    const second = join(root, 'second-gone.pdf')
    const dialogs: string[] = []
    const showErrorDialog = vi.fn<(message: string, err: Error, onClosed?: () => void) => void>(
      (message) => {
        dialogs.push(message) // record but never call onClosed: dialog stays up
      },
    )
    openLaunchFiles(collectLaunchFiles([first, second]), {
      ...deps(),
      openFailure: { showErrorDialog, openFailedMessage: (name) => `Could not open “${name}”` },
    })
    expect(dialogs).toEqual(['Could not open “first-gone.docx”'])
    // dismissing the first dialog surfaces the second
    const onClosed = showErrorDialog.mock.calls[0][2]
    onClosed?.()
    expect(dialogs).toEqual([
      'Could not open “first-gone.docx”',
      'Could not open “second-gone.pdf”',
    ])
  })

  describe('one-file cases behave exactly as before the batch fix', () => {
    it('a single valid path opens one tab, no dialogs', () => {
      const [only] = sixTypePaths()
      const d = deps()
      expect(openLaunchFiles(collectLaunchFiles([only]), d)).toBe(true)
      expect(d.opened).toEqual([only])
      expect(d.dialogs).toEqual([])
      expect(d.warnings).toEqual([])
    })

    it('a single vanished path raises exactly one #154 dialog (BUG-1655 semantics)', () => {
      const gone = join(root, 'vanished.docx')
      const d = deps()
      expect(openLaunchFiles(collectLaunchFiles([gone]), d)).toBe(false)
      expect(d.dialogs).toEqual(['Could not open “vanished.docx”'])
      expect(d.warnings).toEqual([])
    })

    it('a single legacy format raises exactly one unsupported warning', () => {
      const legacy = join(root, 'old.odt')
      writeFileSync(legacy, 'x')
      const d = deps()
      expect(openLaunchFiles(collectLaunchFiles([legacy]), d)).toBe(false)
      expect(d.opened).toEqual([])
      expect(d.warnings).toEqual(['unsupported: odt'])
      expect(d.dialogs).toEqual([])
    })
  })
})

describe('index.ts wiring (BUG-1737)', () => {
  it('launch, second-instance, open-file and the lock payload all use the batch', () => {
    expect(shellMainSource).toContain('let pendingLaunchFiles = collectLaunchFiles(process.argv)')
    expect(shellMainSource).toContain('const files = collectLaunchFiles(argv)')
    expect(shellMainSource).toContain('openLaunchFiles(files, launchOpenDeps)')
    expect(shellMainSource).toContain('openLaunchFiles(pendingLaunchFiles, launchOpenDeps)')
    expect(shellMainSource).toContain('addLaunchPath(pendingLaunchFiles, filePath)')
    // the batch rides the single-instance lock; the single-path key stays for
    // a version-skewed surviving instance
    expect(shellMainSource).toContain('{ launchPaths: paths, launchPath: paths[0] }')
  })

  it('the first-file-wins helpers are gone', () => {
    expect(shellMainSource).not.toContain('function supportedFileIn')
    expect(shellMainSource).not.toContain('function unsupportedFileIn')
    expect(shellMainSource).not.toContain('function missingFileIn')
    expect(shellMainSource).not.toContain('pendingLaunchPath')
  })

  it('the #154 and Home channels stay wired for the single-path callers (#154/#180)', () => {
    // open-file (macOS) keeps its per-path failure report + Home fallback;
    // Home clicks keep going through openFromHome → reportOpenFailure
    expect(shellMainSource).toContain('reportOpenFailure(filePath, openFailureDeps)')
    expect(shellMainSource).toContain('openFromHome(')
  })
})
