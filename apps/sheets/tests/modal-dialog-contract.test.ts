import { describe, expect, it } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'

/**
 * UX-1101 modal contract (sheets twin of the docs contract behind UX-1001):
 * every overlay dialog in the sheets renderer adopts the shared
 * useModalDialog hook (packages/ui/src/modal-dialog.ts) — dialog role /
 * aria-modal / aria-labelledby, focus trap, focus return and a stopped
 * Escape. The docs contract only scans apps/docs, so the PR #67 dialogs
 * (Text to Columns, Outline Settings) shipped in sheets without even an
 * Escape handler; this source-level test is the anti-recurrence gate for
 * that class here.
 *
 * Mechanical trait — the one artifact every adoption shares and nothing else
 * uses:
 *   - a modal ROOT is a JSX element carrying the literal attribute
 *     `className="dialog-backdrop"` or `className="modal-backdrop"` (the two
 *     dimming wrappers sheets dialogs render; menu-dismiss layers like
 *     chart-menu-backdrop are not dialogs and do not count);
 *   - an ADOPTED modal spreads the hook's props onto that root, spelled
 *     `{...dialog.backdropProps}` (the controller variable is `dialog` by
 *     the convention every existing adoption follows).
 *
 * Rule, per .tsx file under src/renderer (recursive):
 *   backdrops - spreads === LEGACY[file] ?? 0
 *
 * LEGACY is the acknowledged pre-UX-1101 debt — sheets started with a large
 * registry of one-off dialogs. It is a ratchet: migrating a legacy dialog to
 * the hook means lowering its number in the same commit — never raising a
 * number, never adding files. A brand-new dialog therefore MUST adopt
 * useModalDialog to keep the balance at zero in every file not listed here.
 */
const SRC = join(__dirname, '../src/renderer')

/** modal roots still on the legacy keyboard-only behavior (no hook) */
const LEGACY: Record<string, number> = {
  // full-load filter prompt
  'App.tsx': 1,
  // AdvancedFilterDialog
  'AdvancedFilterDialog.tsx': 1,
  // AllowEditRangesDialog
  'AllowEditRangesDialog.tsx': 1,
  // chart title/data panels
  'ChartPanels.tsx': 1,
  // ConsolidateDialog
  'ConsolidateDialog.tsx': 1,
  // EquationDialog
  'EquationDialog.tsx': 1,
  // SortDialog + RemoveDuplicatesDialog
  'ExcelShell.tsx': 2,
  // FormatCellsDialog
  'FormatCellsDialog.tsx': 1,
  // GoalSeekDialog
  'GoalSeekDialog.tsx': 1,
  // GoToDialog
  'GoToDialog.tsx': 1,
  // HeaderFooterDialog
  'HeaderFooterDialog.tsx': 1,
  // IconsDialog
  'IconsDialog.tsx': 1,
  // InsertFunctionDialog
  'InsertFunctionDialog.tsx': 1,
  // NameManagerDialog
  'NameManagerDialog.tsx': 1,
  // PivotDialog
  'PivotDialog.tsx': 1,
  // RecommendedChartsDialog
  'RecommendedChartsDialog.tsx': 1,
  // RecoveryDialog
  'RecoveryDialog.tsx': 1,
  // ScreenshotDialog
  'ScreenshotDialog.tsx': 1,
  // SlicerPanel
  'SlicerPanel.tsx': 1,
  // SubtotalDialog
  'SubtotalDialog.tsx': 1,
  // SymbolDialog
  'SymbolDialog.tsx': 1,
  // TimelinePanel
  'TimelinePanel.tsx': 1,
  // PrintDialog (modal-backdrop root)
  'PrintDialog.tsx': 1,
  // ShortcutsDialog (modal-backdrop root, useModalKeys)
  'ShortcutsDialog.tsx': 1,
}

const occurrences = (src: string, needle: string): number => src.split(needle).length - 1

const backdrops = (src: string): number =>
  occurrences(src, 'className="dialog-backdrop') + occurrences(src, 'className="modal-backdrop')

/** every .tsx under the renderer, relative to SRC (dialogs live in the renderer root) */
function tsxFiles(dir: string, prefix = ''): string[] {
  const out: string[] = []
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.isDirectory()) out.push(...tsxFiles(join(dir, entry.name), `${prefix}${entry.name}/`))
    else if (entry.name.endsWith('.tsx')) out.push(prefix + entry.name)
  }
  return out
}

describe('overlay dialogs adopt useModalDialog (UX-1101)', () => {
  const files = tsxFiles(SRC)

  it('finds the renderer tree (scanner sanity)', () => {
    expect(files.length).toBeGreaterThan(20)
  })

  it.each(files)('%s keeps every dialog backdrop on the hook', (file) => {
    const src = readFileSync(join(SRC, file), 'utf8')
    const adopted = occurrences(src, '{...dialog.backdropProps}')
    const legacy = LEGACY[file] ?? 0
    // equality (not <=) keeps LEGACY self-tightening: retire a legacy dialog
    // and its allowance must shrink in the same commit
    expect(backdrops(src) - adopted).toBe(legacy)
  })

  it('lists only files that still render dialog backdrops', () => {
    for (const [file, legacy] of Object.entries(LEGACY)) {
      const src = readFileSync(join(SRC, file), 'utf8')
      expect(backdrops(src), `${file} entry is dead`).toBeGreaterThan(0)
      expect(legacy, `${file} allowance must be positive`).toBeGreaterThan(0)
    }
  })
})
