import { describe, expect, it } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'

/**
 * UX-1101 modal contract (slides twin of the docs contract behind UX-1001):
 * every overlay dialog in the slides renderer adopts the shared
 * useModalDialog hook (packages/ui/src/modal-dialog.ts) — dialog role /
 * aria-modal / aria-labelledby, focus trap, focus return and a stopped
 * Escape. The docs contract only scans apps/docs, so PR #62/#67-class
 * dialogs shipped in slides without the hook; this source-level test is the
 * anti-recurrence gate for that class here.
 *
 * Mechanical trait — the one artifact every adoption shares and nothing else
 * uses:
 *   - a modal ROOT is a JSX element carrying the literal attribute
 *     `className="modal-backdrop"` (every overlay dialog in the renderer
 *     renders exactly this class on its dimming wrapper);
 *   - an ADOPTED modal spreads the hook's props onto that root, spelled
 *     `{...dialog.backdropProps}` (the controller variable is `dialog` by
 *     the convention every existing adoption follows).
 *
 * Rule, per .tsx file under src/renderer (recursive):
 *   backdrops - spreads === LEGACY[file] ?? 0
 *
 * LEGACY is the acknowledged pre-UX-1101 debt (the useModalKeys-era
 * dialogs), each entry naming its dialogs and how many are still unhooked.
 * It is a ratchet: migrating a legacy dialog to the hook means lowering its
 * number in the same commit — never raising a number, never adding files.
 * A brand-new dialog therefore MUST adopt useModalDialog to keep the balance
 * at zero in every file not listed here.
 */
const SRC = join(__dirname, '../src/renderer')

/** modal roots still on the legacy keyboard-only behavior (no hook) */
const LEGACY: Record<string, number> = {
  // rehearse-discard confirm
  'App.tsx': 1,
  // ChartDataDialog
  'components/ChartDataDialog.tsx': 1,
  // ChartTypeDialog
  'components/ChartTypeDialog.tsx': 1,
  // CustomShowDialog
  'components/CustomShowDialog.tsx': 1,
  // CutoutDialog
  'components/CutoutDialog.tsx': 1,
  // LinkDialog + HeaderFooterDialog + EquationDialog
  'components/InsertDialogs.tsx': 3,
  // PrintDialog
  'components/PrintDialog.tsx': 1,
  // ShortcutsDialog (useModalKeys)
  'components/ShortcutsDialog.tsx': 1,
}

const occurrences = (src: string, needle: string): number => src.split(needle).length - 1

/** every .tsx under the renderer, relative to SRC (dialogs live in both the root and components/) */
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

  it.each(files)('%s keeps every modal-backdrop on the hook', (file) => {
    const src = readFileSync(join(SRC, file), 'utf8')
    const backdrops = occurrences(src, 'className="modal-backdrop')
    const adopted = occurrences(src, '{...dialog.backdropProps}')
    const legacy = LEGACY[file] ?? 0
    // equality (not <=) keeps LEGACY self-tightening: retire a legacy dialog
    // and its allowance must shrink in the same commit
    expect(backdrops - adopted).toBe(legacy)
  })

  it('lists only files that still render modal backdrops', () => {
    for (const [file, legacy] of Object.entries(LEGACY)) {
      const src = readFileSync(join(SRC, file), 'utf8')
      expect(
        occurrences(src, 'className="modal-backdrop'),
        `${file} entry is dead`,
      ).toBeGreaterThan(0)
      expect(legacy, `${file} allowance must be positive`).toBeGreaterThan(0)
    }
  })
})
