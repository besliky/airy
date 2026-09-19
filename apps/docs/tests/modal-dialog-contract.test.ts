import { describe, expect, it } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'

/**
 * UX-1001 modal contract: every overlay dialog in the docs renderer adopts
 * the shared useModalDialog hook (packages/ui/src/modal-dialog.ts) — dialog
 * role/aria-modal/aria-labelledby, focus trap, focus return and a stopped
 * Escape. This source-level test is the anti-recurrence gate for that class
 * (UX-904 closed it for one dialog; seven more shipped unhooked in v0.12.0).
 *
 * Mechanical trait — the one artifact every adoption shares and nothing else
 * uses:
 *   - a modal ROOT is a JSX element carrying the literal attribute
 *     `className="modal-backdrop"` (every overlay dialog in components/
 *     renders exactly this class on its dimming wrapper);
 *   - an ADOPTED modal spreads the hook's props onto that root, spelled
 *     `{...dialog.backdropProps}` (the controller variable is `dialog` by
 *     the convention every existing adoption follows).
 *
 * Rule, per .tsx file under src/renderer/components:
 *   backdrops - spreads === LEGACY[file] ?? 0
 *
 * LEGACY is the acknowledged pre-UX-1001 debt (the useModalKeys-era
 * dialogs), each entry naming its dialogs and how many are still unhooked.
 * It is a ratchet: migrating a legacy dialog to the hook means lowering its
 * number in the same commit — never raising a number, never adding files.
 * A brand-new dialog therefore MUST adopt useModalDialog to keep the balance
 * at zero in every file not listed here.
 */
const SRC = join(__dirname, '../src/renderer/components')

/** modal roots still on the legacy keyboard-only behavior (no hook) */
const LEGACY: Record<string, number> = {
  // FontDialog + ParagraphDialog
  'ContextMenu.tsx': 2,
  'EquationModal.tsx': 1,
  // shared margin fields dialog (ribbonMargin* labels)
  'MarginDialog.tsx': 1,
  'PasswordDialog.tsx': 1,
  'PrintDialog.tsx': 1,
  'PromptModal.tsx': 1,
  'ProtectDialog.tsx': 1,
  // BookmarkModal + TableInsertModal
  'ribbon-insert-tab.tsx': 2,
  // CaptionModal + SourceModal
  'ribbon-references-tab.tsx': 2,
  // TablePropertiesDialog + ListDefineDialog
  'Ribbon.tsx': 2,
  'ShortcutsDialog.tsx': 1,
  'WordCountDialog.tsx': 1,
}

const occurrences = (src: string, needle: string): number => src.split(needle).length - 1

describe('overlay dialogs adopt useModalDialog (UX-1001)', () => {
  const files = readdirSync(SRC).filter((f) => f.endsWith('.tsx'))

  it('finds the component directory (scanner sanity)', () => {
    expect(files.length).toBeGreaterThan(10)
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
