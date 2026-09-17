/**
 * The single list of user-facing Sheets keyboard shortcuts.
 *
 * Bindings live where they must fire — the main-process menu accelerators and
 * the Univer shortcut registrations in excel-shortcuts.ts — but every one of
 * them is described here once so the Keyboard Shortcuts dialog reads from one
 * place. Keep this honest: only list chords that are actually bound.
 *
 * Chords are written in Mac notation and rewritten to Ctrl/Alt/Shift form on
 * Windows/Linux, exactly like the docs app's shortcut sheet.
 */
import { macShortcutsToWin, platformShortcuts } from '@airy-office/i18n'
import type { StringKey } from './i18n/locale'

export type ShortcutGroupId = 'file' | 'workbook' | 'selection'

export interface ShortcutDef {
  id: string
  group: ShortcutGroupId
  labelKey: StringKey
  /** Mac notation, e.g. '⇧⌘S'; alternates are listed as 'A / B' */
  keys: string
  /** Windows/Linux chord when it differs from a notation rewrite of `keys` */
  win?: string
}

export const SHORTCUT_GROUPS: readonly { id: ShortcutGroupId; labelKey: StringKey }[] = [
  { id: 'file', labelKey: 'scGroupFile' },
  { id: 'workbook', labelKey: 'scGroupWorkbook' },
  { id: 'selection', labelKey: 'scGroupSelection' },
]

export const SHORTCUTS: readonly ShortcutDef[] = [
  // ---- File (main-process menu accelerators, sent as menu actions) ----
  { id: 'open', group: 'file', labelKey: 'scOpen', keys: '⌘O' },
  { id: 'save', group: 'file', labelKey: 'scSave', keys: '⌘S' },
  { id: 'save-as', group: 'file', labelKey: 'scSaveAs', keys: '⇧⌘S' },
  { id: 'print', group: 'file', labelKey: 'scPrint', keys: '⌘P' },
  { id: 'undo', group: 'file', labelKey: 'scUndo', keys: '⌘Z' },
  { id: 'redo', group: 'file', labelKey: 'scRedo', keys: '⇧⌘Z / ⌘Y' },

  // ---- Workbook navigation (excel-shortcuts.ts) ----
  { id: 'next-sheet', group: 'workbook', labelKey: 'scNextSheet', keys: '⌘PgDn / ⌥→' },
  { id: 'prev-sheet', group: 'workbook', labelKey: 'scPrevSheet', keys: '⌘PgUp / ⌥←' },
  { id: 'sheet-home', group: 'workbook', labelKey: 'scFirstCell', keys: '⌘Home' },
  { id: 'sheet-end', group: 'workbook', labelKey: 'scLastUsedCell', keys: '⌘End' },
  { id: 'row-home', group: 'workbook', labelKey: 'scRowStart', keys: 'Home' },

  // ---- Selection & layout (excel-shortcuts.ts) ----
  // ⌘Space is Spotlight — Excel for mac uses real Ctrl+Space too
  { id: 'select-column', group: 'selection', labelKey: 'scSelectColumn', keys: '⌃Space' },
  { id: 'select-row', group: 'selection', labelKey: 'scSelectRow', keys: '⇧Space' },
  { id: 'hide-rows', group: 'selection', labelKey: 'scHideRows', keys: '⌘9' },
  { id: 'unhide-rows', group: 'selection', labelKey: 'scUnhideRows', keys: '⇧⌘9' },
  { id: 'hide-cols', group: 'selection', labelKey: 'scHideColumns', keys: '⌘0' },
  { id: 'unhide-cols', group: 'selection', labelKey: 'scUnhideColumns', keys: '⇧⌘0' },
]

/** platformShortcuts is the identity on macOS and rewrites the notation elsewhere */
const IS_MAC = platformShortcuts('⌘') === '⌘'

/** the chord as this platform's user sees it */
export function shortcutKeys(def: ShortcutDef, isMac = IS_MAC): string {
  if (isMac) return def.keys
  return def.win ?? macShortcutsToWin(def.keys)
}
