/**
 * The single list of user-facing Slides keyboard shortcuts.
 *
 * Bindings live where they must fire — keyboard-actions.ts's global handler
 * and the main-process menu accelerators — but every one of them is described
 * here once so the Keyboard Shortcuts dialog reads from one place. Keep this
 * honest: only list chords that are actually bound there.
 *
 * Chords are written in Mac notation and rewritten to Ctrl/Alt/Shift form on
 * Windows/Linux, exactly like the docs app's shortcut sheet.
 */
import { macShortcutsToWin, platformShortcuts } from '@airy-office/i18n'
import type { StringKey } from './i18n/locale'

export type ShortcutGroupId = 'show' | 'edit' | 'clipboard' | 'arrange'

export interface ShortcutDef {
  id: string
  group: ShortcutGroupId
  labelKey: StringKey
  /** Mac notation, e.g. '⇧⌘S'; alternates are listed as 'A / B' */
  keys: string
  /** Windows/Linux chord when it differs from a notation rewrite of `keys` */
  win?: string
  /** Binding exists on macOS only (hidden from the sheet elsewhere) */
  macOnly?: boolean
}

export const SHORTCUT_GROUPS: readonly { id: ShortcutGroupId; labelKey: StringKey }[] = [
  { id: 'show', labelKey: 'scGroupShow' },
  { id: 'edit', labelKey: 'scGroupEdit' },
  { id: 'clipboard', labelKey: 'scGroupClipboard' },
  { id: 'arrange', labelKey: 'scGroupArrange' },
]

export const SHORTCUTS: readonly ShortcutDef[] = [
  // ---- Slide show (keyboard-actions.ts) ----
  { id: 'show-start', group: 'show', labelKey: 'scShowFromStart', keys: 'F5' },
  { id: 'show-current', group: 'show', labelKey: 'scShowFromCurrent', keys: '⇧F5' },
  // PowerPoint for macOS: ⌘+Enter starts from the current slide
  {
    id: 'show-current-mac',
    group: 'show',
    labelKey: 'scShowFromCurrent',
    keys: '⌘⏎',
    macOnly: true,
  },

  // ---- Editing (keyboard-actions.ts + menu accelerators) ----
  // label reuses the context menu's "New Slide" string
  { id: 'new-slide', group: 'edit', labelKey: 'appCtxNewSlide', keys: '⌘M' },
  { id: 'open', group: 'edit', labelKey: 'scOpen', keys: '⌘O' },
  { id: 'save', group: 'edit', labelKey: 'scSave', keys: '⌘S' },
  { id: 'save-as', group: 'edit', labelKey: 'scSaveAs', keys: '⇧⌘S' },
  { id: 'print', group: 'edit', labelKey: 'scPrint', keys: '⌘P' },
  { id: 'undo', group: 'edit', labelKey: 'scUndo', keys: '⌘Z' },
  { id: 'redo', group: 'edit', labelKey: 'scRedo', keys: '⇧⌘Z / ⌘Y' },
  { id: 'find', group: 'edit', labelKey: 'scFind', keys: '⌘F' },
  { id: 'ask-ai', group: 'edit', labelKey: 'scAskSelection', keys: '⌘K' },
  { id: 'zoom-in', group: 'edit', labelKey: 'scZoomIn', keys: '⌘=' },
  { id: 'zoom-out', group: 'edit', labelKey: 'scZoomOut', keys: '⌘-' },
  { id: 'zoom-reset', group: 'edit', labelKey: 'scZoomReset', keys: '⌘0' },
  { id: 'select-all', group: 'edit', labelKey: 'scSelectAll', keys: '⌘A' },

  // ---- Clipboard & duplication (keyboard-actions.ts) ----
  { id: 'copy', group: 'clipboard', labelKey: 'scCopy', keys: '⌘C' },
  { id: 'cut', group: 'clipboard', labelKey: 'scCut', keys: '⌘X' },
  { id: 'paste', group: 'clipboard', labelKey: 'scPaste', keys: '⌘V' },
  { id: 'copy-format', group: 'clipboard', labelKey: 'scCopyFormat', keys: '⇧⌘C' },
  { id: 'paste-format', group: 'clipboard', labelKey: 'scPasteFormat', keys: '⇧⌘V' },
  { id: 'duplicate', group: 'clipboard', labelKey: 'scDuplicate', keys: '⌘D' },

  // ---- Arrange & selection (keyboard-actions.ts) ----
  { id: 'group', group: 'arrange', labelKey: 'scGroup', keys: '⌘G' },
  { id: 'ungroup', group: 'arrange', labelKey: 'scUngroup', keys: '⇧⌘G' },
  { id: 'delete', group: 'arrange', labelKey: 'scDelete', keys: '⌫ / ⌦' },
  { id: 'nudge', group: 'arrange', labelKey: 'scNudge', keys: '← ↑ → ↓' },
  { id: 'nudge-far', group: 'arrange', labelKey: 'scNudgeFar', keys: '⇧← ↑ → ↓' },
  {
    id: 'cycle-selection',
    group: 'arrange',
    labelKey: 'scCycleSelection',
    keys: '⇥ / ⇧⇥',
    win: 'Tab / Shift+Tab',
  },
  // nothing selected: arrows/PageUp/PageDown switch slides (thumbnail-pane behavior)
  { id: 'switch-slide', group: 'arrange', labelKey: 'scSwitchSlide', keys: '↑ ↓ / PgUp PgDn' },
  // Esc drops the ink pen, exits format painter and in-group editing
  { id: 'escape', group: 'arrange', labelKey: 'scEscapeTools', keys: 'Esc' },
]

/** platformShortcuts is the identity on macOS and rewrites the notation elsewhere */
const IS_MAC = platformShortcuts('⌘') === '⌘'

/** the chord as this platform's user sees it */
export function shortcutKeys(def: ShortcutDef, isMac = IS_MAC): string {
  if (isMac) return def.keys
  return def.win ?? macShortcutsToWin(def.keys)
}
