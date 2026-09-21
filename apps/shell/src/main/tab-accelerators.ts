import type { TabKind } from '../shared/tabs-api'

/**
 * Ctrl/Cmd+1..8 → tab N, Ctrl/Cmd+9 → last tab (Home counts as tab 1, like
 * Chrome). Digit switching is wired two ways in the shell: a "Select Tab"
 * submenu in the shell-built menus, and a before-input-event hook that covers
 * the editor-owned menus (docs/sheets/slides). Digits an editor reserves for
 * its own Word/Excel-parity shortcuts are excluded per active tab kind so the
 * application never eats them:
 * - docs: Ctrl+1/2/5 line spacing (Word), Ctrl+8 formatting marks
 * - sheets: Ctrl+1 Format Cells (Excel), Ctrl+8 outline symbols (Excel),
 *   Ctrl+9 hide rows (Excel); (Ctrl+0 hide columns is not bound here)
 */
export const RESERVED_TAB_DIGITS: Partial<Record<TabKind, ReadonlySet<number>>> = {
  docs: new Set([1, 2, 5, 8]),
  sheets: new Set([1, 8, 9]),
}

/** which switch-to-tab digits the menu may bind for a given active tab kind */
export function switchableDigitsForKind(kind: TabKind): number[] {
  const reserved = RESERVED_TAB_DIGITS[kind]
  return reserved
    ? [1, 2, 3, 4, 5, 6, 7, 8, 9].filter((d) => !reserved.has(d))
    : [1, 2, 3, 4, 5, 6, 7, 8, 9]
}

/** strip index for a switch digit: 1..8 → digit-1, 9 → last; null when invalid */
export function tabIndexForDigit(digit: number, tabCount: number): number | null {
  if (!Number.isInteger(digit) || digit < 1 || digit > 9) return null
  if (tabCount <= 0) return null
  if (digit === 9) return tabCount - 1
  return digit - 1 < tabCount ? digit - 1 : null
}

/** minimal shape of Electron's Input for the pure tests */
interface InputLike {
  type: string
  control: boolean
  meta: boolean
  alt: boolean
  shift: boolean
  code: string
}

/**
 * The switch digit a keydown carries, or null. Ctrl (Win/Linux) or Cmd (mac),
 * no Shift (sheets' Ctrl+Shift+9 unhide) and no Alt (docs' Alt+Cmd heading
 * styles, macOS special characters).
 */
export function switchDigitFromInput(input: InputLike): number | null {
  if (input.type !== 'keyDown') return null
  if (!(input.control || input.meta)) return null
  if (input.alt || input.shift) return null
  const match = /^Digit([1-9])$/.exec(input.code)
  return match ? Number(match[1]) : null
}

/** minimal tab shape the switch decision needs (TabManager.list() satisfies it) */
export interface TabSwitchTabLike {
  id: string
  kind: TabKind
  active: boolean
}

/**
 * Whether a keydown is the "move the active tab to a new window" chord:
 * Ctrl/Cmd+Shift+K, no Alt. The chord was picked against every editor's
 * keydown map (docs reserves Ctrl+M±Shift, ⌘/Ctrl+T±Shift, ⇧⌘E/G;
 * sheets Ctrl+Y) so the shell never eats an editor shortcut. The one
 * external mnemonic it collides with — Word's small caps — is handled by
 * MOVE_TAB_TO_NEW_WINDOW_RESERVED_KINDS on top of this predicate.
 */
export function isMoveTabToNewWindowInput(input: InputLike): boolean {
  return (
    input.type === 'keyDown' &&
    (input.control || input.meta) &&
    input.shift &&
    !input.alt &&
    input.code === 'KeyK'
  )
}

/**
 * Tab kinds whose editor owns Ctrl/Cmd+Shift+K by external mnemonic: Word
 * uses the chord for small caps and Airy-docs renders that formatting (the
 * docs toggle is not bound yet), so while a docs tab is active the shell
 * must not eat the chord — Word muscle memory would otherwise rip the
 * document into a new window instead of formatting text. Same reservation
 * shape as RESERVED_TAB_DIGITS: the chord stays inert until docs binds its
 * small-caps toggle, at which point it starts working with no shell change.
 */
export const MOVE_TAB_TO_NEW_WINDOW_RESERVED_KINDS: ReadonlySet<TabKind> = new Set(['docs'])

/**
 * The move-to-new-window decision for a keydown while the tab of `kind` is
 * active: the chord minus the kinds that reserve it. Shared by the Home
 * hook and the per-view hook in TabManager, so the reservation behaves the
 * same no matter which view owns keyboard focus (kind undefined — no active
 * tab, never in practice — keeps the chord).
 */
export function isMoveTabToNewWindowInputForKind(
  input: InputLike,
  kind: TabKind | undefined,
): boolean {
  if (!isMoveTabToNewWindowInput(input)) return false
  return kind === undefined || !MOVE_TAB_TO_NEW_WINDOW_RESERVED_KINDS.has(kind)
}

/**
 * The tab id a before-input-event keydown should switch to, or null when the
 * input is not a switch chord, the active tab's kind reserves the digit for an
 * editor shortcut, or the digit points beyond the strip. The single decision
 * shared by the hook on the shell's Home webContents and the hook TabManager
 * attaches to every editor view, so digit switching behaves identically no
 * matter which view owns keyboard focus.
 */
export function tabSwitchTargetForInput(input: InputLike, tabs: TabSwitchTabLike[]): string | null {
  const digit = switchDigitFromInput(input)
  if (digit === null) return null
  const activeKind = tabs.find((t) => t.active)?.kind
  if (activeKind && RESERVED_TAB_DIGITS[activeKind]?.has(digit)) return null
  const index = tabIndexForDigit(digit, tabs.length)
  return index === null ? null : (tabs[index].id ?? null)
}
