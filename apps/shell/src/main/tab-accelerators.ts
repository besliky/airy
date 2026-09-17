import type { TabKind } from '../shared/tabs-api'

/**
 * Ctrl/Cmd+1..8 → tab N, Ctrl/Cmd+9 → last tab (Home counts as tab 1, like
 * Chrome). Digit switching is wired two ways in the shell: a "Select Tab"
 * submenu in the shell-built menus, and a before-input-event hook that covers
 * the editor-owned menus (docs/sheets/slides). Digits an editor reserves for
 * its own Word/Excel-parity shortcuts are excluded per active tab kind so the
 * application never eats them:
 * - docs: Ctrl+1/2/5 line spacing (Word), Ctrl+8 formatting marks
 * - sheets: Ctrl+9 hide rows (Excel); (Ctrl+0 hide columns is not bound here)
 */
export const RESERVED_TAB_DIGITS: Partial<Record<TabKind, ReadonlySet<number>>> = {
  docs: new Set([1, 2, 5, 8]),
  sheets: new Set([9]),
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
