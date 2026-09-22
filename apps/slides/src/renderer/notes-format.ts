/**
 * Whole-body notes formatting helpers (PAR-316). The notes editor is plain
 * text, so formatting applies to the entire notes body — PowerPoint's
 * select-all + toggle semantics. The merge is pure so the pane wiring in
 * App.tsx stays thin and the behavior is unit-testable.
 */
import type { NotesFormat } from '../shared/ipc'
import { DEFAULT_NOTES_PT, NOTES_PT_MAX, NOTES_PT_MIN } from './app-constants'

/** A format patch from the notes toolbar: explicit toggles or a size step. */
export interface NotesFormatPatch {
  bold?: boolean
  italic?: boolean
  /** Font-size step in points (+/−1 from the toolbar's A−/A+ buttons) */
  sizeDelta?: number
}

/**
 * Merge a toolbar patch into the current whole-body format. Absent patch keys
 * keep the current value; size steps clamp to [NOTES_PT_MIN, NOTES_PT_MAX] and
 * a step back to the default size drops the explicit size again.
 */
export function nextNotesFormat(cur: NotesFormat, patch: NotesFormatPatch): NotesFormat {
  const next: NotesFormat = {}
  // Only true is stored: absence means "off" on the wire (the engine writes no
  // b/i attribute for false), so toggling off removes the key again.
  if (patch.bold) next.bold = true
  else if (!('bold' in patch) && cur.bold) next.bold = cur.bold
  if (patch.italic) next.italic = true
  else if (!('italic' in patch) && cur.italic) next.italic = cur.italic

  const curPt = cur.fontSizePt ?? DEFAULT_NOTES_PT
  if (patch.sizeDelta !== undefined) {
    const pt = Math.min(NOTES_PT_MAX, Math.max(NOTES_PT_MIN, curPt + patch.sizeDelta))
    if (pt !== DEFAULT_NOTES_PT) next.fontSizePt = pt
  } else if (cur.fontSizePt !== undefined) {
    next.fontSizePt = cur.fontSizePt
  }
  return next
}

/** The size shown in the pane's size readout (default when unset). */
export function notesPtOr(cur: NotesFormat, fallback = DEFAULT_NOTES_PT): number {
  return cur.fontSizePt ?? fallback
}
