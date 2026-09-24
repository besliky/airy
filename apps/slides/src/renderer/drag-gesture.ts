/**
 * Pure helpers for canvas drag gestures (PERF-1671).
 *
 * During a node drag Konva fires `dragmove` per pointer event. Any state update
 * with a fresh array identity re-renders the whole SlideCanvas (every NodeView)
 * and re-batches the Konva layers, even though the drag chrome (snap guides,
 * spacing indicators) changes content only on the rare moves where an
 * alignment actually snaps. Value-comparing keeps the state identity stable so
 * React bails out and a typical drag move costs no React work at all.
 */
import type { Guide, SpacingIndicator } from './snap'

/** True when two guide lists are content-equal (order-sensitive, cheap). */
export function sameGuides(a: Guide[], b: Guide[]): boolean {
  if (a === b) return true
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) {
    if (a[i].axis !== b[i].axis || a[i].pos !== b[i].pos) return false
  }
  return true
}

/** True when two spacing-indicator lists are content-equal (order-sensitive, cheap). */
export function sameSpacing(a: SpacingIndicator[], b: SpacingIndicator[]): boolean {
  if (a === b) return true
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) {
    if (a[i].axis !== b[i].axis) return false
    if (a[i].from !== b[i].from || a[i].to !== b[i].to || a[i].at !== b[i].at) return false
  }
  return true
}

/**
 * Transform payload for a move-drag release: the Konva group position is the
 * box CENTER (boxPivotProps), so the model top-left is position − half size.
 * Size and rotation are unchanged by a pure move.
 */
export function dragEndTransform(
  box: { w: number; h: number; rotationDeg?: number },
  posX: number,
  posY: number,
): { x: number; y: number; w: number; h: number; rotationDeg: number } {
  return {
    x: posX - box.w / 2,
    y: posY - box.h / 2,
    w: box.w,
    h: box.h,
    rotationDeg: box.rotationDeg ?? 0,
  }
}
