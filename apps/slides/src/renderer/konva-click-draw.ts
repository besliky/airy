/**
 * PERF-1727: Konva's drag-end path redraws the dragged node's whole layer
 * synchronously on every mouseup/touchend over a draggable node — even when the
 * drag never moved, i.e. on a plain selection click. DD._endDragBefore pushes the
 * candidate's layer into its draw list regardless of dragStatus, so every click
 * paid a full content-layer rasterization: ~1.1s on a 100-shape slide under
 * software rasterization (56-70ms at 5 shapes), all of it blocking the main
 * thread before React could commit the selection and update the Format Pane.
 *
 * This patch keeps Konva's semantics for real drags (candidates in the
 * "dragging"/"stopped" state get the exact original bookkeeping and the end-draw
 * the dropped node relies on to snap its position) and skips only the click-time
 * draw, where no node position changed and nothing needs redrawing. Candidate
 * cleanup is untouched: it lives in DD._endDragAfter, a separate window listener.
 *
 * Konva binds _endDragBefore to the window by reference at module load, so the
 * listeners are rebound to the wrapper here; Node.stopDrag resolves the handler
 * dynamically through DD, hence the property is replaced too.
 */
import { DD } from 'konva/lib/DragAndDrop'
// The namespace comes from the entry-free Global module on purpose: a value
// import of bare 'konva' resolves to konva's node entry under vitest, which
// requires the optional native 'canvas' package (absent in this workspace).
import { Konva } from 'konva/lib/Global'

let applied = false

export function suppressClickEndDragDraw(): void {
  if (applied || typeof window === 'undefined') return
  applied = true
  const original = DD._endDragBefore
  const wrapped = (evt?: unknown) => {
    let dragged = false
    DD._dragElements.forEach((elem) => {
      if (elem.dragStatus === 'dragging' || elem.dragStatus === 'stopped') dragged = true
    })
    // A candidate that never left "ready" (click without a drag gesture) has an
    // unchanged position: redraw would rasterize the full layer for no visual change.
    if (dragged) original(evt)
  }
  for (const type of ['mouseup', 'touchend', 'touchcancel'] as const) {
    window.removeEventListener(type, original, true)
    window.addEventListener(type, wrapped, true)
  }
  DD._endDragBefore = wrapped
}

/**
 * PERF-1727: applies dragDistance without Konva's auto-draw. Konva's attr setter
 * calls _requestDraw() unless Konva.autoDrawEnabled is false, so a zoom-only
 * re-apply across every node group (e.g. the auto-refit when a dock opens) would
 * batch a full content-layer redraw (~1.1s of software rasterization at 100
 * shapes). Konva reads dragDistance live at gesture start (DD._drag), so
 * suppressing the auto-draw around the setter is sufficient.
 */
export function setDragDistanceNoRedraw(
  node: { dragDistance(px: number): void },
  px: number,
): void {
  const autoDraw = Konva.autoDrawEnabled
  Konva.autoDrawEnabled = false
  try {
    node.dragDistance(px)
  } finally {
    Konva.autoDrawEnabled = autoDraw
  }
}
