/**
 * Double-click entry into a group (BUG-1725 / SL-EDIT-1): which child the
 * gesture targets and whether that child's text editing opens in the same
 * gesture (PowerPoint-style). Pure geometry/decision helpers — the Konva event
 * plumbing lives in SlideCanvas, the state wiring in App's onEnterGroup.
 */
import type { GroupRenderNode, RenderNode } from '@airy-office/pptx-render'
import { isEditableText } from './konva-adapter'

/** The box fields that decide whether the DOM text overlay can sit over a child. */
export interface OverlayAlignmentBox {
  rotationDeg?: number
  flipH?: boolean
  flipV?: boolean
}

/**
 * The DOM text overlay doesn't follow group rotation/flip; in those cases
 * children can't double-click into text editing (the same gate the entered-mode
 * render uses). ext/chExt scaling is already baked into child geometry
 * (including text layout), so it does not affect overlay alignment.
 */
export function groupAllowsChildTextEdit(box: OverlayAlignmentBox): boolean {
  return !box.rotationDeg && !box.flipH && !box.flipV
}

/** The resolution of a double-click on a not-yet-entered group. */
export interface GroupDblClickTarget {
  /**
   * Topmost child whose box contains the point (group-local px), null when the
   * point misses every child (group padding) or the pointer position is unknown.
   */
  childId: string | null
  /** Open the text-edit overlay for the hit child in the same gesture. */
  editText: boolean
}

/**
 * Resolve a double-click on a group at a group-local point. Children are
 * tested topmost-first (reverse render order): a point inside both the group's
 * background rectangle and a text child resolves to the text child, so the
 * text-edit overlay can never bind to the shape behind it (the SL-EDIT-1
 * finding where typed text landed in the group's background auto shape).
 */
export function resolveGroupDblClick(
  group: Pick<GroupRenderNode, 'children'>,
  allowsChildTextEdit: boolean,
  local: { x: number; y: number } | null,
): GroupDblClickTarget {
  if (!local) return { childId: null, editText: false }
  const { x, y } = local
  const hit = [...group.children]
    .reverse()
    .find(
      (c: RenderNode) =>
        x >= c.box.x && x <= c.box.x + c.box.w && y >= c.box.y && y <= c.box.y + c.box.h,
    )
  if (!hit) return { childId: null, editText: false }
  return { childId: hit.sourceId, editText: allowsChildTextEdit && isEditableText(hit) }
}
