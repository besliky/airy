/**
 * PERF-1671 drag-gesture regression guard.
 *
 * Two contracts live here:
 * 1. The pure helpers the canvas gesture path uses (drag-gesture.ts): identity-stable
 *    guide/spacing state (a dragmove must not re-render the canvas when the snap chrome
 *    content did not change) and the move-release transform math (final position).
 * 2. The gesture→history wiring (source-contract style, like video-export-throttle-restore:
 *    slides-main.ts is an Electron main module that cannot be imported into a unit test).
 *    A move drag commits ONCE — at release — so one whole drag = exactly one history entry;
 *    the first preview frame of a resize/adjust gesture pushes the pre-gesture snapshot,
 *    later previews and the final commit do not.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { dragEndTransform, sameGuides, sameSpacing } from '../src/renderer/drag-gesture'
import type { Guide, SpacingIndicator } from '../src/renderer/snap'

const here = dirname(fileURLToPath(import.meta.url))
const canvasSource = readFileSync(join(here, '../src/renderer/SlideCanvas.tsx'), 'utf8')
const mainSource = readFileSync(join(here, '../src/main/slides-main.ts'), 'utf8')

describe('sameGuides keeps state identity when the chrome content is unchanged', () => {
  it('equal-content fresh arrays compare equal (the dragmove case: no re-render)', () => {
    const prev = [{ axis: 'v' as const, pos: 120 }]
    const next = [{ axis: 'v' as const, pos: 120 }]
    expect(sameGuides(prev, next)).toBe(true)
    // the React state updater must return the PREVIOUS reference so React bails out
    expect((() => (sameGuides(prev, next) ? prev : next))()).toBe(prev)
  })

  it('both empty keeps identity (the typical dragmove: no snap)', () => {
    const prev: Guide[] = []
    const next: Guide[] = []
    expect(sameGuides(prev, next)).toBe(true)
    expect((() => (sameGuides(prev, next) ? prev : next))()).toBe(prev)
  })

  it('different position, axis, or length updates the state', () => {
    expect(sameGuides([{ axis: 'v', pos: 120 }], [{ axis: 'v', pos: 121 }])).toBe(false)
    expect(sameGuides([{ axis: 'v', pos: 120 }], [{ axis: 'h', pos: 120 }])).toBe(false)
    expect(sameGuides([], [{ axis: 'v', pos: 120 }])).toBe(false)
  })
})

describe('sameSpacing keeps state identity when the spacing chrome is unchanged', () => {
  const arr = (a: SpacingIndicator[]) => a

  it('equal-content fresh arrays compare equal', () => {
    const prev = arr([{ axis: 'x', from: 10, to: 30, at: 20 }])
    const next = arr([{ axis: 'x', from: 10, to: 30, at: 20 }])
    expect(sameSpacing(prev, next)).toBe(true)
  })

  it('any component change updates the state', () => {
    const base = arr([{ axis: 'x', from: 10, to: 30, at: 20 }])
    expect(sameSpacing(base, arr([{ axis: 'y', from: 10, to: 30, at: 20 }]))).toBe(false)
    expect(sameSpacing(base, arr([{ axis: 'x', from: 11, to: 30, at: 20 }]))).toBe(false)
    expect(sameSpacing(base, arr([{ axis: 'x', from: 10, to: 31, at: 20 }]))).toBe(false)
    expect(sameSpacing(base, arr([{ axis: 'x', from: 10, to: 30, at: 21 }]))).toBe(false)
    expect(sameSpacing(base, arr([]))).toBe(false)
  })
})

describe('dragEndTransform converts the Konva drop position to the model box', () => {
  const box = { w: 173, h: 115, rotationDeg: 42 }

  it('model top-left is the released center minus half size (boxPivotProps inverse)', () => {
    // Konva group position IS the box center (boxPivotProps); a drag release at
    // (200, 150) must land the model box at (200 - w/2, 150 - h/2)
    expect(dragEndTransform(box, 200, 150)).toEqual({
      x: 200 - 173 / 2,
      y: 150 - 115 / 2,
      w: 173,
      h: 115,
      rotationDeg: 42,
    })
  })

  it('rotation defaults to 0 when the box carries none', () => {
    const out = dragEndTransform({ w: 10, h: 20 }, 5, 10)
    expect(out.rotationDeg).toBe(0)
    expect(out.x).toBe(0)
    expect(out.y).toBe(0)
  })

  it('a move never changes size', () => {
    const out = dragEndTransform(box, 999, -999)
    expect(out.w).toBe(173)
    expect(out.h).toBe(115)
  })
})

describe('the canvas gesture commits exactly once per drag (source contract)', () => {
  it('the dragmove handler never commits to the model', () => {
    // Snap computation, position clamping and guide reporting only; the model
    // commit happens in onDragEnd — this is what makes one whole drag = one
    // history entry (the main process also coalesces preview frames, pinned below).
    const dragMoveBody = canvasSource.slice(
      canvasSource.indexOf('onDragMove: (e: Konva.KonvaEventObject<DragEvent>)'),
      canvasSource.indexOf('onDragEnd: (e: Konva.KonvaEventObject<DragEvent>)'),
    )
    expect(dragMoveBody).not.toContain('onTransform(')
  })

  it('the drag release commits through the tested transform helper', () => {
    const dragEndBody = canvasSource.slice(
      canvasSource.indexOf('onDragEnd: (e: Konva.KonvaEventObject<DragEvent>)'),
      canvasSource.indexOf('onTransformStart:'),
    )
    expect(dragEndBody).toContain('onTransform(')
    expect(dragEndBody).toContain('dragEndTransform(box, e.target.x(), e.target.y())')
  })

  it('guide/spacing state updates route through the identity-stable comparators', () => {
    expect(canvasSource).toContain('sameGuides(prev, g) ? prev : g')
    expect(canvasSource).toContain('sameSpacing(prev, sp ?? []) ? prev : (sp ?? [])')
  })

  it('while a drag is live the slide raster decimates and snaps back at the end', () => {
    expect(canvasSource).toContain(
      'const ratio = dragRaster ? Math.max(base * DRAG_RASTER_FACTOR, DRAG_RASTER_MIN) : base',
    )
  })
})

describe('one whole drag = one undo step (main-process history contract, source-pinned)', () => {
  it('the first preview frame pushes the pre-gesture snapshot; later frames do not', () => {
    expect(mainSource).toContain(
      'Undo semantics for preview gestures: one whole drag = one undo step.',
    )
    expect(mainSource).toContain('if (!session.transformPreview) {')
    expect(mainSource).toContain('pushHistory(session)')
  })

  it('a plain (non-preview) transform commits push their own single history entry', () => {
    expect(mainSource).toContain('} else if (session.transformPreview) {')
    expect(mainSource).toContain('session.transformPreview = false')
  })
})
