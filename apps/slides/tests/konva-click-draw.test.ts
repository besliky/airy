/**
 * PERF-1727 regression guard: a plain selection click must not end in Konva's
 * unconditional whole-layer redraw (DD._endDragBefore draws the candidate's layer
 * even when the drag never moved), while a real drag keeps the original end-draw.
 * Drives the real konva DD module with fake drag elements and a draw spy.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { DD } from 'konva/lib/DragAndDrop'

import { suppressClickEndDragDraw } from '../src/renderer/konva-click-draw'

const here = dirname(fileURLToPath(import.meta.url))
const canvasSource = readFileSync(join(here, '../src/renderer/SlideCanvas.tsx'), 'utf8')
const clickDrawSource = readFileSync(join(here, '../src/renderer/konva-click-draw.ts'), 'utf8')

interface FakeNode {
  getStage(): unknown
  getLayer(): unknown
}

function fakeNode(layer: { draw: () => void } | null): FakeNode {
  return {
    getStage: () => ({
      setPointersPositions: () => undefined,
      _changedPointerPositions: [{ id: 1, x: 5, y: 5 }],
    }),
    getLayer: () => layer,
  }
}

function addCandidate(id: number, dragStatus: 'ready' | 'dragging' | 'stopped', node: FakeNode) {
  DD._dragElements.set(id, {
    node: node as never,
    startPointerPos: { x: 0, y: 0 },
    offset: { x: 0, y: 0 },
    pointerId: 1,
    dragStatus,
  })
}

afterEach(() => {
  DD._dragElements.clear()
})

describe('suppressClickEndDragDraw (PERF-1727)', () => {
  it('replaces DD._endDragBefore exactly once', () => {
    const before = DD._endDragBefore
    suppressClickEndDragDraw()
    expect(DD._endDragBefore).not.toBe(before)
    const afterFirst = DD._endDragBefore
    suppressClickEndDragDraw()
    expect(DD._endDragBefore).toBe(afterFirst)
  })

  it('a click-style end (candidate still "ready") does not redraw the layer', () => {
    suppressClickEndDragDraw()
    const draw = vi.fn()
    addCandidate(1, 'ready', fakeNode({ draw }))
    DD._endDragBefore({})
    // no draw: nothing moved, so the full content-layer rasterization is pure waste
    expect(draw).not.toHaveBeenCalled()
  })

  it('a real drag end (candidate "dragging") keeps the original end-draw and bookkeeping', () => {
    suppressClickEndDragDraw()
    const draw = vi.fn()
    addCandidate(1, 'dragging', fakeNode({ draw }))
    DD._endDragBefore({})
    // the dropped node relies on this draw to snap its imperatively-set position
    expect(draw).toHaveBeenCalledTimes(1)
    expect(DD.justDragged).toBe(true)
    expect(DD._dragElements.get(1)?.dragStatus).toBe('stopped')
  })

  it('mixed candidates: a real drag delegates to the original (draws every candidate layer)', () => {
    suppressClickEndDragDraw()
    const drawDragged = vi.fn()
    const drawReady = vi.fn()
    addCandidate(1, 'dragging', fakeNode({ draw: drawDragged }))
    addCandidate(2, 'ready', fakeNode({ draw: drawReady }))
    DD._endDragBefore({})
    // the original draws every candidate's layer once a real drag is in flight
    expect(drawDragged).toHaveBeenCalledTimes(1)
    expect(drawReady).toHaveBeenCalledTimes(1)

    DD._dragElements.clear()
    addCandidate(3, 'ready', fakeNode({ draw: drawReady }))
    DD._endDragBefore({})
    expect(drawReady).toHaveBeenCalledTimes(1)
  })
})

describe('SlideCanvas keeps zoom-only changes off the content-layer raster (PERF-1727)', () => {
  it('the click-patch is wired at SlideCanvas module scope', () => {
    expect(canvasSource).toContain('suppressClickEndDragDraw()')
  })

  it('dragDistance is applied imperatively with auto-draw suppressed, never through JSX', () => {
    // a JSX attr re-apply on a zoom-only change (dock-open refit) would
    // _requestDraw() every node group and batch a full content-layer redraw
    expect(canvasSource).not.toMatch(/dragDistance:\s*6\s*\/\s*Math\.max\(zoom/)
    expect(canvasSource).toContain('setDragDistanceNoRedraw(g, 6 / Math.max(zoom, 0.1))')
    // the suppression lives in the helper (see konva-click-draw.ts)
    expect(clickDrawSource).toContain('Konva.autoDrawEnabled = false')
    expect(clickDrawSource).toContain('node.dragDistance(px)')
    expect(clickDrawSource).toContain('Konva.autoDrawEnabled = autoDraw')
  })

  it('the background cache skips same-input re-bakes', () => {
    expect(canvasSource).toContain('bgBakedRef.current = { ratio, deps }')
    expect(canvasSource).toMatch(/baked\.ratio === ratio && baked\.deps\.every/)
  })
})
