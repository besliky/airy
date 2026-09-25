/**
 * BUG-1726 regression guard: zoom-state coherence between the applied CSS
 * scale and the app's fit logic.
 *
 * The canvas drag converts the pointer delta to page coordinates through the
 * CSS transform that is actually applied to the stage (Konva normalizes the
 * pointer by rect.width/clientWidth). Cursor and shape therefore stay in sync
 * exactly as long as the applied scale is stable under the gesture. The zoom
 * gestures apply to the CSS transform immediately and commit the state
 * debounced; the stage-wrap ResizeObserver used to decide its re-fit from the
 * LAGGING state zoom, so a scrollbar ripple in that window slammed the canvas
 * back to fit — wiping the user's zoom and, mid-drag, moving the page under
 * the pointer (model delta overshot the cursor by canvasW/pageW = ×1.25 on a
 * 16:9 deck; other gestures never started because the shape left the pointer).
 * These tests pin the fixed decision path: applied-scale reads everywhere.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { appliedStageZoom, resizeFitDecision } from '../src/renderer/zoom-fit'
import { dragEndTransform } from '../src/renderer/drag-gesture'

const here = dirname(fileURLToPath(import.meta.url))
const appSource = readFileSync(join(here, '../src/renderer/App.tsx'), 'utf8')

/** Fake measured stage element: layout width (offsetWidth) + transformed width */
const elAtZoom = (layoutWidth: number, zoom: number) =>
  ({
    offsetWidth: layoutWidth,
    getBoundingClientRect: () => ({ width: layoutWidth * zoom }),
  }) as unknown as HTMLDivElement

describe('appliedStageZoom measures the scale the canvas really displays', () => {
  const LAYOUT = 1600 // 1280px page + 2×160 Konva bleed

  it('equals rect.width / offsetWidth at every zoom (the ratio all pointer deltas go through)', () => {
    // 0.4622 / 0.5788 = the audit's stale-vs-applied pair (cursor ×1.25 desync);
    // 1.24 = a manual zoom-in — the same measurement must hold at all of them
    for (const zoom of [0.4622, 0.5788, 1.24]) {
      expect(appliedStageZoom(elAtZoom(LAYOUT, zoom))).toBeCloseTo(zoom, 10)
    }
  })

  it('returns null before the element is measurable (and for a null element)', () => {
    expect(appliedStageZoom(null)).toBeNull()
    expect(appliedStageZoom(elAtZoom(0, 1))).toBeNull()
  })
})

describe('resizeFitDecision never reads the lagging state zoom (BUG-1726)', () => {
  const FIT = 0.6414
  const base = {
    stateZoom: FIT, // what the state/refs still say during a gesture's commit window
    rawFitZoom: FIT,
    fitZoom: FIT,
    lastFitZoom: FIT, // the app opened in fit mode
    outerResized: false,
  }

  it('a zoom-in inside its commit window is not stomped back to fit (scrollbar ripple)', () => {
    // probe repro: applied 0.74, state still 0.64 → the old code saw "fit mode"
    // and re-fit, silently wiping the zoom ~300ms after every zoom step
    const d = resizeFitDecision({ ...base, appliedZoom: 0.7414 })
    expect(d).toEqual({ action: 'none' })
  })

  it('a manual zoom below fit is left alone (no bounce back to fit)', () => {
    // probe repro: slider to 0.55, state still 0.64 → old code re-fit to 0.6414,
    // moving the page under the next mousedown (drag never started)
    const d = resizeFitDecision({ ...base, appliedZoom: 0.55 })
    expect(d).toEqual({ action: 'none' })
  })

  it('fit mode still follows a real container resize', () => {
    expect(resizeFitDecision({ ...base, appliedZoom: FIT, outerResized: true })).toEqual({
      action: 'refit',
      zoom: FIT,
    })
    expect(resizeFitDecision({ ...base, appliedZoom: FIT, outerResized: false })).toEqual({
      action: 'refit',
      zoom: FIT,
    })
  })

  it('a manual zoom above fit is clamped only by a real (border-box) shrink', () => {
    const manual = { ...base, lastFitZoom: FIT, appliedZoom: 1.2 }
    // scrollbar appearance is not a resize: keep the user's zoom
    expect(resizeFitDecision(manual)).toEqual({ action: 'none' })
    // a genuine pane/window shrink clamps back to fit
    expect(resizeFitDecision({ ...manual, outerResized: true })).toEqual({
      action: 'refit',
      zoom: FIT,
    })
  })

  it('a manual zoom above fit that still fits the viewport is never wiped', () => {
    const d = resizeFitDecision({
      ...base,
      appliedZoom: 1.2,
      rawFitZoom: 1.4,
      outerResized: true,
    })
    expect(d).toEqual({ action: 'none' })
  })

  it('pre-mount (nothing measurable) falls back to the state zoom', () => {
    expect(
      resizeFitDecision({ ...base, appliedZoom: null, stateZoom: FIT, outerResized: true }),
    ).toEqual({ action: 'refit', zoom: FIT })
  })
})

describe('the drag delta converter holds the exact screen↔page proportion at every zoom', () => {
  const box = { w: 173, h: 115, rotationDeg: 0 }
  const SCREEN = 100

  it('a 100-screen-px drag commits exactly 100/zoom page px (dragEndTransform chain)', () => {
    for (const zoom of [0.4622, 0.5788, 1.24]) {
      const applied = appliedStageZoom(elAtZoom(1600, zoom))!
      const konvaDropX = 500 + SCREEN / applied // Konva receives the pointer in stage px…
      const drop = dragEndTransform(box, konvaDropX, 300)
      const start = dragEndTransform(box, 500, 300)
      expect(drop.x - start.x).toBeCloseTo(SCREEN / zoom, 9)
      // …and the same screen distance is slideDelta × applied scale on screen
      expect((drop.x - start.x) * applied).toBeCloseTo(SCREEN, 9)
    }
  })
})

describe('source contract: fit decisions and the live zoom ref read the applied scale', () => {
  it('the resize observer decides via resizeFitDecision + appliedStageZoom', () => {
    expect(appSource).toContain('const decision = resizeFitDecision({')
    expect(appSource).toContain('appliedZoom: appliedStageZoom(stageScaleRef.current),')
    // the stale pattern must not come back
    expect(appSource).not.toContain('Math.abs(zoomLiveRef.current - lf)')
  })

  it('previewZoom syncs the live ref with the applied scale at apply time', () => {
    expect(appSource).toContain('zoomLiveRef.current = g.pending')
  })
})
