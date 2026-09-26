import { describe, it, expect } from 'vitest'
import type { ChartModel } from '@airy-office/pptx-engine'
import { buildChartNode } from '../src/build-chart'
import { linearFit, formatTrendEquation, formatRSquared } from '../src/trendline'
import { HeuristicMetrics } from '../src/metrics'
import { makeViewport } from '../src/coords'

const vp = makeViewport({ cx: 12192000, cy: 6858000 }, 1280)
const metrics = new HeuristicMetrics()
const box = {
  x: 100,
  y: 100,
  w: 600,
  h: 400,
  centerX: 400,
  centerY: 300,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
}

describe('linearFit (least squares)', () => {
  it('fits an exact line with r² 1', () => {
    const fit = linearFit([
      { x: 0, y: 100 },
      { x: 1, y: 200 },
      { x: 2, y: 300 },
      { x: 3, y: 400 },
    ])
    expect(fit).toEqual({ slope: 100, intercept: 100, r2: 1 })
  })

  it('fits noisy data (known least-squares solution)', () => {
    // x = 1..4, y = 1, 2, 2, 3 → slope 0.6, intercept 0.5, r² 0.9
    const fit = linearFit([
      { x: 1, y: 1 },
      { x: 2, y: 2 },
      { x: 3, y: 2 },
      { x: 4, y: 3 },
    ])
    expect(fit!.slope).toBeCloseTo(0.6, 12)
    expect(fit!.intercept).toBeCloseTo(0.5, 12)
    expect(fit!.r2).toBeCloseTo(0.9, 12)
  })

  it('handles a constant series (slope 0) and rejects degenerate input', () => {
    const flat = linearFit([
      { x: 0, y: 5 },
      { x: 1, y: 5 },
      { x: 2, y: 5 },
    ])
    expect(flat).toEqual({ slope: 0, intercept: 5, r2: 1 })
    expect(linearFit([{ x: 1, y: 1 }])).toBeNull()
    expect(linearFit([])).toBeNull()
    expect(
      linearFit([
        { x: 2, y: 1 },
        { x: 2, y: 3 },
      ]),
    ).toBeNull()
  })
})

describe('trendline label formatting', () => {
  it('formats equations like PowerPoint (coefficients of 1 dropped, sign split)', () => {
    expect(formatTrendEquation(0.5, 1.5)).toBe('y = 0.5x + 1.5')
    expect(formatTrendEquation(1, 100)).toBe('y = x + 100')
    expect(formatTrendEquation(-1, 4)).toBe('y = -x + 4')
    expect(formatTrendEquation(2.5, -0.75)).toBe('y = 2.5x - 0.75')
    expect(formatTrendEquation(100, 0)).toBe('y = 100x')
  })

  it('formats R² with four decimals', () => {
    expect(formatRSquared(1)).toBe('R² = 1.0000')
    expect(formatRSquared(6 / 7)).toBe('R² = 0.8571')
  })
})

describe('trendline render (chart builders)', () => {
  it('line chart: draws one dashed 2-point trendline polyline + equation/R² labels inside the box', () => {
    const model: ChartModel = {
      kind: 'line',
      categories: ['2016', '2017', '2018', '2019'],
      series: [
        {
          name: 'A',
          color: '#4CAF50',
          values: [100, 200, 300, 400],
          trendlines: [{ kind: 'linear', dispRSqr: true, dispEq: true }],
        },
      ],
    }
    const node = buildChartNode('r_1', 'el1', model, box, vp, metrics)!
    // Series polyline + 1 trendline
    expect(node.polylines).toHaveLength(2)
    const tl = node.polylines.find((p) => p.points.length === 4)
    expect(tl).toBeDefined()
    expect(tl!.dash).toBeDefined()
    expect(tl!.color).toBe('#4CAF50')
    // Labels carry the fit over x = 0..3: y = 100x + 100
    const texts = node.labels.map((l) => l.text)
    expect(texts).toContain('y = 100x + 100')
    expect(texts).toContain('R² = 1.0000')
    // Everything stays inside the chart box
    for (const p of node.polylines) {
      for (let i = 0; i < p.points.length; i += 2) {
        expect(p.points[i]).toBeGreaterThanOrEqual(0)
        expect(p.points[i]).toBeLessThanOrEqual(box.w)
        expect(p.points[i + 1]).toBeGreaterThanOrEqual(0)
        expect(p.points[i + 1]).toBeLessThanOrEqual(box.h)
      }
    }
  })

  it('scatter chart: the trendline spans the real x extent and honors spPr color/width', () => {
    const model: ChartModel = {
      kind: 'scatter',
      categories: [],
      series: [
        {
          name: 'S',
          color: '#E91E63',
          xValues: [1, 2, 3, 4, 5],
          values: [2.1, 3.9, 6.2, 7.8, 10.1],
          trendlines: [{ kind: 'linear', color: '#333333', widthPt: 2, dash: 'sysDash' }],
        },
      ],
    }
    const node = buildChartNode('r_2', 'el2', model, box, vp, metrics)!
    const tl = node.polylines.find((p) => p.color === '#333333')
    expect(tl).toBeDefined()
    expect(tl!.points).toHaveLength(4)
    // Slope ≈ 2 over x 1..5: the line rises left → right
    expect(tl!.points[3]).toBeLessThan(tl!.points[1]!)
    const pxPerX = (tl!.points[2]! - tl!.points[0]!) / 4
    expect(pxPerX).toBeGreaterThan(0)
    // No dispEq/dispRSqr → no equation labels
    expect(node.labels.some((l) => l.text.startsWith('y ='))).toBe(false)
  })

  it('a non-linear-only trendline never reaches the render (dropped at parse)', () => {
    const model: ChartModel = {
      kind: 'bar',
      categories: ['A', 'B', 'C'],
      series: [{ name: 'B', values: [1, 2, 3] }],
    }
    const node = buildChartNode('r_3', 'el3', model, box, vp, metrics)!
    // No trendlines in the model → only bars, no extra polylines
    expect(node.polylines).toHaveLength(0)
    expect(node.bars).toHaveLength(3)
  })
})
