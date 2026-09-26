import { describe, it, expect } from 'vitest'
import type { ChartModel } from '@airy-office/pptx-engine'
import { buildChartNode } from '../src/build-chart'
import { HeuristicMetrics } from '../src/metrics'
import { makeViewport } from '../src/coords'

const vp = makeViewport({ cx: 12192000, cy: 6858000 }, 1280)
const metrics = new HeuristicMetrics()
const box = {
  x: 0,
  y: 0,
  w: 600,
  h: 400,
  centerX: 300,
  centerY: 200,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
}

// Default 10% split over [70, 15, 5, 3, 2] (total 95, threshold 9.5): from the end
// 2 (acc 2), then 3 (acc 5), then 5 would make 10 > 9.5 → break. So [3, 2] = 5 move
// to the secondary; the main pie keeps 70 + 15 + 5 plus one aggregated 5 wedge.
const model: ChartModel = {
  kind: 'pieOfPie',
  categories: ['A', 'B', 'C', 'D', 'E'],
  series: [{ name: 'S', values: [70, 15, 5, 3, 2] }],
  pieOfPie: { type: 'pie', splitPos: 10, secondPieSizePct: 75, gapWidthPct: 100 },
}

describe('buildPieOfPieNode', () => {
  it('splits trailing small slices: 3 main + 1 aggregated + 2 secondary wedges, 2 connectors', () => {
    const node = buildChartNode('r_1', 'el1', model, box, vp, metrics)!
    expect(node.wedges).toHaveLength(6)
    // All wedges stay inside the box
    for (const w of node.wedges!) {
      expect(w.cx - w.outerR).toBeGreaterThanOrEqual(-1)
      expect(w.cx + w.outerR).toBeLessThanOrEqual(box.w + 1)
      expect(w.cy - w.outerR).toBeGreaterThanOrEqual(-1)
      expect(w.cy + w.outerR).toBeLessThanOrEqual(box.h + 1)
    }
    // The aggregated "other" wedge: shade(slice 3's palette color #FFC000, 0.72) = #b88a00
    const other = node.wedges!.find((w) => w.color === '#b88a00')
    expect(other).toBeDefined()
    expect(other!.sweepDeg).toBeCloseTo((5 / 95) * 360, 5)
    // Connector lines from the main wedge boundaries to the secondary pie
    expect(node.axisLines).toHaveLength(2)
  })

  it('bar secondary (ofPieType="bar"): one stacked bar instead of the secondary pie, no connectors', () => {
    const node = buildChartNode(
      'r_2',
      'el2',
      { ...model, pieOfPie: { ...model.pieOfPie!, type: 'bar' } },
      box,
      vp,
      metrics,
    )!
    // Main pie: 3 slices + aggregated wedge; the secondary is a stacked bar now
    expect(node.wedges).toHaveLength(4)
    expect(node.bars).toHaveLength(2)
    // Stacked top-down, no gaps between segments
    const sorted = [...node.bars].sort((a, b) => a.y - b.y)
    expect(sorted[0]!.y + sorted[0]!.h).toBeCloseTo(sorted[1]!.y, 5)
    expect(node.axisLines).toHaveLength(0)
  })

  it('no splittable slices falls back to a plain pie', () => {
    const node = buildChartNode(
      'r_3',
      'el3',
      // 20 > 10% of 110: nothing moves to the secondary → plain-pie fallback
      { kind: 'pieOfPie', categories: ['A', 'B'], series: [{ values: [90, 20] }] },
      box,
      vp,
      metrics,
    )!
    expect(node.wedges).toHaveLength(2)
    expect(node.axisLines).toHaveLength(0)
  })

  it('legend lists every category; percent labels appear with dataLabels', () => {
    const withLegend: ChartModel = {
      ...model,
      legendPos: 'b',
      dataLabels: true,
      dataLabelsPct: true,
    }
    const node = buildChartNode('r_4', 'el4', withLegend, box, vp, metrics)!
    expect(node.swatches).toHaveLength(5)
    const texts = node.labels.map((l) => l.text)
    expect(texts).toContain('A')
    expect(texts).toContain('E')
    // Percent labels of the grand total on drawn wedges
    expect(texts).toContain('74%')
    expect(texts).toContain('2%')
  })
})
