/**
 * PAR-304: the render model exposes cNvPr alt text (title/description) on
 * alt-text-bearing nodes, keeps the descr payload slots (ink/3D) locked, and
 * flags app-created charts via the new extLst marker.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  addChart,
  addElement,
  addPicture,
  openPptx,
  setElementAltText,
  solidPng,
} from '@airy-office/pptx-engine'
import { buildRenderSlide } from '../src/index'
import type { RenderNode } from '../src/render-tree'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) =>
  readFileSync(join(here, '..', '..', 'pptx-engine', 'tests', 'fixtures', name))

const OFF = { x: 914400, y: 914400, cx: 4572000, cy: 2743200 }

async function deckWithAlt() {
  const opened = await openPptx(fx('01_standard_business.pptx'))
  let slide = opened.deck.slides[0]!
  const box = addElement(slide, {
    kind: 'textbox',
    offset: { ...OFF },
    paragraphs: [{ runs: [{ text: 'Alt' }] }],
  })
  setElementAltText(slide, box.id, { title: 'Title text', descr: 'Description text' })
  addChart(opened, 0, {
    kind: 'bar',
    categories: ['a'],
    series: [{ name: 's', values: [1] }],
    offset: { ...OFF },
  })
  // addChart reparses the slide: re-resolve before further edits (ids regenerate)
  slide = opened.deck.slides[0]!
  const box2 = slide.elements.find((e) => e.descr === 'Description text')!
  const chart = slide.elements.find((e) => e.type === 'chart')!
  setElementAltText(slide, chart.id, { descr: 'Sales by quarter' })
  // An ink-flavored picture: name prefix + descr vector payload (renderer-owned marker)
  const ink = addPicture(opened, slide, {
    bytes: solidPng(8, 8, [0, 0, 255]),
    ext: 'png',
    offset: { ...OFF },
  })!
  ink.name = 'aislides-ink-3'
  ink.descr = '[[0,0],[4,4]]'
  return { opened, slide, boxId: box2.id, chartId: chart.id, inkId: ink.id }
}

const nodeBySource = (nodes: RenderNode[], id: string) => nodes.find((n) => n.sourceId === id)!

describe('render nodes carry alt text (PAR-304)', () => {
  it('exposes title/descr on shape and chart nodes', async () => {
    const { opened, slide, boxId, chartId } = await deckWithAlt()
    const rs = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 1280 })
    expect(nodeBySource(rs.nodes, boxId).altText).toEqual({
      title: 'Title text',
      descr: 'Description text',
    })
    const chartNode = nodeBySource(rs.nodes, chartId)
    expect(chartNode.altText).toEqual({ descr: 'Sales by quarter' })
    expect(chartNode.altTextLocked).toBeUndefined()
  })

  it('locks ink pictures whose descr slot holds the vector payload', async () => {
    const { opened, slide, inkId } = await deckWithAlt()
    const rs = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 1280 })
    const inkNode = nodeBySource(rs.nodes, inkId)
    expect(inkNode.altTextLocked).toBe(true)
    expect(inkNode.altText).toBeUndefined()
  })

  it('marks app-created charts (extLst marker) with appCreated', async () => {
    const { opened, slide, chartId } = await deckWithAlt()
    const rs = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 1280 })
    expect((nodeBySource(rs.nodes, chartId) as { appCreated?: boolean }).appCreated).toBe(true)
  })
})
