/**
 * PAR-304: object alt text (cNvPr title/description) authoring + the app-chart
 * marker decollision — the marker moved from cNvPr@descr (the user's alt-text
 * slot) to a cNvPr extLst ext, with the legacy descr marker still recognized.
 * All round-trips go insert → set → save → reopen.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  addChart,
  addElement,
  addPicture,
  addTable,
  appChartMarkerExtXml,
  editChartElement,
  markChartEditable,
  openPptx,
  savePptx,
  setElementAltText,
  solidPng,
  type ChartElement,
  type PictureElement,
  type TableElement,
  type TextElement,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const OFF = { x: 914400, y: 914400, cx: 4572000, cy: 2743200 }

describe('setElementAltText', () => {
  it('round-trips title+description on a shape and clears with null', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const shape = addElement(slide, {
      kind: 'textbox',
      offset: { ...OFF },
      paragraphs: [{ runs: [{ text: 'Alt' }] }],
    })
    expect(
      setElementAltText(slide, shape.id, { title: 'Board', descr: 'Quarterly board photo' }),
    ).toBe(true)
    expect(shape.title).toBe('Board')
    expect(shape.descr).toBe('Quarterly board photo')

    const reopened = await openPptx(await savePptx(opened))
    const el = reopened.deck.slides[0]!.elements.at(-1) as TextElement
    expect(el.title).toBe('Board')
    expect(el.descr).toBe('Quarterly board photo')
    expect(el.anchor.originalXml).toContain('title="Board"')
    expect(el.anchor.originalXml).toContain('descr="Quarterly board photo"')

    // null clears both attributes
    expect(setElementAltText(reopened.deck.slides[0]!, el.id, { title: null, descr: null })).toBe(
      true,
    )
    const reclosed = await openPptx(await savePptx(reopened))
    const el2 = reclosed.deck.slides[0]!.elements.at(-1) as TextElement
    expect(el2.title).toBeUndefined()
    expect(el2.descr).toBeUndefined()
    expect(el2.anchor.originalXml).not.toContain('descr="')
  })

  it('round-trips on pictures and tables too', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const pic = addPicture(opened, opened.deck.slides[0]!, {
      bytes: solidPng(8, 8, [255, 0, 0]),
      ext: 'png',
      offset: { ...OFF },
    })!
    // Insert ops reparse the slide: re-resolve the live element before editing
    const picLive = opened.deck.slides[0]!.elements.find((e) => e.type === 'picture')!
    expect(setElementAltText(opened.deck.slides[0]!, picLive.id, { descr: 'A red square' })).toBe(
      true,
    )
    expect(pic.id).toBe(picLive.id)
    addTable(opened, 0, { rows: 2, cols: 2, offset: { x: 0, y: 0, cx: 1828800, cy: 914400 } })
    const table = opened.deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement
    expect(
      setElementAltText(opened.deck.slides[0]!, table.id, {
        title: 'Budget',
        descr: 'Quarterly budget',
      }),
    ).toBe(true)

    const reopened = await openPptx(await savePptx(opened))
    const pic2 = reopened.deck.slides[0]!.elements.find(
      (e) => e.type === 'picture',
    ) as PictureElement
    expect(pic2.descr).toBe('A red square')
    const table2 = reopened.deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement
    expect(table2.title).toBe('Budget')
    expect(table2.descr).toBe('Quarterly budget')
  })

  it('refuses ink strokes and 3D posters (descr holds an editor payload)', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const pic = addPicture(opened, slide, {
      bytes: solidPng(8, 8, [0, 255, 0]),
      ext: 'png',
      offset: { ...OFF },
    })!
    // Simulate the renderer's ink (name prefix) and a 3D poster (descr payload)
    pic.name = 'aislides-ink-7'
    expect(setElementAltText(slide, pic.id, { descr: 'nope' })).toBe(false)
    pic.name = 'Picture 3'
    pic.descr = 'aislides-3d:ppt/media/model1.glb'
    expect(setElementAltText(slide, pic.id, { descr: 'nope' })).toBe(false)
  })
})

describe('app-chart marker (PAR-304 decollision)', () => {
  it('new charts carry the extLst marker, not the descr slot', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      categories: ['a'],
      series: [{ name: 's', values: [1] }],
      offset: { ...OFF },
    })!
    const el = opened.deck.slides[0]!.elements.find((e) => e.id === r.elementId) as ChartElement
    expect(el.appCreated).toBe(true)
    expect(el.descr).toBeUndefined()
    expect(el.anchor.originalXml).toContain(appChartMarkerExtXml())
    expect(el.anchor.originalXml).not.toContain('descr="aislides-chart"')

    const reopened = await openPptx(await savePptx(opened))
    const el2 = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el2.appCreated).toBe(true)
    expect(el2.descr).toBeUndefined()
  })

  it('a chart with user alt text stays app-editable through save/reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      categories: ['a'],
      series: [{ name: 's', values: [1] }],
      offset: { ...OFF },
    })!
    const slide = opened.deck.slides[0]!
    expect(
      setElementAltText(slide, r.elementId, { title: 'Sales', descr: 'Sales by quarter' }),
    ).toBe(true)
    expect(editChartElement(opened, 0, r.elementId, { gridlines: true })).toBe(true)

    const reopened = await openPptx(await savePptx(opened))
    const el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.appCreated).toBe(true)
    expect(el.title).toBe('Sales')
    expect(el.descr).toBe('Sales by quarter')
    expect(el.chart.valAxis?.gridColor).toBeTruthy()
    expect(editChartElement(reopened, 0, el.id, { gridlines: false })).toBe(true)
  })

  it('legacy descr marker files still parse as app-created and migrate on alt-text edit', async () => {
    // Build the legacy bytes: a chart frame carrying descr="aislides-chart" and no ext marker
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      categories: ['a'],
      series: [{ name: 's', values: [1] }],
      offset: { ...OFF },
    })!
    const el = opened.deck.slides[0]!.elements.find((e) => e.id === r.elementId) as ChartElement
    el.anchor.originalXml = el.anchor.originalXml
      .replace(appChartMarkerExtXml(), '')
      .replace(/<p:cNvPr\b([^>]*)>/, '<p:cNvPr$1 descr="aislides-chart">')
    opened.deck.slides[0]!.structureDirty = true

    const reopened = await openPptx(await savePptx(opened))
    const legacy = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    // Recognized as app-created; the marker is consumed, not surfaced as alt text
    expect(legacy.appCreated).toBe(true)
    expect(legacy.descr).toBeUndefined()
    expect(editChartElement(reopened, 0, legacy.id, { gridlines: true })).toBe(true)

    // Setting a real description must not lose editability: the ext marker is minted
    expect(setElementAltText(reopened.deck.slides[0]!, legacy.id, { descr: 'Revenue chart' })).toBe(
      true,
    )
    const final = await openPptx(await savePptx(reopened))
    const el3 = final.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el3.appCreated).toBe(true)
    expect(el3.descr).toBe('Revenue chart')
    expect(editChartElement(final, 0, el3.id, { gridlines: false })).toBe(true)
  })

  it('markChartEditable re-tags a foreign chart via the extLst marker', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      categories: ['a'],
      series: [{ name: 's', values: [1] }],
      offset: { ...OFF },
    })!
    const slide = opened.deck.slides[0]!
    // Simulate a foreign chart: strip the marker ext, keep a user description
    const el = slide.elements.find((e) => e.id === r.elementId) as ChartElement
    el.anchor.originalXml = el.anchor.originalXml.replace(appChartMarkerExtXml(), '')
    delete el.appCreated
    setElementAltText(slide, el.id, { descr: 'Foreign chart' })
    const stripped = await openPptx(await savePptx(opened))
    const foreign = stripped.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(foreign.appCreated).toBeUndefined()
    expect(editChartElement(stripped, 0, foreign.id, { gridlines: true })).toBe(false)

    expect(markChartEditable(stripped.deck.slides[0]!, foreign.id)).toBe(true)
    expect(foreign.descr).toBe('Foreign chart') // the user's alt text survives marking
    expect(editChartElement(stripped, 0, foreign.id, { gridlines: true })).toBe(true)

    const reopened = await openPptx(await savePptx(stripped))
    const el2 = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el2.appCreated).toBe(true)
    expect(el2.descr).toBe('Foreign chart')
  })
})
