import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import {
  parseDocx,
  patchChartPartXml,
  saveDocx,
  type ChartDisplay,
  type NewChart,
} from '@airy-office/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { applyChartEdits, chartEditSource, displayFromSpec } from '../src/renderer/editor/chart'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const BAR_SPEC: NewChart = {
  kind: 'bar',
  title: 'Sales',
  categories: ['January', 'February'],
  series: [{ name: 'East', values: [10, 20] }],
}

/** doc with one native chart (word/charts/chart1.xml + embedded workbook) */
async function chartDoc(): Promise<Uint8Array> {
  const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>' })
  const parsed = await parseDocx(source)
  return saveDocx(parsed, [
    { kind: 'original', docxIndex: 0 },
    { kind: 'chart', chart: BAR_SPEC },
  ])
}

async function workbookSheet(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  const wb = await zip.file('word/charts/embeddings/workbook1.xlsx')!.async('uint8array')
  const xlsx = await JSZip.loadAsync(wb)
  return xlsx.file('xl/worksheets/sheet1.xml')!.async('string')
}

describe('chart insert/edit model conversions', () => {
  it('doughnut displays as pie + hole and reads back as doughnut', () => {
    const spec: NewChart = {
      kind: 'doughnut',
      title: 'Share',
      categories: ['A', 'B', 'C'],
      series: [{ name: 'S1', values: [1, 2, 3] }],
    }
    const display = displayFromSpec(spec)
    expect(display.kind).toBe('pie')
    expect(display.holePct).toBe(50)
    // reopening the modal on that display recovers the doughnut kind
    expect(chartEditSource(display, null)?.kind).toBe('doughnut')
  })

  it('scatter display carries derived x values; bubble sizes default uniformly', () => {
    const scatter = displayFromSpec({
      kind: 'scatter',
      title: 'XY',
      categories: ['1', '2', '3'],
      series: [{ name: 'S1', values: [4, 5, 6] }],
    })
    expect(scatter.kind).toBe('scatter')
    expect(scatter.series[0].xValues).toEqual([1, 2, 3])
    const bubble = displayFromSpec({
      kind: 'bubble',
      title: 'B',
      categories: ['1', '2'],
      series: [{ name: 'S1', values: [4, 5] }],
    })
    expect(bubble.series[0].sizes).toEqual([100, 100])
    // non-numeric x cells degrade to gaps, not NaN
    expect(
      displayFromSpec({
        kind: 'scatter',
        title: 'X',
        categories: ['1', 'B', '3'],
        series: [{ name: 'S1', values: [4, 5, 6] }],
      }).series[0].xValues,
    ).toEqual([1, null, 3])
  })
})

describe('chart edit → embedded workbook sync', () => {
  it('a part-only patch desyncs cache and workbook; applyChartEdits resyncs them', async () => {
    const source = await chartDoc()
    const parsed = await parseDocx(source)
    const partPath = 'word/charts/chart1.xml'
    const patch = {
      title: 'Updated Sales',
      series: [{ name: 'East Region', values: [999, null] }],
    }

    // the pre-sync save path: patch the chart part only — the workbook keeps
    // the original numbers (the audit's "Edit Data shows stale numbers" gap)
    const desynced = await saveDocx(
      parsed,
      parsed.blocks
        .filter((b) => !b.hidden)
        .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
      {
        partXml: { [partPath]: patchChartPartXml(parsed.extras.chartParts[partPath], patch) },
      },
    )
    expect(await workbookSheet(desynced)).not.toContain('<v>999</v>')

    // the save pipeline's chart application: cache patch + workbook sync
    const { partXml, partBinary } = await applyChartEdits(source, parsed.extras.chartParts, [
      { partPath, patch },
    ])
    expect(Object.keys(partXml)).toEqual([partPath])
    expect(Object.keys(partBinary)).toEqual(['word/charts/embeddings/workbook1.xlsx'])
    const synced = await saveDocx(
      parsed,
      parsed.blocks
        .filter((b) => !b.hidden)
        .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
      { partXml, partBinary },
    )

    // cache and workbook now agree
    const reparsed = await parseDocx(synced)
    const chart = reparsed.blocks.find((b) => b.chartDisplay)
    expect(chart?.chartDisplay?.title).toBe('Updated Sales')
    expect(chart?.chartDisplay?.series[0]).toEqual({ name: 'East Region', values: [999, 20] })
    const sheet = await workbookSheet(synced)
    expect(sheet).toContain('<v>999</v>')
    expect(sheet).toContain('East Region')
    expect(sheet).not.toContain('<v>10</v>')
  })

  it('the editor save pipeline syncs in-place edits to cache and workbook', async () => {
    const source = await chartDoc()
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    // double-click edit the first data cell, like chart-edit.test.ts
    const value = editor.view.dom.querySelector('.doc-chart-val') as HTMLElement
    value.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))
    value.textContent = '77'
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.chartPatches).toHaveLength(1)
    const { partXml, partBinary } = await applyChartEdits(
      source,
      parsed.extras.chartParts,
      plan.chartPatches,
    )
    const saved = await saveDocx(parsed, plan.saveBlocks, { partXml, partBinary })
    const reparsed = await parseDocx(saved)
    const chart = reparsed.blocks.find((b) => b.chartDisplay)
    expect(chart?.chartDisplay?.series[0].values[0]).toBe(77)
    expect(await workbookSheet(saved)).toContain('<v>77</v>')
    editor.destroy()
  })
})

describe('new chart kinds from the editor', () => {
  const SPECS: NewChart[] = [
    {
      kind: 'area',
      title: 'Area',
      categories: ['Q1', 'Q2'],
      series: [{ name: 'S1', values: [3, 5] }],
    },
    {
      kind: 'scatter',
      title: 'Scatter',
      categories: ['1', '2'],
      series: [{ name: 'S1', values: [4, 6] }],
    },
    {
      kind: 'bubble',
      title: 'Bubble',
      categories: ['1', '2'],
      series: [{ name: 'S1', values: [4, 6], sizes: [20, 40] }],
    },
    {
      kind: 'doughnut',
      title: 'Doughnut',
      categories: ['A', 'B'],
      series: [{ name: 'S1', values: [7, 3] }],
    },
  ]

  it.each(SPECS.map((spec) => [spec.kind, spec] as const))(
    '%s insert saves a chart + workbook and survives reopen/resave',
    async (kind, spec) => {
      const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>' })
      const parsed = await parseDocx(source)
      const editor = new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content: blocksToPmDoc(parsed.blocks) as never,
      })
      editor
        .chain()
        .insertContentAt(editor.state.doc.content.size, {
          type: 'docProtected',
          attrs: {
            docxIndex: null,
            blockType: 'chart',
            label: 'Chart',
            genChart: spec,
            chartDisplay: displayFromSpec(spec),
          },
        })
        .run()
      // the in-session display renders the right shape immediately
      expect(editor.view.dom.querySelector('.doc-chart-svg, svg')).toBeTruthy()

      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const saved = await saveDocx(parsed, plan.saveBlocks)
      const reparsed = await parseDocx(saved)
      const display: ChartDisplay | null =
        reparsed.blocks.find((b) => b.chartDisplay)?.chartDisplay ?? null
      expect(display?.title).toBe(spec.title)
      if (kind === 'doughnut') {
        expect(display?.kind).toBe('pie')
        expect(display?.holePct).toBe(50)
      } else {
        expect(display?.kind).toBe(kind)
      }
      if (kind === 'scatter' || kind === 'bubble') {
        expect(display?.series[0].xValues).toEqual(spec.categories.map(Number))
      }
      if (kind === 'bubble') expect(display?.series[0].sizes).toEqual([20, 40])

      // reopen + resave must not damage the chart or its workbook
      const visible = reparsed.blocks.filter((b) => !b.hidden)
      const resaved = await saveDocx(
        reparsed,
        visible.map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
      )
      const [zipA, zipB] = await Promise.all([JSZip.loadAsync(saved), JSZip.loadAsync(resaved)])
      expect(await zipB.file('word/charts/chart1.xml')!.async('string')).toBe(
        await zipA.file('word/charts/chart1.xml')!.async('string'),
      )
      expect(await zipB.file('word/charts/embeddings/workbook1.xlsx')!.async('uint8array')).toEqual(
        await zipA.file('word/charts/embeddings/workbook1.xlsx')!.async('uint8array'),
      )
      editor.destroy()
    },
  )
})
