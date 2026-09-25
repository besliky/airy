/**
 * PAR-302: chart embedded workbook (ppt/embeddings/*.xlsx) — "Edit Data" parity.
 *
 * App-inserted charts now carry an embedded xlsx (chart part rels …/package rel
 * + <c:externalData> + Content_Types xlsx Default) and every chart edit rewrites
 * the Sheet1 data rectangle so the sheet PowerPoint's own "Edit Data" shows stays
 * in sync with the rendered chart. Regression-covered here: container build/read
 * round-trip, insert wiring, edit sync (chart XML caches + sheet cells), authored
 * content preservation (formulas, cells outside the rectangle, foreign parts),
 * save→open round-trip, and python-pptx reading the updated values.
 */
import { execFileSync } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'
import {
  addChart,
  editChartElement,
  findChartEmbeddingPath,
  findChartPartPath,
  getChartElementWorkbookData,
  openPptx,
  parseChartXml,
  patchChartWorkbookBytes,
  readChartWorkbookBytes,
  readZipContainer,
  savePptx,
  writeZipContainer,
  type ChartElement,
  type PackageArchive,
} from '../src/index'
import { resolveTarget } from '../src/zip'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const OFF = { x: 914400, y: 914400, cx: 4572000, cy: 2743200 }

/** chart part path of the first chart element on slide 0 */
function chartPartOf(archive: PackageArchive): string {
  const slidePath = 'ppt/slides/slide1.xml'
  for (const rel of archive.readRels(slidePath).values()) {
    if (rel.type.endsWith('/chart')) return resolveTarget(slidePath, rel.target)
  }
  throw new Error('no chart relationship on slide 1')
}

/** the single embedding of slide 0's chart (addChart mints exactly one) */
function embeddingOf(archive: PackageArchive): string {
  const emb = findChartEmbeddingPath(archive, chartPartOf(archive))
  if (!emb) throw new Error('chart part must reference an embedded workbook')
  expect(emb).toMatch(/^ppt\/embeddings\/.*\.xlsx$/)
  expect(archive.readBytes(emb)).toBeTruthy()
  return emb
}

describe('embedded workbook on insert (PAR-302)', () => {
  it('addChart wires chart part → embeddings xlsx (rels + externalData + content type)', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      categories: ['Q1', 'Q2', 'Q3'],
      series: [
        { name: 'Sales', values: [10, 20, 30] },
        { name: 'Costs', values: [5, 8, 12] },
      ],
      offset: { ...OFF },
    })!
    const archive = opened.archive
    const chartPath = findChartPartPath(archive, opened.deck.slides[0]!, r.elementId)
    expect(chartPath).toMatch(/^ppt\/charts\/chart\d+\.xml$/)
    const part = chartPath as string
    const chartXml = archive.readText(part)!
    // <c:externalData> points at the package rel and follows the c:chart block
    expect(chartXml).toContain('<c:externalData r:id="')
    expect(chartXml.indexOf('<c:externalData')).toBeGreaterThan(chartXml.indexOf('</c:chart>'))
    // chart part rels carry the …/package relationship into ppt/embeddings/
    const rels = [...archive.readRels(part).values()]
    expect(
      rels.some((rel) => rel.type.endsWith('/package') && rel.target.includes('embeddings')),
    ).toBe(true)
    // [Content_Types].xml covers the .xlsx extension
    expect(archive.readText('[Content_Types].xml')).toContain('Extension="xlsx"')

    // the sheet mirrors the inserted data
    const table = readChartWorkbookBytes(archive.readBytes(embeddingOf(archive))!)!
    expect(table.categories).toEqual(['Q1', 'Q2', 'Q3'])
    expect(table.series).toEqual([
      { name: 'Sales', values: [10, 20, 30] },
      { name: 'Costs', values: [5, 8, 12] },
    ])
  })

  it('readChartWorkbookBytes round-trips shared strings, numbers and empty cells', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    addChart(opened, 0, {
      kind: 'line',
      categories: ['a', '', 'c'],
      series: [
        { name: 's1', values: [1.5, 2, 3] },
        { name: '', values: [9, 8, 7] },
      ],
      offset: { ...OFF },
    })
    const table = readChartWorkbookBytes(opened.archive.readBytes(embeddingOf(opened.archive))!)!
    expect(table.categories).toEqual(['a', '', 'c'])
    expect(table.series).toEqual([
      { name: 's1', values: [1.5, 2, 3] },
      { name: '', values: [9, 8, 7] },
    ])
  })
})

describe('embedded workbook sync on edit (PAR-302)', () => {
  it('editChartElement updates the chart XML caches and the Sheet1 rectangle together', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      categories: ['Q1', 'Q2'],
      series: [{ name: 'Sales', values: [10, 20] }],
      offset: { ...OFF },
    })!
    const archive = opened.archive
    const chartPath = findChartPartPath(archive, opened.deck.slides[0]!, r.elementId)!

    expect(
      editChartElement(opened, 0, r.elementId, {
        categories: ['Jan', 'Feb', 'Mar'],
        series: [{ name: 'Revenue', values: [111, 122, 133] }],
      }),
    ).toBe(true)

    // chart part: caches regenerated from the patch
    const model = parseChartXml(archive.readText(chartPath)!)!
    expect(model.categories).toEqual(['Jan', 'Feb', 'Mar'])
    expect(model.series[0]!.values).toEqual([111, 122, 133])

    // workbook: the same numbers
    const table = readChartWorkbookBytes(archive.readBytes(embeddingOf(archive))!)!
    expect(table.categories).toEqual(['Jan', 'Feb', 'Mar'])
    expect(table.series).toEqual([{ name: 'Revenue', values: [111, 122, 133] }])
  })

  it('getChartElementWorkbookData reads the sheet the Edit Data dialog opens', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'pie',
      categories: ['North', 'South'],
      series: [{ name: 'Share', values: [60, 40] }],
      offset: { ...OFF },
    })!
    expect(getChartElementWorkbookData(opened, 0, r.elementId)).toEqual({
      categories: ['North', 'South'],
      series: [{ name: 'Share', values: [60, 40] }],
    })
    // non-chart ids have nothing to read
    expect(getChartElementWorkbookData(opened, 0, 'nope')).toBeNull()
  })

  it('a second chart mints a second, non-conflicting embedding name', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    addChart(opened, 0, {
      kind: 'bar',
      categories: ['a'],
      series: [{ name: 's', values: [1] }],
      offset: { ...OFF },
    })
    addChart(opened, 0, {
      kind: 'bar',
      categories: ['b'],
      series: [{ name: 't', values: [2] }],
      offset: { x: 9144000, y: 914400, cx: 4572000, cy: 2743200 },
    })
    const names = [...opened.archive.entries.keys()].filter((p) =>
      /^ppt\/embeddings\/.*\.xlsx$/.test(p),
    )
    expect(names).toHaveLength(2)
    expect(new Set(names).size).toBe(2)
    for (const name of names) {
      expect(opened.archive.readBytes(name)!.length).toBeGreaterThan(0)
    }
  })
})

describe('patch preserves authored workbook content', () => {
  it('keeps formula cells, out-of-rectangle cells and foreign parts byte-identical', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    addChart(opened, 0, {
      kind: 'bar',
      categories: ['Q1', 'Q2'],
      series: [{ name: 'Sales', values: [10, 20] }],
      offset: { ...OFF },
    })
    const archive = opened.archive
    const bytes = archive.readBytes(embeddingOf(archive))!

    // author a workbook by hand: a formula inside the rectangle (Excel keeps its
    // own cached <v>), a helper cell + styled row outside it, a foreign part
    const parts = readZipContainer(bytes)!
    const sheetXml = Buffer.from(parts.get('xl/worksheets/sheet1.xml')!).toString('utf8')
    const authoredSheet = sheetXml
      .replace('<c r="B3"><v>20</v></c>', '<c r="B3"><f>B2*2</f><v>40</v></c>')
      .replace(
        '</sheetData>',
        '<row r="9" ht="14" customHeight="1"/>' +
          '<row r="10"><c r="E10" t="inlineStr"><is><t>note</t></is></c></row></sheetData>',
      )
    parts.set('xl/worksheets/sheet1.xml', Buffer.from(authoredSheet, 'utf8'))
    parts.set('xl/tables/table1.xml', Buffer.from('<table xmlns="x"/>', 'utf8'))
    const authored = writeZipContainer([...parts].map(([name, data]) => ({ name, data })))

    const updated = patchChartWorkbookBytes(authored, {
      categories: ['Q1', 'Q2', 'Q3'],
      series: [{ name: 'Sales', values: [11, 22, 33] }],
    })
    expect(updated).toBeTruthy()
    const after = readZipContainer(updated!)!

    // foreign part survives untouched
    expect(Buffer.from(after.get('xl/tables/table1.xml')!).toString('utf8')).toBe(
      '<table xmlns="x"/>',
    )
    const afterSheet = Buffer.from(after.get('xl/worksheets/sheet1.xml')!).toString('utf8')
    // formula cell byte-identical (formula + its own recalc cache)
    expect(afterSheet).toContain('<c r="B3"><f>B2*2</f><v>40</v></c>')
    // out-of-rectangle helper cell + styled row survive
    expect(afterSheet).toContain('<c r="E10" t="inlineStr"><is><t>note</t></is></c>')
    expect(afterSheet).toContain('<row r="9" ht="14" customHeight="1"/>')
    // dimension covers the union of old extent, new rectangle and survivors
    expect(afterSheet).toContain('<dimension ref="A1:E10"/>')
    // the readable rectangle got the new numbers; the formula column keeps its
    // own cache (40), exactly what Excel's first recalc would reproduce
    const table = readChartWorkbookBytes(updated!)!
    expect(table.categories).toEqual(['Q1', 'Q2', 'Q3'])
    expect(table.series).toEqual([{ name: 'Sales', values: [11, 40, 33] }])
    // the helper column never leaks into the chart as a series
    expect(table.series).toHaveLength(1)
  })

  it('returns null for a corrupt container instead of throwing', () => {
    expect(
      patchChartWorkbookBytes(new Uint8Array([1, 2, 3]), { categories: [], series: [] }),
    ).toBeNull()
    expect(readChartWorkbookBytes(new Uint8Array([1, 2, 3]))).toBeNull()
  })
})

describe('save → open round-trip', () => {
  it('embedding, rels and edited values survive a full save/reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      categories: ['Q1', 'Q2'],
      series: [{ name: 'Sales', values: [10, 20] }],
      offset: { ...OFF },
    })!
    editChartElement(opened, 0, r.elementId, {
      categories: ['Jan', 'Feb'],
      series: [{ name: 'Revenue', values: [50, 75] }],
    })
    const reopened = await openPptx(await savePptx(opened))

    // element ids refresh on the reopen reparse: find the chart by kind
    const el = reopened.deck.slides[0]!.elements.find((e) => e.type === 'chart') as ChartElement
    expect(el.type).toBe('chart')
    expect(el.chart.categories).toEqual(['Jan', 'Feb'])
    expect(el.chart.series[0]!.values).toEqual([50, 75])

    // the workbook travels with the deck and still matches
    expect(getChartElementWorkbookData(reopened, 0, el.id)).toEqual({
      categories: ['Jan', 'Feb'],
      series: [{ name: 'Revenue', values: [50, 75] }],
    })
  })
})

// ── python-pptx cross-check (opportunistic, same probe as save-lo-interop) ──

function hasPythonPptx(): boolean {
  try {
    execFileSync('python3', ['-c', 'import pptx'], { stdio: 'ignore' })
    return true
  } catch {
    return false
  }
}
const withPptx = describe.skipIf(!hasPythonPptx())

withPptx('python-pptx reads the edited values (PAR-302 DoD)', () => {
  let dir = ''
  afterEach(() => {
    if (dir) {
      rmSync(dir, { recursive: true, force: true })
      dir = ''
    }
  })

  it('insert → edit → save → python-pptx sees the new categories/series', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      categories: ['Q1', 'Q2'],
      series: [{ name: 'Sales', values: [10, 20] }],
      offset: { ...OFF },
    })!
    editChartElement(opened, 0, r.elementId, {
      categories: ['Jan', 'Feb', 'Mar'],
      series: [{ name: 'Revenue', values: [42, 43, 44] }],
    })
    const bytes = await savePptx(opened)
    dir = mkdtempSync(join(tmpdir(), 'par302-'))
    const file = join(dir, 'deck.pptx')
    writeFileSync(file, bytes)

    const script = [
      'import json',
      'from pptx import Presentation',
      `p = Presentation(${JSON.stringify(file)})`,
      'out = []',
      'for slide in p.slides:',
      '    for shape in slide.shapes:',
      '        if shape.has_chart:',
      '            c = shape.chart',
      '            out.append({',
      '                "categories": [str(x) for x in c.plots[0].categories],',
      '                "series": [{"name": s.name, "values": list(s.values)} for s in c.series],',
      '            })',
      'print(json.dumps(out))',
    ].join('\n')
    const stdout = execFileSync('python3', ['-c', script], { encoding: 'utf8' })
    const charts = JSON.parse(stdout.trim()) as Array<{
      categories: string[]
      series: Array<{ name: string; values: number[] }>
    }>
    expect(charts).toHaveLength(1)
    expect(charts[0].categories).toEqual(['Jan', 'Feb', 'Mar'])
    expect(charts[0].series[0].name).toBe('Revenue')
    expect(charts[0].series[0].values).toEqual([42, 43, 44])
  })
})
