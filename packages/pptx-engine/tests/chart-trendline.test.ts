import { execFileSync } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, it, expect } from 'vitest'
import { parseChartXml } from '../src/chart'
import { buildChartSpaceXml } from '../src/chart-insert'
import { addChart, openPptx, savePptx } from '../src/index'

const LINE_TRENDLINE = `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:lineChart><c:ser>
  <c:idx val="0"/><c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
  <c:cat><c:strRef><c:f>S!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>
    <c:pt idx="0"><c:v>2016</c:v></c:pt><c:pt idx="1"><c:v>2017</c:v></c:pt><c:pt idx="2"><c:v>2018</c:v></c:pt></c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>S!$B$2:$B$4</c:f><c:numCache><c:ptCount val="3"/>
    <c:pt idx="0"><c:v>100</c:v></c:pt><c:pt idx="1"><c:v>120</c:v></c:pt><c:pt idx="2"><c:v>145</c:v></c:pt></c:numCache></c:numRef></c:val>
  <c:trendline><c:name>Linear (Revenue)</c:name>
    <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
    <c:trendlineType val="linear"/><c:dispRSqr val="1"/><c:dispEq val="1"/>
  </c:trendline>
</c:ser></c:lineChart>
</c:plotArea></c:chart></c:chartSpace>`

const baseOpts = {
  kind: 'line' as const,
  categories: ['2016', '2017', '2018'],
  series: [
    { name: 'Revenue', values: [100, 120, 145] },
    { name: 'Cost', values: [80, 95, 110] },
  ],
  offset: { x: 0, y: 0, cx: 0, cy: 0 },
}

describe('c:trendline parsing', () => {
  it('parses a linear trendline: name, spPr stroke, dash, dispRSqr/dispEq', () => {
    const m = parseChartXml(LINE_TRENDLINE)!
    expect(m.series[0]!.trendlines).toHaveLength(1)
    expect(m.series[0]!.trendlines![0]).toEqual({
      kind: 'linear',
      name: 'Linear (Revenue)',
      color: '#FF0000',
      widthPt: 1.5,
      dash: 'dash',
      dispRSqr: true,
      dispEq: true,
    })
  })

  it('drops non-linear regressions (exp/log/poly/power/movingAvg are not modeled)', () => {
    const xml = LINE_TRENDLINE.replace(
      '<c:trendlineType val="linear"/>',
      '<c:trendlineType val="exp"/>',
    )
    const m = parseChartXml(xml)!
    expect(m.series[0]!.trendlines).toBeUndefined()
  })

  it('parses a chart with no trendline into no trendlines field', () => {
    const m = parseChartXml(LINE_TRENDLINE.replace(/<c:trendline>[\s\S]*?<\/c:trendline>/, ''))!
    expect(m.series[0]!.trendlines).toBeUndefined()
  })

  it('parses trendlines on scatter series (schema position before c:xVal)', () => {
    const xml = `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:scatterChart><c:ser>
  <c:idx val="0"/>
  <c:trendline><c:trendlineType val="linear"/><c:dispEq val="1"/></c:trendline>
  <c:xVal><c:numRef><c:f>x</c:f><c:numCache><c:ptCount val="3"/>
    <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:xVal>
  <c:yVal><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="3"/>
    <c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:yVal>
</c:ser></c:scatterChart>
</c:plotArea></c:chart></c:chartSpace>`
    const m = parseChartXml(xml)!
    expect(m.kind).toBe('scatter')
    expect(m.series[0]!.trendlines).toEqual([{ kind: 'linear', dispEq: true }])
  })
})

describe('trendline serialization', () => {
  it('writes c:trendline into the requested series only (schema position before c:cat)', () => {
    const xml = buildChartSpaceXml({
      ...baseOpts,
      trendline: { seriesIdx: 1, dispRSqr: true, dispEq: true },
    })
    expect(xml).toContain(
      '<c:trendline><c:trendlineType val="linear"/><c:dispRSqr val="1"/><c:dispEq val="1"/></c:trendline>',
    )
    // Exactly one trendline, inside the second series, before its c:cat
    expect(xml.match(/<c:trendline>/g)).toHaveLength(1)
    const ser2 = xml.slice(xml.indexOf('<c:ser><c:idx val="1"/>'))
    expect(ser2.indexOf('<c:trendline>')).toBeLessThan(ser2.indexOf('<c:cat>'))
  })

  it('writes no c:trendline without the option and omits absent flags', () => {
    const plain = buildChartSpaceXml(baseOpts)
    expect(plain).not.toContain('<c:trendline')
    const bare = buildChartSpaceXml({ ...baseOpts, trendline: {} })
    expect(bare).toContain('<c:trendline><c:trendlineType val="linear"/></c:trendline>')
  })

  it('writes the trendline before c:xVal on scatter charts', () => {
    const xml = buildChartSpaceXml({
      kind: 'scatter',
      categories: ['1', '2', '3'],
      series: [{ name: 'Y', values: [2, 4, 7] }],
      offset: { x: 0, y: 0, cx: 0, cy: 0 },
      trendline: { dispEq: true },
    })
    const ser = xml.slice(xml.indexOf('<c:ser>'))
    expect(ser.indexOf('<c:trendline>')).toBeLessThan(ser.indexOf('<c:xVal>'))
  })

  it('round-trips: a serialized trendline parses back with the same flags', () => {
    const xml = buildChartSpaceXml({
      ...baseOpts,
      trendline: { seriesIdx: 0, dispRSqr: true, dispEq: true },
    })
    const m = parseChartXml(xml)!
    expect(m.series[0]!.trendlines).toEqual([{ kind: 'linear', dispRSqr: true, dispEq: true }])
    expect(m.series[1]!.trendlines).toBeUndefined()
  })
})

// ── python-pptx cross-check (same opportunistic probe as chart-workbook.test.ts) ──

const here = dirname(fileURLToPath(import.meta.url))

function hasPythonPptx(): boolean {
  try {
    execFileSync('python3', ['-c', 'import pptx'], { stdio: 'ignore' })
    return true
  } catch {
    return false
  }
}
const withPptx = describe.skipIf(!hasPythonPptx())

withPptx('trendline round-trip through a saved deck (structural check)', () => {
  let dir = ''
  afterEach(() => {
    if (dir) {
      rmSync(dir, { recursive: true, force: true })
      dir = ''
    }
  })

  it('insert → save → python-pptx opens the chart and finds the trendline in schema order', async () => {
    const opened = await openPptx(readFileSync(join(here, 'fixtures', '01_standard_business.pptx')))
    const r = addChart(opened, 0, {
      kind: 'line',
      categories: ['Q1', 'Q2', 'Q3'],
      series: [{ name: 'Sales', values: [10, 20, 30] }],
      offset: { x: 914400, y: 914400, cx: 4572000, cy: 2743200 },
      trendline: { seriesIdx: 0, dispRSqr: true, dispEq: true },
    })!
    expect(r).not.toBeNull()
    const bytes = await savePptx(opened)
    dir = mkdtempSync(join(tmpdir(), 'par211-'))
    const file = join(dir, 'deck.pptx')
    writeFileSync(file, bytes)

    const script = [
      'import json',
      'from lxml import etree',
      'from pptx import Presentation',
      `p = Presentation(${JSON.stringify(file)})`,
      'out = []',
      'for slide in p.slides:',
      '    for shape in slide.shapes:',
      '        if shape.has_chart:',
      '            root = etree.fromstring(shape.chart._chartSpace.xml.encode())',
      '            local = lambda tag: [e for e in root.iter() if etree.QName(e).localname == tag]',
      '            ser_order = [[etree.QName(ch).localname for ch in tl if etree.QName(ch).localname != "trendline"] for tl in local("trendline")]',
      '            out.append({',
      '                "series_values": [[float(v) for v in s.values] for s in shape.chart.plots[0].series],',
      '                "trendline_count": len(local("trendline")),',
      '                "trendline_type": [e.get("val") for e in local("trendlineType")],',
      '                "flags": sorted(etree.QName(e).localname for e in root.iter() if etree.QName(e).localname in ("dispRSqr", "dispEq")),',
      '                "ser_child_order": ser_order,',
      '            })',
      'print(json.dumps(out))',
    ].join('\n')
    const stdout = execFileSync('python3', ['-c', script], { encoding: 'utf8' })
    const charts = JSON.parse(stdout.trim()) as Array<{
      series_values: number[][]
      trendline_count: number
      trendline_type: string[]
      flags: string[]
      ser_child_order: string[][]
    }>
    expect(charts).toHaveLength(1)
    // python-pptx parsed the deck fine and sees the series data
    expect(charts[0].series_values).toEqual([[10, 20, 30]])
    expect(charts[0].trendline_count).toBe(1)
    expect(charts[0].trendline_type).toEqual(['linear'])
    expect(charts[0].flags).toEqual(['dispEq', 'dispRSqr'])
    // CT_TrendLine schema order inside c:trendline: trendlineType then the boolean flags
    expect(charts[0].ser_child_order[0]).toEqual(['trendlineType', 'dispRSqr', 'dispEq'])
  })
})
