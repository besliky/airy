import { describe, it, expect } from 'vitest'
import { parseChartXml } from '../src/chart'

const ofPie = (
  body: string,
) => `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:ofPieChart><c:varyColors val="1"/>${body}
</c:ofPieChart></c:plotArea></c:chart></c:chartSpace>`

const SER = `<c:ser><c:idx val="0"/>
  <c:cat><c:strRef><c:f>S!$A$2:$A$6</c:f><c:strCache><c:ptCount val="5"/>
    <c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt>
    <c:pt idx="3"><c:v>D</c:v></c:pt><c:pt idx="4"><c:v>E</c:v></c:pt></c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>S!$B$2:$B$6</c:f><c:numCache><c:ptCount val="5"/>
    <c:pt idx="0"><c:v>70</c:v></c:pt><c:pt idx="1"><c:v>15</c:v></c:pt><c:pt idx="2"><c:v>5</c:v></c:pt>
    <c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser>`

describe('c:ofPieChart parsing (pie-of-pie)', () => {
  it('parses kind pieOfPie with explicit split settings', () => {
    const m = parseChartXml(
      ofPie(
        `<c:ofPieType val="bar"/>${SER}<c:gapWidth val="150"/><c:splitType val="auto"/><c:splitPos val="20"/><c:secondPieSize val="80"/>`,
      ),
    )!
    expect(m.kind).toBe('pieOfPie')
    expect(m.series[0]!.values).toEqual([70, 15, 5, 3, 2])
    expect(m.pieOfPie).toEqual({
      type: 'bar',
      splitType: 'auto',
      splitPos: 20,
      secondPieSizePct: 80,
      gapWidthPct: 150,
    })
  })

  it('falls back to PowerPoint defaults when the settings are absent', () => {
    const m = parseChartXml(ofPie(`<c:ofPieType val="pie"/>${SER}`))!
    expect(m.kind).toBe('pieOfPie')
    expect(m.pieOfPie).toEqual({
      type: 'pie',
      splitPos: 10,
      secondPieSizePct: 75,
      gapWidthPct: 100,
    })
  })

  it('a plain pie chart keeps kind pie (no pieOfPie field)', () => {
    const xml = ofPie(`<c:ofPieType val="pie"/>${SER}`).replace('c:ofPieChart', 'c:pieChart')
    const m = parseChartXml(xml)!
    expect(m.kind).toBe('pie')
    expect(m.pieOfPie).toBeUndefined()
  })
})
