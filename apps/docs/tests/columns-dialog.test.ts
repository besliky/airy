import { describe, expect, it } from 'vitest'
import {
  DEFAULT_SECTION,
  applySectionSettings,
  sectionSettingsFromXml,
  type SectionSettings,
} from '@airy-office/docx-engine'
import { columnPresetSpec, presetOf } from '../src/renderer/components/ColumnsDialog'
import { sectionColGeom } from '../src/renderer/pagination-sections'
import type { SectionInfo } from '@airy-office/docx-engine'

const A4_CONTENT = 11906 - 1440 - 1440 // 9026 twips

describe('Columns dialog presets', () => {
  it('equal presets carry no explicit widths; Left/Right split 1.69:1', () => {
    expect(columnPresetSpec('one', A4_CONTENT, 720)).toEqual({ columns: 1 })
    expect(columnPresetSpec('two', A4_CONTENT, 720)).toEqual({ columns: 2 })
    expect(columnPresetSpec('three', A4_CONTENT, 720)).toEqual({ columns: 3 })
    const left = columnPresetSpec('left', 9026, 720)
    const right = columnPresetSpec('right', 9026, 720)
    expect(left.colWidths![0]).toBeGreaterThan(left.colWidths![1]!)
    expect(right.colWidths![0]!).toBeLessThan(right.colWidths![1]!)
    // widths + spacing fill the content area
    for (const spec of [left, right]) {
      expect(spec.colWidths![0]! + spec.colWidths![1]! + 720).toBeCloseTo(9026, -1)
    }
  })

  it('presetOf reads a section back into its preset', () => {
    const base = { ...DEFAULT_SECTION }
    expect(presetOf(base)).toBe('one')
    expect(presetOf({ ...base, columns: 2 })).toBe('two')
    expect(presetOf({ ...base, columns: 3 })).toBe('three')
    expect(presetOf({ ...base, columns: 2, colWidths: [6000, 3000] })).toBe('left')
    expect(presetOf({ ...base, columns: 2, colWidths: [3000, 6000] })).toBe('right')
    // near-equal explicit widths read as the equal preset
    expect(presetOf({ ...base, columns: 2, colWidths: [4513, 4513] })).toBe('two')
  })

  it('a preset applied through the engine round-trips (w:cols equalWidth=0 + w:col)', () => {
    const sectPr =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>'
    const spec = columnPresetSpec('left', A4_CONTENT, 720)
    const settings: SectionSettings = {
      ...sectionSettingsFromXml(sectPr),
      columns: spec.columns,
      colSpace: 720,
      colWidths: spec.colWidths,
      columnSep: true,
    }
    const out = applySectionSettings(sectPr, settings)
    expect(out).toContain('w:equalWidth="0"')
    expect(out).toContain('<w:sep/>')
    const parsed = sectionSettingsFromXml(out)
    expect(parsed.colWidths).toEqual(spec.colWidths)
    expect(parsed.columnSep).toBe(true)
    expect(presetOf(parsed)).toBe('left')
  })
})

describe('renderer column geometry (unequal widths + separator)', () => {
  it('sectionColGeom respects explicit per-column widths and gaps', () => {
    const sectPr =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
      '<w:cols w:num="2" w:space="720" w:equalWidth="0"><w:sep/><w:col w:w="6000" w:space="720"/><w:col w:w="3026"/></w:cols></w:sectPr>'
    const sec = { settings: sectionSettingsFromXml(sectPr), sectPrXml: sectPr } as SectionInfo
    const geom = sectionColGeom(sec)
    expect(geom.cols).toBe(2)
    expect(geom.equalWidth).toBe(false)
    expect(Math.round(geom.widths[0]!)).toBe(400) // 6000 twips = 400 px @96dpi
    expect(Math.round(geom.widths[1]!)).toBeCloseTo(202, 0)
    expect(sec.settings.columnSep).toBe(true)
  })
})
