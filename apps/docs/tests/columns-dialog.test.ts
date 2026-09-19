import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  DEFAULT_SECTION,
  applySectionSettings,
  sectionSettingsFromXml,
  type SectionSettings,
} from '@airy-office/docx-engine'
import { columnPresetSpec, presetOf } from '../src/renderer/components/ColumnsDialog'
import { sectionColGeom } from '../src/renderer/pagination-sections'
import { syncColumnRules } from '../src/renderer/editor/pagination-gaps'
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

describe('column separator surfaces (PAR-113 w:sep, TEST-1004)', () => {
  // sectPr with w:sep: A4 portrait, 1440-twip margins, two equal columns
  // with a 720-twip gap (the values the geometry assertions below derive from)
  const SECT_WSEP =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>' +
    '<w:cols w:num="2" w:space="720"><w:sep/></w:cols></w:sectPr>'
  const sectionOf = (sectPr: string) =>
    ({ settings: sectionSettingsFromXml(sectPr), sectPrXml: sectPr }) as SectionInfo

  it('syncColumnRules inks one centered hairline per column gap on every page', () => {
    const rectOf = (top: number, height: number) => ({ top, height }) as DOMRect
    const wrap = document.createElement('div')
    wrap.getBoundingClientRect = () => rectOf(0, 1000)
    // one 40px gap decoration at 480: paper bounds become [0,480,520,1000]
    const gap = document.createElement('div')
    gap.className = 'page-gap'
    gap.getBoundingClientRect = () => rectOf(480, 40)
    wrap.appendChild(gap)
    document.body.appendChild(wrap)

    const slices = [
      { start: 0, end: 5, section: 0 },
      { start: 5, end: 10, section: 0 },
    ]
    syncColumnRules(wrap, slices, [sectionOf(SECT_WSEP)], 1)

    const rules = [...wrap.querySelectorAll<HTMLElement>('.page-colrule')]
    expect(rules.length).toBe(2) // one per page, one gap between the two columns
    const [first, second] = rules
    // rules live inside the overlay layer, not the page content flow
    expect(first!.closest('.page-colrule-overlays')).toBe(
      wrap.querySelector('.page-colrule-overlays'),
    )
    // vertical extent: paper bounds minus the 1440-twip (96px) margins
    expect(first!.style.top).toBe('96px')
    expect(first!.style.height).toBe('288px')
    expect(second!.style.top).toBe('616px')
    // horizontal: left margin (96px) + first column + half the 720-twip gap
    //   content = 9026 twips = 601.73px; column = (601.73 - 48) / 2 = 276.87px
    expect(parseFloat(first!.style.left)).toBeCloseTo(96 + 276.87 + 24, 1)
    expect(first!.style.left).toBe(second!.style.left)
    wrap.remove()
  })

  it('removes the overlay layer when no section carries a separator', () => {
    const wrap = document.createElement('div')
    wrap.getBoundingClientRect = () => ({ top: 0, height: 1000 }) as DOMRect
    const gap = document.createElement('div')
    gap.className = 'page-gap'
    gap.getBoundingClientRect = () => ({ top: 480, height: 40 }) as DOMRect
    wrap.appendChild(gap)
    document.body.appendChild(wrap)
    const layer = document.createElement('div')
    layer.className = 'page-colrule-overlays'
    wrap.appendChild(layer)

    const noSep = SECT_WSEP.replace('<w:sep/>', '')
    syncColumnRules(wrap, [{ start: 0, end: 5, section: 0 }], [sectionOf(noSep)], 1)
    expect(wrap.querySelector('.page-colrule-overlays')).toBeNull()
    wrap.remove()
  })

  it('the preview surface and the simple-flow CSS both carry the separator', () => {
    // Source contract (split-button-a11y pattern): mounting the paginated
    // preview needs the whole pagination fixture; what matters statically is
    // that both remaining surfaces gate their rule on the section's w:sep
    // flag. Preview: a hairline centered in the gap, only BETWEEN columns.
    const preview = readFileSync(
      join(__dirname, '../src/renderer/components/PaginationPreview.tsx'),
      'utf8',
    )
    expect(preview).toContain('rSec?.settings.columnSep && ci < region.columns.length - 1')
    expect(preview).toContain('pv-colrule')
    // Simple (non-regioned) flow: the doc-page style gets a CSS column-rule.
    const app = readFileSync(join(__dirname, '../src/renderer/App.tsx'), 'utf8')
    expect(app).toContain("columnSep ? ' column-rule: 1px solid var(--docs-paper-ink);'")
  })
})
