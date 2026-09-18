import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  applySectionSettings,
  parseDocx,
  readSectionSettings,
  saveDocx,
  sectionSettingsFromXml,
  type SaveBlock,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

const BASE_SECT_PR =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="425"/></w:sectPr>'

describe('w:sep parse (line between columns)', () => {
  it('reads w:sep from an expanded w:cols; self-closing / absent = undefined', () => {
    const withSep =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
      '<w:cols w:num="2" w:space="425"><w:sep/></w:cols></w:sectPr>'
    const s = sectionSettingsFromXml(withSep)
    expect(s.columns).toBe(2)
    expect(s.columnSep).toBe(true)
    expect(sectionSettingsFromXml(BASE_SECT_PR).columnSep).toBeUndefined()
  })

  it('reads w:sep alongside explicit w:col children', () => {
    const xml =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
      '<w:cols w:num="2" w:space="425" w:equalWidth="0"><w:sep/><w:col w:w="6000" w:space="425"/><w:col w:w="4046"/></w:cols></w:sectPr>'
    const s = sectionSettingsFromXml(xml)
    expect(s.columnSep).toBe(true)
    expect(s.colWidths).toEqual([6000, 4046])
  })
})

describe('w:sep serialize (applySectionSettings)', () => {
  it('expands a self-closing w:cols for the separator and collapses it again', () => {
    const base = sectionSettingsFromXml(BASE_SECT_PR)
    const two = applySectionSettings(BASE_SECT_PR, { ...base, columns: 2 })
    const withSep = applySectionSettings(two, {
      ...sectionSettingsFromXml(two),
      columnSep: true,
    })
    expect(withSep).toContain('<w:cols w:num="2" w:space="425"><w:sep/></w:cols>')
    expect(sectionSettingsFromXml(withSep).columnSep).toBe(true)
    const without = applySectionSettings(withSep, {
      ...sectionSettingsFromXml(withSep),
      columnSep: false,
    })
    expect(without).not.toContain('<w:sep/>')
    expect(sectionSettingsFromXml(without).columnSep).toBeUndefined()
  })

  it('keeps w:col children when the separator toggles', () => {
    const xml =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
      '<w:cols w:num="2" w:space="425" w:equalWidth="0"><w:col w:w="6000" w:space="425"/><w:col w:w="4046"/></w:cols></w:sectPr>'
    const withSep = applySectionSettings(xml, {
      ...sectionSettingsFromXml(xml),
      columnSep: true,
    })
    // CT_Columns order: w:sep first, then the w:col list
    expect(withSep).toContain(
      '<w:cols w:num="2" w:space="425" w:equalWidth="0"><w:sep/><w:col w:w="6000" w:space="425"/><w:col w:w="4046"/></w:cols>',
    )
    const without = applySectionSettings(withSep, {
      ...sectionSettingsFromXml(withSep),
      columnSep: false,
    })
    expect(without).toContain('<w:col w:w="6000" w:space="425"/><w:col w:w="4046"/>')
    expect(without).not.toContain('<w:sep/>')
  })

  it('colWidths rebuild emits w:sep and round-trips', () => {
    const base = sectionSettingsFromXml(BASE_SECT_PR)
    const out = applySectionSettings(BASE_SECT_PR, {
      ...base,
      columns: 2,
      colSpace: 720,
      colWidths: [5500, 4046],
      columnSep: true,
    })
    expect(out).toContain(
      '<w:cols w:num="2" w:space="720" w:equalWidth="0"><w:sep/><w:col w:w="5500" w:space="720"/><w:col w:w="4046"/></w:cols>',
    )
    const parsed = sectionSettingsFromXml(out)
    expect(parsed.colWidths).toEqual([5500, 4046])
    expect(parsed.columnSep).toBe(true)
    // byte-stable: applying the same settings again changes nothing
    expect(applySectionSettings(out, parsed)).toBe(out)
  })

  it('creates an expanded w:cols with separator when the sectPr had none', () => {
    const xml =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
    const out = applySectionSettings(xml, {
      ...sectionSettingsFromXml(xml),
      columns: 3,
      columnSep: true,
    })
    // parsed default colSpace is 720 (no w:cols in the source), so the fresh
    // element carries it explicitly
    expect(out).toContain('<w:cols w:num="3" w:space="720"><w:sep/></w:cols>')
  })
})

describe('columns round-trip through saveDocx', () => {
  it('SaveOptions.section writes unequal columns + separator; parses back whole', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('正文') }))
    const blocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    const out = await saveDocx(parsed, blocks, {
      section: {
        ...readSectionSettings(parsed),
        columns: 2,
        colSpace: 425,
        colWidths: [6036, 4046],
        columnSep: true,
      },
    })
    const docXml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string')
    expect(docXml).toContain('w:equalWidth="0"')
    expect(docXml).toContain('<w:sep/>')
    expect(docXml).toContain('<w:col w:w="6036" w:space="425"/>')
    expect(docXml).toContain('<w:col w:w="4046"/>')
    const reparsed = await parseDocx(out)
    const s = readSectionSettings(reparsed)
    expect(s.columns).toBe(2)
    expect(s.colWidths).toEqual([6036, 4046])
    expect(s.columnSep).toBe(true)
  })
})
