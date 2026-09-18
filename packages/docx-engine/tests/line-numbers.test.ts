import { describe, expect, it } from 'vitest'
import {
  applySectionSettings,
  parseDocx,
  readSections,
  readSectionSettings,
  saveDocx,
  sectionSettingsFromXml,
  type SaveBlock,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

const BASE_SECT_PR =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="425"/></w:sectPr>'

describe('w:lnNumType parse', () => {
  it('reads all attributes; absent attributes stay undefined', () => {
    const xml =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:lnNumType w:count="5" w:start="3" w:restart="continuous" w:distance="240"/></w:sectPr>'
    expect(sectionSettingsFromXml(xml).lineNumbers).toEqual({
      countBy: 5,
      start: 3,
      restart: 'continuous',
      distance: 240,
    })
    const minimal = xml.replace(
      ' w:count="5" w:start="3" w:restart="continuous" w:distance="240"',
      '',
    )
    expect(sectionSettingsFromXml(minimal).lineNumbers).toEqual({})
  })

  it('no tag = undefined (numbering off)', () => {
    expect(sectionSettingsFromXml(BASE_SECT_PR).lineNumbers).toBeUndefined()
  })

  it('per-section: paragraph-level sectPr keeps its own settings', async () => {
    const sect1 =
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
      '<w:lnNumType w:restart="newSection"/></w:sectPr></w:pPr></w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('a') + sect1 + P('b') }))
    const sections = readSections(parsed)
    expect(sections).toHaveLength(2)
    expect(sections[0].settings.lineNumbers).toEqual({ restart: 'newSection' })
    expect(sections[1].settings.lineNumbers).toBeUndefined()
  })
})

describe('w:lnNumType serialize (applySectionSettings)', () => {
  it('inserts before pgNumType/cols per CT_SectPr order and replaces an existing tag', () => {
    const withNum = applySectionSettings(BASE_SECT_PR, {
      ...sectionSettingsFromXml(BASE_SECT_PR),
      lineNumbers: { countBy: 5, restart: 'continuous' },
    })
    expect(withNum).toContain('<w:lnNumType w:count="5" w:restart="continuous"/><w:cols')
    // replacing: exactly one tag, updated attrs
    const replaced = applySectionSettings(withNum, {
      ...sectionSettingsFromXml(withNum),
      lineNumbers: { start: 10, distance: 360 },
    })
    expect(replaced.match(/<w:lnNumType/g)).toHaveLength(1)
    expect(replaced).toContain('<w:lnNumType w:start="10" w:distance="360"/><w:cols')
    // round-trips through the parser
    expect(sectionSettingsFromXml(replaced).lineNumbers).toEqual({
      start: 10,
      distance: 360,
    })
  })

  it('undefined removes the tag (numbering off)', () => {
    const withNum = applySectionSettings(BASE_SECT_PR, {
      ...sectionSettingsFromXml(BASE_SECT_PR),
      lineNumbers: { countBy: 2 },
    })
    expect(withNum).toContain('<w:lnNumType w:count="2"/>')
    const without = applySectionSettings(withNum, {
      ...sectionSettingsFromXml(withNum),
      lineNumbers: undefined,
    })
    expect(without).not.toContain('lnNumType')
  })

  it('writes before pgNumType when present (schema order: pgBorders < lnNumType < pgNumType)', () => {
    const xml =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:start="3"/><w:cols w:space="425"/></w:sectPr>'
    const out = applySectionSettings(xml, {
      ...sectionSettingsFromXml(xml),
      lineNumbers: { restart: 'newPage' },
    })
    expect(out).toContain('<w:lnNumType w:restart="newPage"/><w:pgNumType w:start="3"/>')
  })

  it('falls back to after pgMar when only later-schema elements exist (vAlign, no cols)', () => {
    const xml =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:vAlign w:val="center"/></w:sectPr>'
    const out = applySectionSettings(xml, {
      ...sectionSettingsFromXml(xml),
      lineNumbers: { countBy: 3 },
    })
    expect(out).toContain('<w:lnNumType w:count="3"/><w:vAlign w:val="center"/>')
  })

  it('falls back to after pgBorders when no later-schema element exists (BUG-914: pgBorders sits between pgMar and lnNumType)', () => {
    const xml =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
      '<w:pgBorders w:offsetFrom="page">' +
      '<w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/>' +
      '<w:left w:val="single" w:sz="4" w:space="24" w:color="auto"/>' +
      '<w:bottom w:val="single" w:sz="4" w:space="24" w:color="auto"/>' +
      '<w:right w:val="single" w:sz="4" w:space="24" w:color="auto"/>' +
      '</w:pgBorders></w:sectPr>'
    const out = applySectionSettings(xml, {
      ...sectionSettingsFromXml(xml),
      lineNumbers: { restart: 'newPage' },
    })
    // pgMar < pgBorders < lnNumType per CT_SectPr (applySectionSettings rebuilds
    // the visible border right after pgMar; lnNumType must land after it)
    expect(out.indexOf('<w:pgMar')).toBeLessThan(out.indexOf('<w:pgBorders'))
    expect(out).toContain('</w:pgBorders><w:lnNumType w:restart="newPage"/>')
    expect(sectionSettingsFromXml(out).lineNumbers).toEqual({ restart: 'newPage' })
  })

  it('falls back to after paperSrc when pgBorders is absent (paperSrc sits between pgMar and lnNumType)', () => {
    const xml =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:paperSrc w:first="1" w:other="2"/></w:sectPr>'
    const out = applySectionSettings(xml, {
      ...sectionSettingsFromXml(xml),
      lineNumbers: { countBy: 2 },
    })
    expect(out).toContain('<w:paperSrc w:first="1" w:other="2"/><w:lnNumType w:count="2"/>')
  })
})

describe('w:lnNumType round-trip through saveDocx', () => {
  it('SaveOptions.section writes the final section and parses back with all options', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('正文') }))
    const blocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const settings = readSectionSettings(parsed)
    const saved = await saveDocx(parsed, blocks, {
      section: {
        ...settings,
        lineNumbers: { countBy: 5, start: 2, restart: 'continuous', distance: 240 },
      },
    })
    const reparsed = await parseDocx(saved)
    expect(readSectionSettings(reparsed).lineNumbers).toEqual({
      countBy: 5,
      start: 2,
      restart: 'continuous',
      distance: 240,
    })
  })

  it('byte-stable: re-applying parsed settings leaves the sectPr unchanged', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('a') }))
    const settings = readSectionSettings(parsed)
    const withLn = applySectionSettings(BASE_SECT_PR, { ...settings, lineNumbers: {} })
    const again = applySectionSettings(withLn, sectionSettingsFromXml(withLn))
    expect(again).toBe(withLn)
  })
})
