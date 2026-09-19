/**
 * PAR-114 Modify Style: the dialog's edit model flattens into a StyleUpsert
 * (name / font / size / B-I / color / align / spacing / outline level), the
 * engine writes it back into styles.xml (built-ins without w:customStyle,
 * EA fonts and non-auto line rules preserved), and the live document style CSS
 * shows the new definition so paragraphs carrying the pStyle update on screen.
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, saveDocx, type StyleUpsert } from '@airy-office/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { styleEditsFromInfo, styleUpsertFromEdits } from '../src/renderer/components/StyleDialog'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const STYLES =
  '<w:style w:type="paragraph" w:styleId="MyQuote"><w:name w:val="My Quote"/>' +
  '<w:pPr><w:spacing w:before="120" w:after="120"/><w:ind w:left="720"/></w:pPr>' +
  '<w:rPr><w:i/><w:color w:val="595959"/><w:sz w:val="22"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/>' +
  '<w:pPr><w:outlineLvl w:val="3"/></w:pPr></w:style>'

async function openDoc() {
  const parsed = await parseDocx(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>', extraStylesXml: STYLES }),
  )
  return parsed
}

describe('styleEditsFromInfo seeds', () => {
  it('reads the style definition incl. docDefaults size fallback', async () => {
    const parsed = await openDoc()
    expect(styleEditsFromInfo(parsed.styles.get('MyQuote'))).toMatchObject({
      name: 'My Quote',
      italic: true,
      color: '595959',
      sizePt: 11,
      beforePt: 6,
      afterPt: 6,
      outline: 0,
    })
    expect(styleEditsFromInfo(parsed.styles.get('Heading4')).outline).toBe(4)
  })
})

describe('styleUpsertFromEdits', () => {
  it('builds the flattened upsert and the next live display', async () => {
    const parsed = await openDoc()
    const info = parsed.styles.get('MyQuote')!
    const edits = {
      ...styleEditsFromInfo(info),
      bold: true,
      sizePt: 14,
      color: 'C00000',
      align: 'center' as const,
    }
    const { upsert, display, headingLevel } = styleUpsertFromEdits(info, edits)
    expect(upsert).toMatchObject({
      styleId: 'MyQuote',
      type: 'paragraph',
      name: 'My Quote',
      rPr: { bold: true, color: 'C00000', sizeHalfPoints: 28 },
      pPr: { align: 'center' },
    })
    // BUG-1102: untouched fields stay unset — the engine keeps their bytes and
    // the basedOn chain keeps supplying inherited values
    expect(upsert.rPr?.italic).toBeUndefined()
    expect(upsert.rPr?.font).toBeUndefined()
    expect(upsert.pPr?.spaceBeforeTwips).toBeUndefined()
    expect(upsert.pPr?.spaceAfterTwips).toBeUndefined()
    expect(upsert.pPr?.indentLeftTwips).toBeUndefined()
    expect(display).toMatchObject({ bold: true, italic: true, color: 'C00000', sizeHalfPoints: 28 })
    expect(headingLevel).toBeNull()
  })

  it('outline level and dual fonts round through the model', async () => {
    const parsed = await openDoc()
    const info = parsed.styles.get('MyQuote')!
    const edits = {
      ...styleEditsFromInfo(info),
      outline: 2,
      font: 'Calibri',
      fontEa: 'SimSun',
    }
    const { upsert, headingLevel } = styleUpsertFromEdits(info, edits)
    expect(upsert.pPr?.outlineLevel).toBe(2)
    expect(upsert.rPr).toMatchObject({ font: 'Calibri', fontEa: 'SimSun' })
    expect(headingLevel).toBe(2)
  })
})

describe('styleUpserts save round-trip (engine)', () => {
  it('writes the modified definition and no w:customStyle for built-ins', async () => {
    const parsed = await openDoc()
    const info = parsed.styles.get('MyQuote')!
    const { upsert } = styleUpsertFromEdits(info, {
      ...styleEditsFromInfo(info),
      sizePt: 14,
      outline: 3,
    })
    const blocks = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, blocks, { styleUpserts: [upsert] })
    const zip = await JSZip.loadAsync(saved)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    expect(stylesXml.match(/w:styleId="MyQuote"/g)).toHaveLength(1)
    expect(stylesXml).toContain('<w:sz w:val="28"/>')
    expect(stylesXml).toContain('<w:outlineLvl w:val="2"/>')
    expect(stylesXml).toContain('w:customStyle="1"')
    const reparsed = await parseDocx(saved)
    expect(reparsed.styles.get('MyQuote')!.display).toMatchObject({
      sizeHalfPoints: 28,
      italic: true,
    })
    expect(reparsed.styles.get('MyQuote')!.headingLevel).toBe(3)
  })

  it('a built-in Heading4 modify keeps the built-in marker and outline level', async () => {
    const parsed = await openDoc()
    const info = parsed.styles.get('Heading4')!
    const { upsert } = styleUpsertFromEdits(info, { ...styleEditsFromInfo(info), bold: true })
    const blocks = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, blocks, { styleUpserts: [upsert] })
    const zip = await JSZip.loadAsync(saved)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    const heading4 = /<w:style [^>]*w:styleId="Heading4"[\s\S]*?<\/w:style>/.exec(stylesXml)![0]
    expect(heading4).not.toContain('w:customStyle')
    expect(heading4).toContain('<w:outlineLvl w:val="3"/>')
    expect(heading4).toContain('<w:b/>')
    const reparsed = await parseDocx(saved)
    expect(reparsed.styles.get('Heading4')).toMatchObject({ headingLevel: 4 })
    expect(reparsed.styles.get('Heading4')!.display?.bold).toBe(true)
  })

  it('exact line rules and east-asian fonts survive an untouched modify', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        extraStylesXml:
          '<w:style w:type="paragraph" w:styleId="Rule"><w:name w:val="Rule"/>' +
          '<w:pPr><w:spacing w:line="480" w:lineRule="exact"/></w:pPr>' +
          '<w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="SimSun"/></w:rPr></w:style>',
      }),
    )
    const info = parsed.styles.get('Rule')!
    const edits = styleEditsFromInfo(info)
    expect(edits.lineSpacing).toBe(0) // exact rule: no multiple to edit
    // BUG-1102: nothing changed → nothing but the name is written; the exact
    // line rule and both font slots are no longer flattened into the upsert
    const { upsert } = styleUpsertFromEdits(info, edits)
    expect(upsert.pPr?.lineRule).toBeUndefined()
    expect(upsert.pPr?.lineRawTwips).toBeUndefined()
    expect(upsert.pPr?.lineSpacing).toBeUndefined()
    expect(upsert.rPr?.font).toBeUndefined()
    expect(upsert.rPr?.fontEa).toBeUndefined()
    const saved = await saveDocWith(parsed, [upsert])
    const reparsed = await parseDocx(saved)
    expect(reparsed.styles.get('Rule')!.display).toMatchObject({
      lineRule: 'exact',
      lineRawTwips: 480,
      fontAscii: 'Calibri',
      font: 'SimSun',
    })
  })
})

describe('live display update', () => {
  it('the swapped style definition reaches the regenerated document style CSS', async () => {
    const parsed = await openDoc()
    const info = parsed.styles.get('MyQuote')!
    const { display } = styleUpsertFromEdits(info, {
      ...styleEditsFromInfo(info),
      color: '0070C0',
      sizePt: 14,
    })
    // the App-side live update: replace the map entry, regenerate the CSS
    parsed.styles.set('MyQuote', { ...info, display })
    const css = docStyleCss(parsed)
    const rule = css.match(/\[data-style="MyQuote"\][^{]*\{[^}]*\}/g) ?? []
    expect(rule.join('\n')).toContain('#0070C0')
    expect(rule.join('\n')).toContain('font-size:14pt')
  })
})

const DOC_DEFAULTS_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:styleId="Heading9"><w:name w:val="heading 9"/>' +
  '<w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="8"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="OwnSpacing"><w:name w:val="Own Spacing"/>' +
  '<w:pPr><w:spacing w:before="120" w:after="80"/><w:jc w:val="center"/><w:outlineLvl w:val="2"/></w:pPr>' +
  '<w:rPr><w:b/><w:color w:val="FF0000"/><w:sz w:val="28"/></w:rPr></w:style>' +
  '</w:styles>'

async function openDefaultsDoc() {
  return parseDocx(
    await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      stylesXml: DOC_DEFAULTS_STYLES,
    }),
  )
}

async function saveDocWith(parsed: Awaited<ReturnType<typeof parseDocx>>, upserts: StyleUpsert[]) {
  const blocks = parsed.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
  return saveDocx(parsed, blocks, { styleUpserts: upserts })
}

describe('BUG-1002: spacing the chain does not define stays unset', () => {
  it('seeding 0pt never serializes w:before/after="0" over docDefaults', async () => {
    const parsed = await openDefaultsDoc()
    expect(parsed.docDefaults?.spaceAfterTwips).toBe(160)
    const info = parsed.styles.get('Heading9')!
    const seed = styleEditsFromInfo(info, parsed.docDefaults)
    expect(seed.beforePt).toBe(0)
    expect(seed.afterPt).toBe(0)
    // modify only the font — the untouched spacing facets stay out of the upsert
    const { upsert, display } = styleUpsertFromEdits(info, { ...seed, font: 'Calibri' })
    expect(upsert.pPr?.spaceBeforeTwips).toBeUndefined()
    expect(upsert.pPr?.spaceAfterTwips).toBeUndefined()
    expect(display?.spaceBeforeTwips).toBeUndefined()
    expect(display?.spaceAfterTwips).toBeUndefined()
    const saved = await saveDocWith(parsed, [upsert])
    const zip = await JSZip.loadAsync(saved)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    const heading9 = /<w:style [^>]*w:styleId="Heading9"[\s\S]*?<\/w:style>/.exec(stylesXml)![0]
    expect(heading9).not.toContain('<w:spacing')
    // BUG-1001 (docs side): the unedited keepNext/keepLines/outline survive too
    expect(heading9).toContain('<w:keepNext/><w:keepLines/><w:outlineLvl w:val="8"/>')
    // the dialog seeds no distinct EA face, so only the latin slots are written
    expect(heading9).toContain('<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>')
    const reparsed = await parseDocx(saved)
    expect(reparsed.docDefaults?.spaceAfterTwips).toBe(160)
    expect(reparsed.styles.get('Heading9')!.display?.spaceAfterTwips).toBeUndefined()
  })

  it('an explicit user zero is kept as a real override', async () => {
    const parsed = await openDefaultsDoc()
    const info = parsed.styles.get('OwnSpacing')!
    const seed = styleEditsFromInfo(info)
    expect(seed.beforePt).toBe(6)
    const { upsert } = styleUpsertFromEdits(info, { ...seed, beforePt: 0 })
    expect(upsert.pPr?.spaceBeforeTwips).toBe(0)
    // untouched after stays out of the upsert (BUG-1102) and in the file
    expect(upsert.pPr?.spaceAfterTwips).toBeUndefined()
    const saved = await saveDocWith(parsed, [upsert])
    const zip = await JSZip.loadAsync(saved)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    const own = /<w:style [^>]*w:styleId="OwnSpacing"[\s\S]*?<\/w:style>/.exec(stylesXml)![0]
    expect(own).toContain('<w:spacing w:before="0" w:after="80"/>')
  })

  it('clearing a facet removes it from the definition instead of leaving it', async () => {
    const parsed = await openDefaultsDoc()
    const info = parsed.styles.get('OwnSpacing')!
    const seed = styleEditsFromInfo(info)
    const { upsert } = styleUpsertFromEdits(info, {
      ...seed,
      bold: false,
      italic: true,
      color: '',
      sizePt: 0,
      align: '',
      outline: 0,
    })
    expect(upsert.rPr?.bold).toBeNull()
    expect(upsert.rPr?.italic).toBe(true)
    expect(upsert.rPr?.color).toBeNull()
    expect(upsert.rPr?.sizeHalfPoints).toBeNull()
    expect(upsert.pPr?.align).toBeNull()
    // BUG-1101: outline clears as an explicit body-text marker (w:outlineLvl 9)
    expect(upsert.pPr?.outlineLevel).toBe(false)
    // BUG-1102: the untouched interval stays out of the upsert (and in the file)
    expect(upsert.pPr?.spaceAfterTwips).toBeUndefined()
    const saved = await saveDocWith(parsed, [upsert])
    const zip = await JSZip.loadAsync(saved)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    const own = /<w:style [^>]*w:styleId="OwnSpacing"[\s\S]*?<\/w:style>/.exec(stylesXml)![0]
    expect(own).not.toContain('<w:b/>')
    expect(own).not.toContain('<w:color')
    expect(own).not.toContain('<w:sz ')
    expect(own).not.toContain('<w:jc ')
    expect(own).toContain('<w:outlineLvl w:val="9"/>')
    expect(own).toContain('<w:i/>')
    expect(own).toContain('<w:spacing w:before="120" w:after="80"/>')
  })

  it('a style without facets keeps its definition when nothing is set', async () => {
    const parsed = await openDefaultsDoc()
    const info = parsed.styles.get('Plain')!
    const { upsert } = styleUpsertFromEdits(
      info,
      styleEditsFromInfo(info, parsed.docDefaults),
      parsed.docDefaults,
    )
    const saved = await saveDocWith(parsed, [upsert])
    const zip = await JSZip.loadAsync(saved)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    const plain = /<w:style [^>]*w:styleId="Plain"[\s\S]*?<\/w:style>/.exec(stylesXml)![0]
    // BUG-1102: no pPr/rPr is invented — even the docDefaults-seeded size is
    // no longer flattened into an untouched style
    expect(plain).toBe(
      '<w:style w:type="paragraph" w:styleId="Plain" w:customStyle="1">' +
        '<w:name w:val="Plain"/></w:style>',
    )
    expect(upsert.pPr?.outlineLevel).toBeUndefined()
    expect(upsert.rPr?.sizeHalfPoints).toBeUndefined()
  })
})

const HEADING_CHAIN_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr>' +
  '<w:rPr><w:b/><w:i/><w:color w:val="2E74B5"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>' +
  '<w:basedOn w:val="Heading1"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr>' +
  '<w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>' +
  '</w:styles>'

async function openChainDoc() {
  return parseDocx(
    await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      stylesXml: HEADING_CHAIN_STYLES,
    }),
  )
}

describe('BUG-1101: clearing an inherited facet writes an explicit off', () => {
  it('unchecking Bold on Heading2 (basedOn a bold Heading1) survives the reopen', async () => {
    const parsed = await openChainDoc()
    const info = parsed.styles.get('Heading2')!
    // the seed is the resolved view: bold/italic/color all arrive via Heading1
    expect(info.display).toMatchObject({ bold: true, italic: true, color: '2E74B5' })
    expect(info.ownDisplay?.bold).toBe(true)
    expect(info.chainDisplay?.bold).toBe(true)
    const seed = styleEditsFromInfo(info, parsed.docDefaults)
    const { upsert, display } = styleUpsertFromEdits(info, {
      ...seed,
      bold: false,
      italic: false,
      color: '',
    })
    // not null: a removal would let Heading1 re-supply everything after reopen
    expect(upsert.rPr?.bold).toBe(false)
    expect(upsert.rPr?.italic).toBe(false)
    expect(upsert.rPr?.color).toBe('auto')
    // the live view already shows the cleared state, matching the file
    expect(display).toMatchObject({ bold: false, italic: false, color: 'auto' })
    parsed.styles.set('Heading2', { ...info, display })
    const css = docStyleCss(parsed)
    const rule = css.match(/\[data-style="Heading2"\][^{]*\{[^}]*\}/g) ?? []
    expect(rule.join('\n')).toContain('font-weight:400')
    expect(rule.join('\n')).toContain('font-style:normal')
    expect(rule.join('\n')).toContain('color:var(--docs-paper-ink)')
    // save → reopen: the explicit offs win over the chain
    const saved = await saveDocWith(parsed, [upsert])
    const zip = await JSZip.loadAsync(saved)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    const h2 = /<w:style [^>]*w:styleId="Heading2"[\s\S]*?<\/w:style>/.exec(stylesXml)![0]
    expect(h2).toContain('<w:b w:val="0"/>')
    expect(h2).toContain('<w:i w:val="0"/>')
    expect(h2).toContain('<w:color w:val="auto"/>')
    const reparsed = await parseDocx(saved)
    expect(reparsed.styles.get('Heading2')!.display).toMatchObject({
      bold: false,
      italic: false,
      color: 'auto',
    })
    // Heading1 itself keeps its facets
    expect(reparsed.styles.get('Heading1')!.display).toMatchObject({
      bold: true,
      italic: true,
      color: '2E74B5',
    })
  })

  it('an own-only facet still clears by removal', async () => {
    const parsed = await openChainDoc()
    const info = parsed.styles.get('Heading2')!
    const seed = styleEditsFromInfo(info, parsed.docDefaults)
    // size is Heading2's own; the chain does not supply one — removal is enough
    const { upsert, display } = styleUpsertFromEdits(info, { ...seed, sizePt: 0 })
    expect(upsert.rPr?.sizeHalfPoints).toBeNull()
    expect(display?.sizeHalfPoints).toBeUndefined()
  })
})

const CHAIN_FLATTEN_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Parent"><w:name w:val="Parent"/>' +
  '<w:pPr><w:spacing w:after="160"/><w:ind w:left="720"/></w:pPr>' +
  '<w:rPr><w:b/><w:u w:val="single"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Child"><w:name w:val="Child"/>' +
  '<w:basedOn w:val="Parent"/></w:style>' +
  '</w:styles>'

describe('BUG-1102: a modify no longer flattens chain values into the style', () => {
  async function openFlattenDoc() {
    return parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        stylesXml: CHAIN_FLATTEN_STYLES,
      }),
    )
  }

  it('an untouched modify leaves the child style bare — the parent keeps supplying', async () => {
    const parsed = await openFlattenDoc()
    const info = parsed.styles.get('Child')!
    // the seed resolves everything through Parent (spacing, indent, bold, underline)
    expect(info.display).toMatchObject({
      bold: true,
      underline: true,
      spaceAfterTwips: 160,
      indentLeftTwips: 720,
    })
    const { upsert } = styleUpsertFromEdits(info, styleEditsFromInfo(info))
    expect(upsert.pPr?.spaceAfterTwips).toBeUndefined()
    expect(upsert.pPr?.indentLeftTwips).toBeUndefined()
    expect(upsert.rPr?.bold).toBeUndefined()
    expect(upsert.rPr?.underline).toBeUndefined()
    const saved = await saveDocWith(parsed, [upsert])
    const zip = await JSZip.loadAsync(saved)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    const child = /<w:style [^>]*w:styleId="Child"[\s\S]*?<\/w:style>/.exec(stylesXml)![0]
    // none of Parent's facets were pinned into the child definition
    expect(child).not.toContain('<w:spacing')
    expect(child).not.toContain('<w:ind ')
    expect(child).not.toContain('<w:b/>')
    expect(child).not.toContain('<w:u ')
    const reparsed = await parseDocx(saved)
    expect(reparsed.styles.get('Child')!.display).toMatchObject({
      bold: true,
      underline: true,
      spaceAfterTwips: 160,
    })
  })

  it('a changed spacing interval becomes the child override', async () => {
    const parsed = await openFlattenDoc()
    const info = parsed.styles.get('Child')!
    const seed = styleEditsFromInfo(info)
    expect(seed.afterPt).toBe(8) // Parent's 160 twips
    const { upsert, display } = styleUpsertFromEdits(info, { ...seed, afterPt: 3 })
    expect(upsert.pPr?.spaceAfterTwips).toBe(60)
    expect(upsert.pPr?.spaceBeforeTwips).toBeUndefined()
    expect(display?.spaceAfterTwips).toBe(60)
    const saved = await saveDocWith(parsed, [upsert])
    const zip = await JSZip.loadAsync(saved)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    const child = /<w:style [^>]*w:styleId="Child"[\s\S]*?<\/w:style>/.exec(stylesXml)![0]
    expect(child).toContain('<w:spacing w:after="60"/>')
    const reparsed = await parseDocx(saved)
    expect(reparsed.styles.get('Child')!.display?.spaceAfterTwips).toBe(60)
  })
})
