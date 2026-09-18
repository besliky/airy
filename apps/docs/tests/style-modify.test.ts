/**
 * PAR-114 Modify Style: the dialog's edit model flattens into a StyleUpsert
 * (name / font / size / B-I / color / align / spacing / outline level), the
 * engine writes it back into styles.xml (built-ins without w:customStyle,
 * EA fonts and non-auto line rules preserved), and the live document style CSS
 * shows the new definition so paragraphs carrying the pStyle update on screen.
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@airy-office/docx-engine'
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
      rPr: { bold: true, italic: true, color: 'C00000', sizeHalfPoints: 28 },
      pPr: { align: 'center', spaceBeforeTwips: 120, spaceAfterTwips: 120, indentLeftTwips: 720 },
    })
    // dialog keeps fields it does not edit
    expect(upsert.rPr?.font).toBeUndefined()
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

  it('exact line rules and east-asian fonts survive the upsert', async () => {
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
    const { upsert } = styleUpsertFromEdits(info, edits)
    expect(upsert.pPr).toMatchObject({ lineRule: 'exact', lineRawTwips: 480 })
    expect(upsert.rPr).toMatchObject({ font: 'Calibri', fontEa: 'SimSun' })
    const blocks = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, blocks, { styleUpserts: [upsert] })
    const reparsed = await parseDocx(saved)
    expect(reparsed.styles.get('Rule')!.display).toMatchObject({
      lineRule: 'exact',
      lineRawTwips: 480,
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
