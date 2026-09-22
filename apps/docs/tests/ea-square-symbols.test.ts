/**
 * BUG-1543: the CJK-encoding geometric shapes (U+25CB and siblings, the CJK
 * date-line class: kanji numerals around a U+25CB circle) are East Asian
 * characters for Word — they render with the run's eastAsia font at a
 * fullwidth (1em) advance. Chromium resolves them through the Latin head of
 * the font chain (Carlito U+25CB = 0.55em), so the renderer wraps them in
 * EA-font spans and the width engine counts them as fullwidth CJK.
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@airy-office/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { runSpanSpecs } from '../src/renderer/editor/protected-render'
import { HeuristicMetrics, eaSquareSymbolRanges } from '../src/renderer/line-metrics'

const STYLE_12PT = { fontFamily: 'Calibri', fontSizePx: 16, bold: false, italic: false }

describe('eaSquareSymbolRanges', () => {
  it('wraps the U+25CB of a CJK date line', () => {
    expect(eaSquareSymbolRanges('二○二六年七月二十六日')).toEqual([{ from: 1, to: 2 }])
  })

  it('groups adjacent symbols and finds every stretch', () => {
    expect(eaSquareSymbolRanges('○●☆ 記録 ■完了')).toEqual([
      { from: 0, to: 3 },
      { from: 7, to: 8 },
    ])
  })

  it('covers the GB2312/JIS core geometric set', () => {
    for (const ch of '■□▲△◆◇○◎●★☆') {
      expect(eaSquareSymbolRanges(ch)).toEqual([{ from: 0, to: 1 }])
    }
  })

  it('Latin-context symbols stay on the Latin chain', () => {
    expect(eaSquareSymbolRanges('№ 470-EL')).toEqual([])
    expect(eaSquareSymbolRanges('plain text 12%')).toEqual([])
    expect(eaSquareSymbolRanges('')).toEqual([])
  })
})

describe('EA square symbol width (BUG-1543)', () => {
  const metrics = new HeuristicMetrics()

  it('U+25CB measures as fullwidth CJK, not the 0.52em Latin fallback', () => {
    // doc 10 measured 6.59pt at 12pt (0.55em Latin face); Word/LO render 12.00pt
    expect(metrics.measure('○', STYLE_12PT)).toBe(metrics.measure('二', STYLE_12PT))
    expect(metrics.measure('○', STYLE_12PT)).toBe(16)
    expect(metrics.measure('★', STYLE_12PT)).toBe(16)
  })

  it('the whole doc 10 date line sums to 11 em (11 chars at 1em)', () => {
    expect(metrics.measure('二○二六年七月二十六日', STYLE_12PT)).toBe(11 * 16)
  })
})

describe('EA square symbol rendering', () => {
  it('the editor decorates symbol stretches with the EA-font span', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>二○二六年七月二十六日</w:t></w:r></w:p>' +
          '<w:p><w:r><w:t>plain Latin</w:t></w:r></w:p>',
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.storage.listNumbering.defs = parsed.numbering
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const spans = Array.from(editor.view.dom.querySelectorAll('.doc-ea-symbol'))
    expect(spans.map((s) => s.textContent)).toEqual(['○'])
    editor.destroy()
  })

  it('protected-render wraps the same stretches (textbox/header paths)', () => {
    const specs = runSpanSpecs({ text: '二○二六年度' })
    expect(JSON.stringify(specs)).toContain('["span",{"class":"doc-ea-symbol"},"○"]')
    expect(JSON.stringify(runSpanSpecs({ text: 'no symbols' }))).not.toContain('doc-ea-symbol')
  })
})
