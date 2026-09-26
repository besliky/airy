/**
 * BUG-1765 render smoke: glyph runs of a mixed emoji+CJK+Cyrillic LTR deck must
 * reach Konva without a direction=rtl stamp (neutrals never flip LTR text), the
 * Cyrillic word stays one contiguous LTR unit in logical order, and genuine RTL
 * runs (Arabic) still carry direction=rtl.
 */
import { describe, expect, it } from 'vitest'
import {
  HeuristicMetrics,
  layoutText,
  makeViewport,
  runDrawsRtl,
  type RunStyle,
} from '@airy-office/pptx-render'
import { glyphToDraw, layoutGlyphs } from '../src/renderer/konva-adapter'
import type { RenderTextLayout } from '@airy-office/pptx-render'

const metrics = new HeuristicMetrics()
const vp = makeViewport({ cx: 9525 * 1000, cy: 9525 * 1000 }, 1000)

function layout(paragraphs: Array<{ text: string; rtl?: boolean }>): RenderTextLayout {
  return layoutText({
    body: {
      anchor: 'top',
      insets: { l: 0, t: 0, r: 0, b: 0 },
      autofit: 'none',
      wrap: false,
      paragraphs: paragraphs.map((p) => ({
        runs: [{ text: p.text, style: monoStyle() }],
        ...(p.rtl != null ? { rtl: p.rtl } : {}),
      })),
    },
    boxWidthPx: 4000,
    boxHeightPx: 300,
    metrics,
    vp,
  })
}

const style: RunStyle = { fontFamily: 'Courier New', fontSizePx: 48, bold: false, italic: false }
function monoStyle(): RunStyle {
  return { ...style }
}

describe('BUG-1765 run direction reaches Konva from strong bidi characters', () => {
  it('audit vector "Mixed 🚀тест 中文 end": no GlyphDraw carries direction=rtl and Cyrillic stays in logical order', () => {
    const draws = layoutGlyphs(layout([{ text: 'Mixed 🚀тест 中文 end' }]))
    expect(draws.length).toBeGreaterThan(0)
    expect(draws.every((g) => g.direction === undefined)).toBe(true)
    // The Cyrillic word is one contiguous LTR glyph run, in logical letter order
    const word = draws.filter((g) => g.text.includes('т'))
    expect(word.map((g) => g.text).join('')).toContain('тест')
    expect(draws.some((g) => g.text === 'тсет')).toBe(false)
  })

  it('pure LTR and pure Cyrillic runs never draw rtl', () => {
    for (const text of ['Hello world', 'Слайд экспорт № 42', '中文测试 no flip']) {
      const draws = layoutGlyphs(layout([{ text }]))
      expect(draws.every((g) => g.direction === undefined)).toBe(true)
    }
  })

  it('Arabic runs keep direction=rtl on the Konva draw', () => {
    const draws = layoutGlyphs(layout([{ text: 'السلام عليكم', rtl: true }]))
    expect(draws.length).toBeGreaterThan(0)
    expect(draws.every((g) => g.direction === 'rtl')).toBe(true)
  })

  it('glyphToDraw maps the run rtl flag to direction=rtl only (no stamp on plain runs)', () => {
    const plain = glyphToDraw({
      ...baseGlyphRun(),
      text: '🚀тест 中文 end',
    })
    expect(plain.direction).toBeUndefined()
    const arabic = glyphToDraw({ ...baseGlyphRun(), text: 'سلام', rtl: true })
    expect(arabic.direction).toBe('rtl')
  })

  it('runDrawsRtl: strong characters decide — neutrals and emoji never flip an LTR run', () => {
    // Audit vector segments (LTR deck): no odd level anywhere
    expect(runDrawsRtl('Mixed 🚀тест 中文 end', undefined)).toBe(false)
    expect(runDrawsRtl('Mixed 🚀тест 中文 end', 0)).toBe(false)
    // Strong RTL justifies the stamp
    expect(runDrawsRtl('السلام عليكم', 1)).toBe(true)
    // Pure neutrals keep the odd-level stamp (paired-bracket mirroring parity)
    expect(runDrawsRtl('🚀', 1)).toBe(true)
    // Strong-LTR text mixed with neutrals never draws rtl, whatever the level
    expect(runDrawsRtl('🚀тест', 1)).toBe(false)
  })
})

function baseGlyphRun() {
  return {
    text: '',
    x: 0,
    baselineY: 24,
    fontFamily: 'Courier New',
    fontSizePx: 48,
    color: '#000000',
    bold: false,
    italic: false,
    underline: false,
    widthPx: 100,
    ascentPx: 38,
  }
}
