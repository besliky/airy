/**
 * BUG-1541: docStyleCss may only emit hyphens:auto when this Chromium build
 * really hyphenates. Stock Electron ships without hyphenation dictionaries,
 * and jsdom has no layout at all, so both must report unsupported; a layout
 * engine that wraps the probe word into several line boxes reports supported.
 */
/** @vitest-environment jsdom */
import { afterEach, describe, expect, it } from 'vitest'
import {
  cssHyphenationSupported,
  resetCssHyphenationProbe,
} from '../src/renderer/hyphenation-support'

interface RectListStub {
  length: number
}

afterEach(() => {
  resetCssHyphenationProbe()
  delete (Range.prototype as Partial<RangePrototypeStub>).getClientRects
})

interface RangePrototypeStub {
  getClientRects?: () => RectListStub
}

function stubClientRects(count: number): void {
  ;(Range.prototype as RangePrototypeStub).getClientRects = () => ({ length: count })
}

describe('cssHyphenationSupported (BUG-1541)', () => {
  it('reports unsupported without a layout engine (jsdom: zero rects)', () => {
    expect(cssHyphenationSupported()).toBe(false)
  })

  it('reports unsupported when the probe word stays on one overflowing line', () => {
    stubClientRects(1)
    resetCssHyphenationProbe()
    expect(cssHyphenationSupported()).toBe(false)
  })

  it('reports supported when the probe word hyphenates into several line boxes', () => {
    stubClientRects(3)
    resetCssHyphenationProbe()
    expect(cssHyphenationSupported()).toBe(true)
  })

  it('caches the probe result and removes the probe element from the DOM', () => {
    stubClientRects(2)
    resetCssHyphenationProbe()
    expect(cssHyphenationSupported()).toBe(true)
    // second call must not re-probe: the DOM stub is gone, so a re-probe
    // would flip the answer back to unsupported
    delete (Range.prototype as RangePrototypeStub).getClientRects
    expect(cssHyphenationSupported()).toBe(true)
    expect(document.body.children.length).toBe(0)
  })
})
