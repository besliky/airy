import { describe, expect, it } from 'vitest'
import { warpGlyphs } from '../src/renderer/text-warp'
import type { GlyphDraw } from '../src/renderer/konva-adapter'

const glyph = (text: string, x: number): GlyphDraw => ({
  text,
  x,
  y: 0,
  fontSize: 20,
  fontFamily: 'Arial',
  fill: '#000',
  fontStyle: 'normal',
  textDecoration: '',
})
const measure = (text: string) => [...text].length * 10

describe('warpGlyphs', () => {
  it('unknown presets return null (caller keeps the straight layout)', () => {
    expect(warpGlyphs([glyph('Hi', 0)], 100, 40, { prst: 'textNoShape' }, measure)).toBeNull()
    expect(warpGlyphs([glyph('Hi', 0)], 100, 40, { prst: 'textRingInside' }, measure)).toBeNull()
  })

  it('wave1 lifts the quarter-way character and drops the three-quarter one', () => {
    const out = warpGlyphs([glyph('abcd', 0)], 200, 40, { prst: 'textWave1' }, measure)!
    // repacked at 10px/char, centers at u=1/8,3/8,5/8,7/8: sin>0 → +y for u<0.5
    expect(out[1]!.y).toBeGreaterThan(out[2]!.y)
    expect(out[0]!.rotation).toBe(0) // waves stay upright
  })

  it('arch rotates characters along the tangent and centers the block', () => {
    const out = warpGlyphs([glyph('abcd', 0)], 200, 40, { prst: 'textArchUp' }, measure)!
    expect(out[0]!.rotation!).toBeLessThan(0) // rising into the arch
    expect(out[3]!.rotation!).toBeGreaterThan(0)
    const centers = out.map((g) => g.x)
    expect((centers[0]! + centers[3]!) / 2).toBeCloseTo(100, 0) // centered in the box
  })

  it('triangleInverted squeezes mid-box characters via scaleY', () => {
    const out = warpGlyphs([glyph('abcd', 0)], 200, 40, { prst: 'textTriangleInverted' }, measure)!
    expect(out[1]!.scaleY!).toBeLessThan(out[0]!.scaleY!)
  })

  it('splits multi-char runs and skips whitespace glyphs', () => {
    const out = warpGlyphs([glyph('a b', 0)], 100, 40, { prst: 'textWave1' }, measure)!
    expect(out.map((g) => g.text)).toEqual(['a', 'b'])
    expect(out[1]!.x).toBeGreaterThan(out[0]!.x + 10) // the swallowed space keeps its advance
  })

  it('circle marches the characters around an ellipse, upright over the top', () => {
    const out = warpGlyphs([glyph('abcde', 0)], 100, 80, { prst: 'textCircle' }, measure)!
    // all characters land on the ellipse path (inside the box, touching the em-inset ring)
    for (const g of out) {
      expect(g.x).toBeGreaterThanOrEqual(10)
      expect(g.x).toBeLessThanOrEqual(90)
      expect(g.y).toBeGreaterThanOrEqual(10)
      expect(g.y).toBeLessThanOrEqual(70)
    }
    // clockwise over the top: left → top (char 1 is the highest) → right edge
    expect(out[1]!.y).toBeLessThan(out[0]!.y)
    expect(out[2]!.x).toBeGreaterThan(out[1]!.x)
    expect(out[2]!.x).toBeGreaterThan(out[0]!.x)
    // the top character is upright and scaled to cover the circumference
    expect(Math.abs(out[1]!.rotation ?? 0)).toBeLessThan(90)
    expect(out[1]!.scaleX!).toBeGreaterThan(1)
    // the walk continues down the far side: the last char is below the first
    expect(out[4]!.y).toBeGreaterThan(out[0]!.y)
  })

  it('button keeps the middle full height and squeezes the rounded ends', () => {
    const out = warpGlyphs([glyph('abcde', 0)], 200, 40, { prst: 'textButton' }, measure)!
    expect(out[2]!.scaleY!).toBeGreaterThan(out[0]!.scaleY!)
    expect(out[2]!.scaleY!).toBeGreaterThan(out[4]!.scaleY!)
    // vertical centers stay on the box midline (symmetric envelope)
    for (const g of out) expect(Math.abs(g.y - 20)).toBeLessThan(0.001)
  })
})
