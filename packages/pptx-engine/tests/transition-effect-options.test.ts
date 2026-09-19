import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import type { SlideTransitionDir, SlideTransitionKind } from '../src/index'
import {
  getSlideTransitionSpec,
  openPptx,
  readSlideTransitionSpecXml,
  patchSlideTransitionXml,
  savePptx,
  setSlideTransition,
  TRANSITION_DIR_INFO,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const BODY = '</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'

/** Every (kind, options) combination the Effect Options UI can produce. */
function allOptionCombos(): Array<{
  kind: SlideTransitionKind
  dir?: SlideTransitionDir
  orient?: 'horz' | 'vert'
  durationMs?: number
}> {
  const out: Array<{
    kind: SlideTransitionKind
    dir?: SlideTransitionDir
    orient?: 'horz' | 'vert'
    durationMs?: number
  }> = [{ kind: 'none' }, { kind: 'fade', durationMs: 1250 }]
  for (const kind of Object.keys(TRANSITION_DIR_INFO) as SlideTransitionKind[]) {
    const { dirs } = TRANSITION_DIR_INFO[kind]
    if (!dirs.length) continue
    out.push({ kind, dir: dirs[0]! })
    out.push({ kind, dir: dirs[dirs.length - 1]!, durationMs: 300 })
    if (kind === 'split') out.push({ kind, orient: 'vert', dir: 'in', durationMs: 2000 })
  }
  return out
}

describe('transition Effect Options: direction serialization', () => {
  it('serializes side directions to the OOXML travel dir (From Bottom travels up)', () => {
    for (const [from, xml] of [
      ['fromBottom', 'u'],
      ['fromTop', 'd'],
      ['fromLeft', 'r'],
      ['fromRight', 'l'],
    ] as const) {
      const s = patchSlideTransitionXml(BODY, 'push', { dir: from })
      expect(s).toContain(`<p:push dir="${xml}"/>`)
      expect(readSlideTransitionSpecXml(s).dir).toBe(from)
    }
  })

  it('serializes diagonal directions for cover/pull and reads them back', () => {
    for (const [from, xml] of [
      ['fromBottomLeft', 'ru'],
      ['fromBottomRight', 'lu'],
      ['fromTopLeft', 'rd'],
      ['fromTopRight', 'ld'],
    ] as const) {
      for (const kind of ['cover', 'pull'] as const) {
        const s = patchSlideTransitionXml(BODY, kind, { dir: from })
        expect(s).toContain(`<p:${kind} dir="${xml}"/>`)
        expect(readSlideTransitionSpecXml(s).dir).toBe(from)
      }
    }
  })

  it('serializes split orientation + in/out and zoom in/out', () => {
    const vertIn = patchSlideTransitionXml(BODY, 'split', { orient: 'vert', dir: 'in' })
    expect(vertIn).toContain('<p:split orient="vert" dir="in"/>')
    expect(readSlideTransitionSpecXml(vertIn)).toMatchObject({
      kind: 'split',
      orient: 'vert',
      dir: 'in',
    })

    const zoomOut = patchSlideTransitionXml(BODY, 'zoom', { dir: 'out' })
    expect(zoomOut).toContain('<p:zoom dir="out"/>')
    expect(readSlideTransitionSpecXml(zoomOut)).toMatchObject({ kind: 'zoom', dir: 'out' })
  })

  it('writes the legacy plain bytes when no options are given', () => {
    expect(patchSlideTransitionXml(BODY, 'push')).toContain('<p:push dir="u"/>')
    expect(patchSlideTransitionXml(BODY, 'wipe')).toContain('<p:wipe dir="l"/>')
    expect(patchSlideTransitionXml(BODY, 'split')).toContain('<p:split orient="horz" dir="out"/>')
    expect(patchSlideTransitionXml(BODY, 'zoom')).toContain('<p:zoom/>')
    expect(patchSlideTransitionXml(BODY, 'fade')).not.toContain('AlternateContent')
  })

  it('maps an unknown dir attribute back to no direction', () => {
    const s = '</p:cSld><p:transition><p:push dir="xx"/></p:transition></p:sld>'
    expect(readSlideTransitionSpecXml(s)).toEqual({ kind: 'push', durationMs: null })
  })
})

describe('transition Effect Options: duration serialization', () => {
  it('writes an AlternateContent p14:dur wrapper with a legacy spd Fallback', () => {
    const s = patchSlideTransitionXml(BODY, 'fade', { durationMs: 1500 })
    expect(s).toContain('<mc:AlternateContent')
    expect(s).toContain('Requires="p14"')
    expect(s).toContain('<p:transition spd="slow" p14:dur="1500">')
    expect(s).toContain('<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition>')
    expect(readSlideTransitionSpecXml(s)).toEqual({ kind: 'fade', durationMs: 1500 })
  })

  it('buckets spd by duration (fast ≤ 500ms, med ≤ 1000ms, slow above)', () => {
    expect(patchSlideTransitionXml(BODY, 'fade', { durationMs: 400 })).toContain('spd="fast"')
    expect(patchSlideTransitionXml(BODY, 'fade', { durationMs: 700 })).toContain('spd="med"')
    expect(patchSlideTransitionXml(BODY, 'fade', { durationMs: 2000 })).toContain('spd="slow"')
  })

  it('morph honors an explicit duration (default stays 800ms)', () => {
    const s = patchSlideTransitionXml(BODY, 'morph', { durationMs: 2000 })
    expect(s).toContain('p14:dur="2000"')
    expect(s).toContain('spd="slow"')
    expect(readSlideTransitionSpecXml(s)).toEqual({ kind: 'morph', durationMs: 2000 })
    expect(readSlideTransitionSpecXml(patchSlideTransitionXml(BODY, 'morph')).durationMs).toBe(800)
  })
})

describe('transition Effect Options: round-trip through save/reopen', () => {
  it('round-trips every kind×direction×duration combination the UI offers', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slides = opened.deck.slides
    const combos = allOptionCombos().filter((c) => c.kind !== 'none')
    // slides are reused round-robin; the last write per slide is the one that must survive
    const lastPerSlide = new Map<number, (typeof combos)[number]>()
    // zoom dir="in" is the schema default and serializes as the plain <p:zoom/>
    const readDir = (o: { kind: SlideTransitionKind; dir?: SlideTransitionDir }) =>
      o.kind === 'zoom' && o.dir === 'in' ? undefined : o.dir

    combos.forEach((opts, i) => {
      const at = i % slides.length
      setSlideTransition(slides[at]!, opts.kind, opts)
      lastPerSlide.set(at, opts)
      const back = getSlideTransitionSpec(slides[at]!)
      expect(back.kind).toBe(opts.kind)
      expect(back.dir).toBe(readDir(opts))
      if (opts.kind === 'split' && opts.orient) expect(back.orient).toBe(opts.orient)
      expect(back.durationMs).toBe(opts.durationMs ?? null)
    })

    const reopened = await openPptx(await savePptx(opened))
    for (const [at, opts] of lastPerSlide) {
      const back = getSlideTransitionSpec(reopened.deck.slides[at]!)
      expect(back.kind).toBe(opts.kind)
      expect(back.dir).toBe(readDir(opts))
      if (opts.kind === 'split' && opts.orient) expect(back.orient).toBe(opts.orient)
      expect(back.durationMs).toBe(opts.durationMs ?? null)
    }
  })
})
