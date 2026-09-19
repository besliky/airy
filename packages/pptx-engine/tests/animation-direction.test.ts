import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  ANIM_DIRECTIONS,
  ANIM_DEFAULT_DIR,
  buildTimingXml,
  elementSpid,
  getSlideAnimations,
  openPptx,
  patchSlideTimingIncrementalXml,
  readSlideTimingXml,
  savePptx,
  setSlideAnimations,
  type AnimDirection,
  type AnimEffectKind,
  type SlideAnimation,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const anim = (
  o: Partial<SlideAnimation> & Pick<SlideAnimation, 'spid' | 'effect'>,
): SlideAnimation => ({ trigger: 'onClick', durationMs: 500, delayMs: 0, ...o })

describe('animation direction: serialization', () => {
  it('lists direction options per effect (empty = effect has none)', () => {
    expect(ANIM_DIRECTIONS.flyIn).toContain('fromLeft')
    expect(ANIM_DIRECTIONS.wipe).toHaveLength(4)
    expect(ANIM_DIRECTIONS.appear).toHaveLength(0)
    expect(ANIM_DEFAULT_DIR.flyIn).toBe('fromBottom')
    expect(ANIM_DEFAULT_DIR.wipeOut).toBe('fromTop')
  })

  it('writes fly-in offsets for each direction (corners move both axes)', () => {
    const xml = buildTimingXml([
      anim({ spid: 4, effect: 'flyIn', direction: 'fromLeft' }),
      anim({ spid: 5, effect: 'flyIn', direction: 'fromTop' }),
      anim({ spid: 6, effect: 'flyIn', direction: 'fromTopRight' }),
      anim({ spid: 7, effect: 'flyIn' }),
      anim({ spid: 8, effect: 'flyOut', direction: 'fromRight' }),
    ])
    // fromLeft: x starts at the 0-edge, y stays identity (the default's shape)
    expect(xml).toContain('val="0-#ppt_w/2"')
    expect(xml).toContain('val="0-#ppt_h/2"')
    // fromTopRight: x at the 1-edge, y at the 0-edge
    expect(xml).toContain('val="1+#ppt_w/2"')
    // default (no direction) keeps the legacy from-bottom shape: identity x + bottom y
    expect(xml).toContain('val="1+#ppt_h/2"')
    // flyOut toward the right edge: y identity, x to the 1-edge
    expect(xml).toContain('<p:strVal val="1+#ppt_w/2"/></p:val></p:tav>')
  })

  it('writes wipe travel per direction and split axis variants', () => {
    const xml = buildTimingXml([
      anim({ spid: 4, effect: 'wipe', direction: 'fromLeft' }),
      anim({ spid: 4, effect: 'wipe', direction: 'fromTop' }),
      anim({ spid: 4, effect: 'splitIn', direction: 'vertOut' }),
    ])
    expect(xml).toContain('filter="wipe(right)"') // fromLeft travels right
    expect(xml).toContain('filter="wipe(down)"') // fromTop travels down
    expect(xml).toContain('filter="split(outVertical)"')
  })

  it('writes zoom out variants (entrance settles from 200%, exit grows to 300%)', () => {
    const xml = buildTimingXml([
      anim({ spid: 4, effect: 'zoom', direction: 'out' }),
      anim({ spid: 4, effect: 'zoomOut', direction: 'out' }),
    ])
    expect(xml).toContain('<p:from x="200000" y="200000"/><p:to x="100000" y="100000"/>')
    expect(xml).toContain('<p:to x="300000" y="300000"/>')
  })

  it('writes spin counter-clockwise as a negative rotation', () => {
    expect(buildTimingXml([anim({ spid: 4, effect: 'spin', direction: 'ccw' })])).toContain(
      '<p:animRot by="-21600000">',
    )
  })

  it('encodes directions in presetSubtype (PowerPoint directional bitmask)', () => {
    const subOf = (a: SlideAnimation) => /presetSubtype="(\d+)"/.exec(buildTimingXml([a]))?.[1]
    expect(subOf(anim({ spid: 4, effect: 'flyIn', direction: 'fromTop' }))).toBe('1')
    expect(subOf(anim({ spid: 4, effect: 'flyIn', direction: 'fromRight' }))).toBe('2')
    expect(subOf(anim({ spid: 4, effect: 'flyIn', direction: 'fromBottom' }))).toBe('4')
    expect(subOf(anim({ spid: 4, effect: 'flyIn', direction: 'fromLeft' }))).toBe('8')
    expect(subOf(anim({ spid: 4, effect: 'flyIn', direction: 'fromTopRight' }))).toBe('3')
    expect(subOf(anim({ spid: 4, effect: 'flyIn', direction: 'fromBottomLeft' }))).toBe('12')
    // wipe family defaults: wipe = fromBottom (4), wipeDown/wipeOut = fromTop (1)
    expect(subOf(anim({ spid: 4, effect: 'wipe' }))).toBe('4')
    expect(subOf(anim({ spid: 4, effect: 'wipeDown' }))).toBe('1')
    expect(subOf(anim({ spid: 4, effect: 'wipeOut' }))).toBe('1')
    // no direction → the effect's own preset default (split/zoom keep theirs)
    expect(subOf(anim({ spid: 4, effect: 'flyIn' }))).toBe('4')
    expect(subOf(anim({ spid: 4, effect: 'zoom' }))).toBe('16')
  })
})

describe('animation direction: read-back', () => {
  it('round-trips every effect×direction combination the UI offers', () => {
    // Canonicalization on read-back: the plain form IS the default direction, and
    // in the wipe family the kind itself encodes the default travel (down-travel
    // reads as wipeDown, everything else as wipe + direction)
    const normalized = (effect: AnimEffectKind, direction: AnimDirection): SlideAnimation => {
      if (effect === 'wipe' || effect === 'wipeDown') {
        if (direction === 'fromTop') return anim({ spid: 9, effect: 'wipeDown' })
        if (direction === 'fromBottom') return anim({ spid: 9, effect: 'wipe' })
        return anim({ spid: 9, effect: 'wipe', direction })
      }
      if (direction === ANIM_DEFAULT_DIR[effect]) return anim({ spid: 9, effect })
      return anim({ spid: 9, effect, direction })
    }
    for (const effect of Object.keys(ANIM_DIRECTIONS) as AnimEffectKind[]) {
      for (const direction of ANIM_DIRECTIONS[effect]) {
        const a = anim({ spid: 9, effect, direction })
        const back = readSlideTimingXml(`</p:cSld>${buildTimingXml([a])}</p:sld>`)
        expect(back).toEqual([normalized(effect, direction)])
      }
    }
  })

  it('keeps directionless effects directionless (defaults read back as undefined)', () => {
    const anims = [
      anim({ spid: 4, effect: 'flyIn' }),
      anim({ spid: 4, effect: 'wipe' }),
      anim({ spid: 4, effect: 'wipeDown' }),
      anim({ spid: 4, effect: 'wipeOut' }),
      anim({ spid: 4, effect: 'splitIn' }),
      anim({ spid: 4, effect: 'zoom' }),
      anim({ spid: 4, effect: 'spin' }),
    ]
    expect(readSlideTimingXml(`</p:cSld>${buildTimingXml(anims)}</p:sld>`)).toEqual(anims)
  })

  it('derives directions from foreign presetSubtypes (PowerPoint-written files)', () => {
    // A subtype-only file (behaviors replaced by PowerPoint defaults, no filter)
    // still maps through the directional bitmask: 8 = From Left
    const plain = buildTimingXml([anim({ spid: 4, effect: 'flyIn' })]).replace(
      'presetSubtype="4"',
      'presetSubtype="8"',
    )
    expect(readSlideTimingXml(`</p:cSld>${plain}</p:sld>`)[0]!.direction).toBe('fromLeft')
  })

  it('reads legacy Airy wipes (sub 1 + filter up) as the plain bottom wipe', () => {
    // Old builds wrote wipe with presetSubtype 1 but a wipe(up) filter — the
    // filter is ground truth, so the model canonicalizes to 'wipe'
    const xml = buildTimingXml([anim({ spid: 4, effect: 'wipe' })]).replace(
      'presetSubtype="4"',
      'presetSubtype="1"',
    )
    const back = readSlideTimingXml(`</p:cSld>${xml}</p:sld>`)
    expect(back[0]!.effect).toBe('wipe')
    expect(back[0]!.direction).toBeUndefined()
  })

  it('round-trips directions through save/reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const spids = slide.elements.map((el) => elementSpid(el)!)
    const anims: SlideAnimation[] = [
      anim({ spid: spids[0]!, effect: 'flyIn', direction: 'fromTopLeft' }),
      anim({ spid: spids[1]!, effect: 'wipe', direction: 'fromLeft', trigger: 'withPrev' }),
      anim({ spid: spids[0]!, effect: 'splitIn', direction: 'vertIn', trigger: 'afterPrev' }),
      anim({ spid: spids[1]!, effect: 'spin', direction: 'ccw' }),
      anim({ spid: spids[0]!, effect: 'zoomOut', direction: 'out' }),
      anim({ spid: spids[1]!, effect: 'flyOut', direction: 'fromBottomRight' }),
    ]
    setSlideAnimations(slide, anims)
    expect(getSlideAnimations(slide)).toEqual(anims)
    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideAnimations(reopened.deck.slides[0]!)).toEqual(anims)
  })

  it('incremental patch handles a tail direction-only edit', () => {
    const old = [anim({ spid: 4, effect: 'flyIn' })]
    const next = [anim({ spid: 4, effect: 'flyIn', direction: 'fromLeft' })]
    const xml = `</p:cSld>${buildTimingXml(old)}</p:sld>`
    const patched = patchSlideTimingIncrementalXml(xml, next)
    expect(patched).not.toBeNull()
    expect(readSlideTimingXml(patched!)).toEqual(next)
  })
})
