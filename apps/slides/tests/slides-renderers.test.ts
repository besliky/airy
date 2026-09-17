import { beforeEach, describe, expect, it } from 'vitest'

import {
  isSlidesRenderer,
  resetSlidesRenderers,
  trackSlidesRenderer,
  untrackSlidesRenderer,
} from '../src/main/slides-renderers'

beforeEach(() => {
  resetSlidesRenderers()
})

describe('slides renderer registry', () => {
  it('tracks created slides webContents and forgets destroyed ones', () => {
    trackSlidesRenderer(3)
    trackSlidesRenderer(7)
    expect(isSlidesRenderer(3)).toBe(true)
    expect(isSlidesRenderer(7)).toBe(true)
    untrackSlidesRenderer(3)
    expect(isSlidesRenderer(3)).toBe(false)
    expect(isSlidesRenderer(7)).toBe(true)
  })

  it('rejects unknown senders and malformed ids', () => {
    expect(isSlidesRenderer(99)).toBe(false)
    expect(isSlidesRenderer(undefined)).toBe(false)
  })

  it('re-tracking an id is harmless', () => {
    trackSlidesRenderer(5)
    trackSlidesRenderer(5)
    untrackSlidesRenderer(5)
    expect(isSlidesRenderer(5)).toBe(false)
  })

  it('a foreign editor renderer never becomes a slides renderer by asking', () => {
    // the slides:recent and display-media membership check: a webContents
    // from another app in the shared process (fake ids here) is refused
    // until it is tracked as a slides renderer — which only happens for
    // webContents this module's main created
    const sheetsTabId = 12
    const slidesViewId = 40
    trackSlidesRenderer(slidesViewId)
    expect(isSlidesRenderer(sheetsTabId)).toBe(false)
    expect(isSlidesRenderer(slidesViewId)).toBe(true)
  })
})
