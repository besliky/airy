import { beforeEach, describe, expect, it } from 'vitest'

import {
  isDocsRenderer,
  resetDocsRenderers,
  trackDocsRenderer,
  untrackDocsRenderer,
} from '../src/main/docs-renderers'

beforeEach(() => {
  resetDocsRenderers()
})

describe('docs renderer registry', () => {
  it('tracks created docs webContents and forgets destroyed ones', () => {
    trackDocsRenderer(3)
    trackDocsRenderer(7)
    expect(isDocsRenderer(3)).toBe(true)
    expect(isDocsRenderer(7)).toBe(true)
    untrackDocsRenderer(3)
    expect(isDocsRenderer(3)).toBe(false)
    expect(isDocsRenderer(7)).toBe(true)
  })

  it('rejects unknown senders and malformed ids', () => {
    expect(isDocsRenderer(99)).toBe(false)
    expect(isDocsRenderer(undefined)).toBe(false)
  })

  it('re-tracking an id is harmless', () => {
    trackDocsRenderer(5)
    trackDocsRenderer(5)
    untrackDocsRenderer(5)
    expect(isDocsRenderer(5)).toBe(false)
  })

  it('a foreign editor renderer never becomes a docs renderer by asking', () => {
    // the docs:recent / win:new membership check: a webContents from
    // another app in the shared process (fake ids here) is refused until
    // tracked — which only happens for webContents docs' main created
    const sheetsTabId = 12
    const docsViewId = 40
    trackDocsRenderer(docsViewId)
    expect(isDocsRenderer(sheetsTabId)).toBe(false)
    expect(isDocsRenderer(docsViewId)).toBe(true)
  })
})
