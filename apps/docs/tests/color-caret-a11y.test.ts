import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { en } from '../src/renderer/i18n/ribbon/en'

/**
 * UX-1611 source contract: the two ribbon color carets (`rb-caret
 * rb-color-caret` — the palette-opening half of the highlight / font color
 * split buttons) are icon-only buttons and must carry an accessible name,
 * otherwise screen readers announce a bare «button». Checked at the source
 * level like split-button-a11y.test.ts: mounting the whole Ribbon tab just
 * for two static attributes is not worth the fixture.
 */

const SRC = join(__dirname, '../src/renderer/components/Ribbon.tsx')

/**
 * Opening tag of every button styled `rb-color-caret`. The tag ends at the
 * first '>' at JSX-brace depth 0, so the `=>` arrows inside the className
 * template and the onClick handler cannot end the tag early.
 */
function colorCaretTags(src: string): string[] {
  const tags: string[] = []
  let i = src.indexOf('rb-color-caret')
  while (i >= 0) {
    const start = src.lastIndexOf('<button', i)
    if (start < 0) throw new Error('color caret without an opening <button')
    let depth = 0
    let j = i
    while (j < src.length) {
      const c = src[j]
      if (c === '{') depth++
      else if (c === '}') depth--
      else if (c === '>' && depth === 0) break
      j++
    }
    if (j >= src.length) throw new Error('unterminated color caret tag')
    tags.push(src.slice(start, j + 1))
    i = src.indexOf('rb-color-caret', j)
  }
  return tags
}

/** the i18n key referenced by the tag's `aria-label={t('...')}` */
const ariaKey = (tag: string): string | null =>
  tag.match(/aria-label=\{t\('([A-Za-z]+)'\)\}/)?.[1] ?? null

describe('ribbon color carets have accessible names (UX-1611)', () => {
  const tags = colorCaretTags(readFileSync(SRC, 'utf8'))

  it('finds the two color caret buttons (scanner sanity)', () => {
    expect(tags).toHaveLength(2)
  })

  it.each(tags.map((tag, i) => [i, tag] as const))(
    'caret #%i names itself through an i18n aria-label',
    (_i, tag) => {
      const key = ariaKey(tag)
      expect(key, `no aria-label={t('…')} in: ${tag}`).not.toBeNull()
      // the referenced key must resolve in the reference dictionary
      expect(en, `key ${key} is missing from the en dictionary`).toHaveProperty(key!)
    },
  )

  it('labels the two carets with distinct keys (highlight vs font color)', () => {
    const keys = tags.map((tag) => ariaKey(tag))
    expect(new Set(keys).size).toBe(2)
  })
})
