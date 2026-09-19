import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

/**
 * UX-1103 source contract (slides twin of the docs UX-909/UX-1003
 * split-button gate): the Ribbon's inline dropdown triggers — buttons that
 * toggle a per-panel `set<Name>Open` boolean, unlike docs' shared
 * toggleDropdown key — must expose their open state: `aria-haspopup="menu"`
 * plus an `aria-expanded` bound to the very state the trigger toggles
 * (screen readers announce collapsed / expanded). PR #68's Effect Options
 * shipped without it because the docs contract did not reach slides.
 *
 * Checked at the source level because mounting the whole Ribbon for a
 * static attribute is not worth the fixture.
 */
const SRC = join(__dirname, '../src/renderer/components/Ribbon.tsx')

/** the opening tag of the <button> whose handler toggles the given setter call */
const triggerTag = (src: string, call: string, at: number): string => {
  // scan back to the tag start; forward from the call's end to its close
  // (the handler's arrow contains '>' so a plain [^>]* span cannot cross it)
  const start = src.lastIndexOf('<button', at)
  const end = src.indexOf('>', at + call.length)
  if (start < 0 || end < 0 || start > at) throw new Error(`malformed trigger: ${call}`)
  return src.slice(start, end + 1)
}

describe('ribbon dropdown trigger semantics (UX-1103)', () => {
  const src = readFileSync(SRC, 'utf8')

  it('finds the inline menu toggles (scanner sanity)', () => {
    expect([...src.matchAll(/set([A-Z]\w*?)Open\(\(v\) => !v\)/g)].length).toBeGreaterThanOrEqual(
      10,
    )
  })

  it.each(
    [...src.matchAll(/set([A-Z]\w*?)Open\(\(v\) => !v\)/g)].map((m) => [m[1], m[0], m.index]),
  )('the %s menu trigger exposes its open state', (name, call, at) => {
    const tag = triggerTag(src, call as string, at as number)
    expect(tag).toContain('aria-haspopup="menu"')
    // setTransOptionsOpen toggles transOptionsOpen (first char lowercased)
    expect(tag).toContain(`aria-expanded={${name[0]!.toLowerCase() + name.slice(1)}Open}`)
  })
})
