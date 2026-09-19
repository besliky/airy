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
 * UX-11s extended the ratchet from Ribbon.tsx to the two legacy tab files
 * (RibbonHomeTab / RibbonInsertTab, vintage ≤2026-09-10) that the first pass
 * deliberately skipped. Two shapes live here:
 * - plain menu buttons — the `<button>` itself carries the handler;
 * - split buttons (`rb-split`) — the toggle sits on a nested caret-hit
 *   `<span>` zone, so the attributes go on the enclosing focusable
 *   `<button>`, which announces the composite "action + menu" control.
 *
 * FormatPane galleries and ColorWell already declare their state
 * (aria-expanded + haspopup) at the component level; FormatPane's
 * `secHeader` collapses are accordion sections, not popups — neither is a
 * dropdown trigger and neither matches the scanner (no capital between
 * `set` and `Open` / not inside a button tag).
 *
 * Checked at the source level because mounting the whole Ribbon for a
 * static attribute is not worth the fixture.
 */
const SRC_DIR = join(__dirname, '../src/renderer/components')

/** files swept by the contract + the minimum toggles each must still carry */
const FILES: ReadonlyArray<readonly [file: string, min: number]> = [
  ['Ribbon.tsx', 10],
  ['RibbonHomeTab.tsx', 9],
  ['RibbonInsertTab.tsx', 2],
]

/** the opening tag of the <button> whose handler toggles the given setter call */
const triggerTag = (src: string, call: string, at: number): string => {
  // scan back to the tag start; forward from the call's end to its close
  // (the handler's arrow contains '>' so a plain [^>]* span cannot cross it)
  const start = src.lastIndexOf('<button', at)
  const end = src.indexOf('>', at + call.length)
  if (start < 0 || end < 0 || start > at) throw new Error(`malformed trigger: ${call}`)
  return src.slice(start, end + 1)
}

describe.each(FILES)('ribbon dropdown trigger semantics (UX-1103): %s', (file, min) => {
  const src = readFileSync(join(SRC_DIR, file), 'utf8')
  const toggles = [...src.matchAll(/set([A-Z]\w*?)Open\(\(v\) => !v\)/g)]

  it(`finds the inline menu toggles (scanner sanity, >= ${min})`, () => {
    expect(toggles.length).toBeGreaterThanOrEqual(min)
  })

  it.each(toggles.map((m) => [m[1], m[0], m.index] as const))(
    'the %s menu trigger exposes its open state',
    (name, call, at) => {
      const tag = triggerTag(src, call as string, at as number)
      expect(tag).toContain('aria-haspopup="menu"')
      // setTransOptionsOpen toggles transOptionsOpen (first char lowercased)
      expect(tag).toContain(`aria-expanded={${name[0]!.toLowerCase() + name.slice(1)}Open}`)
    },
  )
})
