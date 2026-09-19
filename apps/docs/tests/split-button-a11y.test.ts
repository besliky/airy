import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

/**
 * UX-909 source contract: dropdown split buttons must expose their open
 * state — `aria-haspopup="menu"` plus an `aria-expanded` bound to the very
 * dropdown key the trigger toggles (screen readers announce collapsed /
 * expanded). Checked at the source level because mounting the whole Ribbon
 * tab just for two static attributes is not worth the fixture.
 */

const SRC = join(__dirname, '../src/renderer/components')

/** the opening tag of the <button> whose handler toggles the given key */
const triggerTag = (file: string, key: string): string => {
  const src = readFileSync(join(SRC, file), 'utf8')
  const call = src.indexOf(`toggleDropdown(setDropdown, '${key}')`)
  if (call < 0) throw new Error(`trigger not found: ${file} / ${key}`)
  // scan back to the tag start and forward to its close (the handler's arrow
  // contains '>' so a plain [^>]* span cannot cross it)
  const start = src.lastIndexOf('<button', call)
  const end = src.indexOf('>', call)
  if (start < 0 || end < 0 || start > call) throw new Error(`malformed trigger: ${file} / ${key}`)
  return src.slice(start, end + 1)
}

describe('split-button dropdown semantics (UX-909)', () => {
  it.each([
    ['ribbon-tabs.tsx', 'compare'],
    ['ribbon-layout-tab.tsx', 'linenumbers'],
    // UX-1003: the Effects/border gallery triggers follow the same contract
    ['Ribbon.tsx', 'shapeEffects'],
    ['Ribbon.tsx', 'picBorder'],
    ['Ribbon.tsx', 'picEffects'],
  ] as const)('%s exposes the %s menu state', (file, key) => {
    const tag = triggerTag(file, key)
    expect(tag).toContain('aria-haspopup="menu"')
    expect(tag).toContain(`aria-expanded={dropdown === '${key}'}`)
  })
})
