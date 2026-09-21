/**
 * BUG-1223: session restore used to anchor EVERY secondary window at the
 * primary's bounds — `createShellWindow({ cascadeFrom: primaryBounds() })`
 * inside the restore loop — so restoring three windows opened them exactly
 * on top of each other. The cascade source now walks with the windows
 * actually created (primary first, then each restored window's real bounds),
 * the same discipline as the "Move to New Window" path.
 *
 * Source-wiring test (same style as generated-open-sender-routing.test.ts):
 * index.ts is an Electron main module that cannot be imported into a unit
 * test, so the wiring itself is the contract under regression guard.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const source = readFileSync(join(here, '../src/main/index.ts'), 'utf8')

describe('session restore window cascade (BUG-1223)', () => {
  it('the restore loop cascades from the previously restored window, not the primary', () => {
    // the walking cascade source is initialized once before the loop…
    expect(source).toContain('let cascadeFrom = primaryBounds()')
    // …each restored window is created from it…
    expect(source).toContain('createShellWindow({ cascadeFrom })')
    // …and the source advances to the window just created/restored, so
    // windows 2..N step across the screen instead of stacking
    expect(source).toContain(
      'const bounds = entry.win.getNormalBounds()\n      if (bounds.width > 0 && bounds.height > 0) cascadeFrom = bounds',
    )
  })

  it('no restore-time window creation anchors at the primary inside the loop', () => {
    expect(source).not.toContain('createShellWindow({ cascadeFrom: primaryBounds() })')
    // the menu-driven secondary window ("New Window") still legitimately
    // defaults to the primary when it passes no cascade source
    expect(source).toContain('options.cascadeFrom ?? primaryBounds()')
  })
})
