/**
 * BUG-1542: Word's autoSpaceDE/DN CJK-Latin gap is ~1/4em — 3pt at the 12pt
 * reference size (Word probe 2026-09-01, kept in autospace-doc-off.test.ts;
 * audit 2026-09-21 measured the shipped product at exactly 1/8em, half the
 * target). The gap is delivered by .doc-autospace-pad's start margin because
 * Chromium's native text-autospace never applies across the pad span — the
 * element boundary ends the shaped run — so the pad must carry the FULL Word
 * gap, not half of it. jsdom cannot lay text out, so this pins the stylesheet
 * contract: the custom property, the fallback in the pad rule, and the Word
 * arithmetic of the chosen value.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const css = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
  'utf8',
)

describe('CJK-Latin autospace pad width (BUG-1542)', () => {
  it('the pad carries the full Word gap: .doc-page sets --doc-autospace-pad to 0.25em', () => {
    const defs = css.match(/--doc-autospace-pad:\s*0\.25em/g) ?? []
    expect(defs.length).toBe(1)
  })

  it('the pad rule reads the variable with the same full-gap fallback', () => {
    expect(css).toMatch(
      /\.doc-autospace-pad\s*\{[^}]*margin-inline-start:\s*var\(--doc-autospace-pad,\s*0\.25em\)/,
    )
  })

  it('no stale half-gap value remains anywhere in the stylesheet', () => {
    expect(css).not.toContain('0.125em')
  })

  it('0.25em at the 12pt reference size is the audited 3pt Word target', () => {
    // Word probe: gap = 3pt at 12pt; 1em = the run's font size, so the em
    // fraction must reproduce that ratio at every size
    expect(0.25 * 12).toBe(3)
  })
})
