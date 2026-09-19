/**
 * UX-1005: the Effects shadow-gallery swatch must bed its preview on light
 * paper in both themes. The preset shadows are authored document data (black
 * ~43%, shadow-effects.ts — CLAUDE.md rule 4: never themed), so on the dark
 * chrome's --hover (#333) every swatch read as the same gray chip and the
 * gallery was unusable. The fix is a theme-static --swatch-bg bed token plus
 * a bed/face structure: outer drop-shadows land on the bed, the inner preset
 * shades the face itself.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const HERE = dirname(fileURLToPath(import.meta.url))
const DOCS_CSS = readFileSync(join(HERE, '../src/renderer/styles.css'), 'utf8')
const TOKENS_CSS = readFileSync(join(HERE, '../../../packages/ui/src/tokens.css'), 'utf8')

/** `--name: value;` (or plain property) declarations of the first rule whose
 * selector matches; comments are stripped first — a comment ending between
 * two semicolons would otherwise glue the surrounding declarations together */
function ruleDecls(css: string, selectorRe: RegExp): Map<string, string> {
  const m = new RegExp(`${selectorRe.source}\\s*\\{([^}]*)\\}`, 'm').exec(css)
  if (!m) throw new Error(`rule not found: ${selectorRe}`)
  const out = new Map<string, string>()
  const body = m[1].replace(/\/\*[\s\S]*?\*\//g, '')
  for (const d of body.matchAll(/(--[\w-]+|\w[\w-]*):\s*([^;]+);/g)) out.set(d[1], d[2].trim())
  return out
}

function channel(hex: string, shift: number): number {
  return (parseInt(hex.slice(1), 16) >> shift) & 255
}

describe('shadow gallery swatch stays legible in dark (UX-1005)', () => {
  it('--swatch-bg is defined light in all three token blocks (theme-static)', () => {
    const blocks = [
      ['light :root', ruleDecls(TOKENS_CSS, /:root(?!\s*:not)/)],
      ["[data-theme='dark']", ruleDecls(TOKENS_CSS, /\[data-theme='dark'\]/)],
      [
        'system-dark fallback',
        ruleDecls(TOKENS_CSS, /:root:not\(\[data-theme='light'\]\):not\(\[data-theme='dark'\]\)/),
      ],
    ] as const
    const values = blocks.map(([name, decls]) => {
      const v = decls.get('--swatch-bg')
      expect(v, `${name} defines --swatch-bg`).toMatch(/^#[0-9a-f]{6}$/i)
      return v!
    })
    // identical in every block: the bed never follows the theme
    expect(new Set(values).size).toBe(1)
    // light enough to contrast Word's ~43% black shadow data
    for (const ch of [16, 8, 0]) expect(channel(values[0]!, ch)).toBeGreaterThan(200)
  })

  it('the swatch bed reads the token, not --hover', () => {
    const bed = ruleDecls(DOCS_CSS, /\.shadow-menu \.shadow-menu-preview(?!\w)/)
    expect(bed.get('background')).toBe('var(--swatch-bg)')
    // the face (not the bed) carries the shadow declarations
    const face = ruleDecls(DOCS_CSS, /\.shadow-menu \.shadow-menu-preview-face/)
    expect(face.get('background')).toBe('var(--swatch-face)')
  })
})
