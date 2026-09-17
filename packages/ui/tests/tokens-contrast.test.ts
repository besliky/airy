import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

import { describe, expect, it } from 'vitest'

const tokensPath = join(dirname(fileURLToPath(import.meta.url)), '..', 'src', 'tokens.css')
const css = readFileSync(tokensPath, 'utf8')

/** Parse `--token: value;` declarations out of a single CSS block body.
 * The selector regex must anchor the rule's opening brace so a selector
 * mentioned inside a comment (the file header documents these) never matches. */
function blockDeclarations(selector: RegExp): Map<string, string> {
  const match = css.match(selector)
  if (!match) throw new Error(`selector not found in tokens.css: ${String(selector)}`)
  const open = css.indexOf('{', match.index! + match[0].length - 1)
  let depth = 1
  let end = open
  while (depth > 0 && end < css.length) {
    end += 1
    const ch = css[end]
    if (ch === '{') depth += 1
    if (ch === '}') depth -= 1
  }
  const body = css.slice(open + 1, end)
  const decls = new Map<string, string>()
  for (const match of body.matchAll(/(--[\w-]+)\s*:\s*([^;]+);/g)) {
    decls.set(match[1], match[2].trim())
  }
  return decls
}

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace('#', '')
  if (!/^[0-9a-fA-F]{6}$/.test(h)) throw new Error(`not a #rrggbb color: ${hex}`)
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]
}

function channel(c: number): number {
  return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4)
}

function relativeLuminance(hex: string): number {
  const [r, g, b] = hexToRgb(hex).map((v) => channel(v / 255))
  return 0.2126 * r + 0.7152 * g + 0.0722 * b
}

/** WCAG 2.x contrast ratio between two #rrggbb colors. */
export function contrastRatio(fg: string, bg: string): number {
  const l1 = Math.max(relativeLuminance(fg), relativeLuminance(bg))
  const l2 = Math.min(relativeLuminance(fg), relativeLuminance(bg))
  return (l1 + 0.05) / (l2 + 0.05)
}

const light = blockDeclarations(/:root\s*\{/)
const dark = blockDeclarations(/\[data-theme='dark'\]\s*\{/)
const systemDark = blockDeclarations(
  /:root:not\(\[data-theme='light'\]\):not\(\[data-theme='dark'\]\)\s*\{/,
)

describe('tokens.css WCAG contrast', () => {
  it('light-theme tertiary text meets AA (4.5:1) on the chrome surface', () => {
    const surface = light.get('--surface')
    expect(surface).toBe('#ffffff')
    for (const token of ['--text-tertiary', '--color-text-tertiary']) {
      const ratio = contrastRatio(light.get(token) as string, surface as string)
      expect(ratio, `${token} ${light.get(token)} on ${surface}`).toBeGreaterThanOrEqual(4.5)
    }
  })

  it('light-theme muted icons meet the non-text minimum (3:1) on the chrome surface', () => {
    const ratio = contrastRatio(
      light.get('--icon-muted') as string,
      light.get('--surface') as string,
    )
    expect(
      ratio,
      `--icon-muted ${light.get('--icon-muted')} on ${light.get('--surface')}`,
    ).toBeGreaterThanOrEqual(3)
  })

  it('dark-theme muted icons meet the non-text minimum (3:1) on the dark surface', () => {
    const ratio = contrastRatio(dark.get('--icon-muted') as string, dark.get('--surface') as string)
    expect(
      ratio,
      `--icon-muted ${dark.get('--icon-muted')} on ${dark.get('--surface')}`,
    ).toBeGreaterThanOrEqual(3)
  })

  it('keeps tertiary lighter than secondary in light theme (visual hierarchy)', () => {
    const tertiary = contrastRatio(light.get('--text-tertiary') as string, '#ffffff')
    const secondary = contrastRatio(light.get('--text-secondary') as string, '#ffffff')
    expect(tertiary).toBeLessThan(secondary)
  })

  it('keeps dark muted icons dimmer than dark tertiary text (visual hierarchy)', () => {
    const surface = dark.get('--surface') as string
    const iconMuted = contrastRatio(dark.get('--icon-muted') as string, surface)
    const tertiary = contrastRatio(dark.get('--text-tertiary') as string, surface)
    expect(iconMuted).toBeLessThan(tertiary)
  })

  it('system-dark fallback matches the dark-theme block for audited tokens', () => {
    for (const token of ['--text-tertiary', '--color-text-tertiary', '--icon-muted']) {
      expect(systemDark.get(token)).toBe(dark.get(token))
    }
  })
})
