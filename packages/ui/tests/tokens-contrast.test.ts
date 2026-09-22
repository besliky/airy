import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

import { describe, expect, it } from 'vitest'

const uiRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const tokensPath = join(uiRoot, 'src', 'tokens.css')
const css = readFileSync(tokensPath, 'utf8')

/** Parse `--token: value;` declarations out of a single CSS block body.
 * The selector regex must anchor the rule's opening brace so a selector
 * mentioned inside a comment (the file header documents these) never matches. */
function blockDeclarations(text: string, selector: RegExp): Map<string, string> {
  const match = text.match(selector)
  if (!match) throw new Error(`selector not found: ${String(selector)}`)
  const open = text.indexOf('{', match.index! + match[0].length - 1)
  let depth = 1
  let end = open
  while (depth > 0 && end < text.length) {
    end += 1
    const ch = text[end]
    if (ch === '{') depth += 1
    if (ch === '}') depth -= 1
  }
  const body = text.slice(open + 1, end)
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

const light = blockDeclarations(css, /:root\s*\{/)
const dark = blockDeclarations(css, /\[data-theme='dark'\]\s*\{/)
const systemDark = blockDeclarations(
  css,
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

/* ── App ring tokens (UX-831) ──
 * Focus/selection rings are alpha colors; WCAG 2.x contrast applies to what
 * they composite into over the surface they are painted on. Every ring token
 * of the pdf/docs/markdown apps must reach the non-text minimum (3:1) on the
 * worst-case surface of its theme:
 *   light theme — the lightest chrome surface, --surface #ffffff;
 *   dark theme  — --surface #1e1e1e and the lighter --surface-subtle /
 *                 --color-bg-subtle #2a2a2a (hover+focus states).
 * Paper overlays (pdf text-edit/link marks) are theme-static and always sit
 * on the white page. Systems-dark blocks must mirror the dark values. */

type RingSurfaces = { light: string[]; dark: string[] }

const CHROME: RingSurfaces = { light: ['#ffffff'], dark: ['#1e1e1e', '#2a2a2a'] }
const PAPER: RingSurfaces = { light: ['#ffffff'], dark: ['#ffffff'] }

const APP_RING_TOKENS: Array<{
  app: string
  token: string
  surfaces: RingSurfaces
}> = [
  { app: 'pdf', token: '--pdf-focus-ring', surfaces: CHROME },
  { app: 'pdf', token: '--pdf-swatch-ring', surfaces: CHROME },
  { app: 'pdf', token: '--pdf-input-ring', surfaces: CHROME },
  { app: 'pdf', token: '--pdf-error-ring', surfaces: CHROME },
  { app: 'pdf', token: '--pdf-link-hover-ring', surfaces: PAPER },
  { app: 'pdf', token: '--pdf-textedit-ring', surfaces: PAPER },
  { app: 'pdf', token: '--pdf-textblock-ring', surfaces: PAPER },
  { app: 'pdf', token: '--pdf-textblock-ring-hover', surfaces: PAPER },
  { app: 'docs', token: '--docs-focus-ring', surfaces: CHROME },
  { app: 'docs', token: '--docs-focus-ring-blue', surfaces: CHROME },
  { app: 'docs', token: '--docs-brand-ring', surfaces: CHROME },
  { app: 'docs', token: '--docs-input-ring', surfaces: CHROME },
  { app: 'docs', token: '--docs-error-ring', surfaces: CHROME },
  { app: 'markdown', token: '--md-focus-ring', surfaces: CHROME },
  { app: 'markdown', token: '--md-focus-ring-blue', surfaces: CHROME },
]

const appsRoot = join(uiRoot, '..', '..')

type Rgba = { r: number; g: number; b: number; a: number }

function parseRgba(value: string): Rgba {
  const m = value.match(/^rgba?\(\s*([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*\/\s*([\d.]+)%\s*\)$/)
  if (!m) throw new Error(`ring token must be rgb(R G B / A%): ${value}`)
  return { r: Number(m[1]), g: Number(m[2]), b: Number(m[3]), a: Number(m[4]) / 100 }
}

function compositeOver(fg: Rgba, bg: [number, number, number]): [number, number, number] {
  return [
    fg.a * fg.r + (1 - fg.a) * bg[0],
    fg.a * fg.g + (1 - fg.a) * bg[1],
    fg.a * fg.b + (1 - fg.a) * bg[2],
  ]
}

function contrastOver(fg: Rgba, bgHex: string): number {
  const bg = hexToRgb(bgHex)
  const l1 = relativeLuminance(
    `#${compositeOver(fg, bg)
      .map((v) => Math.round(v).toString(16).padStart(2, '0'))
      .join('')}`,
  )
  const l2 = relativeLuminance(bgHex)
  return (Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05)
}

describe('app ring tokens composite to >= 3:1 on their surfaces (UX-831)', () => {
  for (const { app, token, surfaces } of APP_RING_TOKENS) {
    const appCss = readFileSync(
      join(appsRoot, 'apps', app, 'src', 'renderer', 'styles.css'),
      'utf8',
    )
    const appLight = blockDeclarations(appCss, /:root\s*\{/)
    const appDark = blockDeclarations(appCss, /\[data-theme='dark'\]\s*\{/)
    const appSystemDark = blockDeclarations(
      appCss,
      /:root:not\(\[data-theme='light'\]\):not\(\[data-theme='dark'\]\)\s*\{/,
    )
    // Theme-static (paper) rings are declared once in :root and apply to both themes.
    const themeValue = (theme: 'light' | 'dark'): string => {
      const value = (theme === 'dark' ? appDark : appLight).get(token)
      if (value === undefined && theme === 'dark') return appLight.get(token)!
      if (value === undefined) throw new Error(`${token} not found in ${app} styles.css`)
      return value
    }
    const mirroredInSystemDark = appDark.has(token)
      ? appSystemDark.get(token) === appDark.get(token)
      : true

    it(`${token} (${app}) meets the non-text minimum in both themes`, () => {
      for (const theme of ['light', 'dark'] as const) {
        const rgba = parseRgba(themeValue(theme))
        for (const surface of surfaces[theme]) {
          expect(
            contrastOver(rgba, surface),
            `${token} [${theme}] ${themeValue(theme)} on ${surface}`,
          ).toBeGreaterThanOrEqual(3)
        }
      }
      expect(mirroredInSystemDark, `${token} system-dark must mirror dark`).toBe(true)
    })
  }
})
