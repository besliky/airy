import { describe, expect, it } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'

/**
 * TEST-1103: unit literals must not be appended after a translated string in
 * JSX (the UX-908/UX-1009 recurrence class). A literal ` (pt)` / ` (cm)`
 * suffix after a t() expression does not translate — Japanese writes the
 * point unit as a word ("ポイント"), and the suffix stays mispositioned in
 * RTL — so the unit belongs inside the dictionary string (ribbonPt,
 * ribbonLnDistance, layoutColWidth (cm), ...). Both incidents were exactly
 * this shape: `{t('layoutColSpacing')} (cm)` raw JSX text (UX-1009) and the
 * LineNumbersDialog distance label `{t('ribbonLnDistance')} (pt)` (UX-908,
 * previously pinned only inside line-numbers-dialog.test.ts — now merged
 * into this tree-wide contract).
 *
 * Mechanical traits — the shapes a suffix after an expression can take in
 * .tsx source (comments are stripped before matching, so prose like
 * "column (cm):" in a comment does not count):
 *   - `{...} (pt)` / `{...} (cm)` — raw JSX text after any expression close
 *     (covers `{t('k')} (cm)` and `{cond ? t('a') : t('b')} (cm)`);
 *   - `t(...) + ' (pt)'` — concatenation onto a translated string;
 *   - `${...} (pt)` — the same suffix inside a template literal.
 *
 * The scanner is a ratchet at zero: there is no LEGACY baseline because both
 * known instances are fixed and a new occurrence is always a bug.
 */
const SRC = join(__dirname, '../src/renderer/components')

const RULES: ReadonlyArray<readonly [name: string, re: RegExp]> = [
  ['raw JSX suffix after an expression', /\}\s*\((?:pt|cm)\)/],
  ['concatenated suffix after t()', /\bt\([^)]*\)\s*\+\s*['"`]\s*\((?:pt|cm)\)/],
  ['suffix after a template interpolation', /\$\{[^}]*\}\s*\((?:pt|cm)\)/],
]

function stripComments(src: string): string {
  return src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/(^|[^:])\/\/.*$/gm, '$1')
}

describe('no unit literals after translated strings in tsx (TEST-1103)', () => {
  const files = readdirSync(SRC).filter((f) => f.endsWith('.tsx'))

  it('finds the component directory (scanner sanity)', () => {
    expect(files.length).toBeGreaterThan(10)
  })

  it.each(files)('%s carries no appended unit literal', (file) => {
    const src = stripComments(readFileSync(join(SRC, file), 'utf8'))
    for (const [name, re] of RULES) {
      const hit = re.exec(src)
      expect(hit, `${file}: ${name} — ${JSON.stringify(hit?.[0] ?? '')}`).toBeNull()
    }
  })
})
