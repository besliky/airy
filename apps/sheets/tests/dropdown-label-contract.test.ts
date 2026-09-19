import { describe, expect, it } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'

/**
 * UX-1206 dropdown naming contract (sheets twin of the slides rule in
 * ribbon-dropdown-a11y.test.ts): the shared Dropdown (packages/ui) always
 * sets aria-label on its trigger, falling back to the current option's
 * label — which overrides any visible field label it sits next to. Without
 * an explicit `ariaLabel` the control answers to its current VALUE («Sort
 * A to Z», «Sum»), not its field.
 *
 * Mechanical trait: a self-closing `<Dropdown … />` element whose source
 * text carries no `ariaLabel=` attribute.
 *
 * Rule, per .tsx file under src/renderer (recursive):
 *   unlabeled dropdowns === ALLOWANCE[file] ?? 0
 *
 * ALLOWANCE is the acknowledged debt this wave chose not to fix here (the
 * audit's finding was slides-side; the sheets dropdowns below are the same
 * class, registered so the number can only shrink). A new dropdown MUST
 * carry an ariaLabel to stay at zero in unlisted files.
 */
const SRC = join(__dirname, '../src/renderer')

/** dropdowns still named by their current value (no explicit ariaLabel) */
const ALLOWANCE: Record<string, number> = {
  // Sort dialog: per-level order (a/d)
  'ExcelShell.tsx': 1,
  // Insert Function: category filter
  'InsertFunctionDialog.tsx': 1,
  // Pivot builder: axis field + value field + aggregation
  'PivotDialog.tsx': 3,
}

/** every .tsx under the renderer, relative to SRC */
function tsxFiles(dir: string, prefix = ''): string[] {
  const out: string[] = []
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.isDirectory()) out.push(...tsxFiles(join(dir, entry.name), `${prefix}${entry.name}/`))
    else if (entry.name.endsWith('.tsx')) out.push(prefix + entry.name)
  }
  return out
}

/**
 * Opening tags of every <Dropdown …> in the source. The tag ends at the
 * first '>' at JSX-brace depth 0, so attribute values containing '>' (arrow
 * callbacks, comparisons) cannot end the tag early — a plain regex to the
 * first '/>' would silently span two elements when a dropdown carries
 * children instead of being self-closing.
 */
function dropdownTags(src: string): string[] {
  const tags: string[] = []
  let i = src.indexOf('<Dropdown')
  while (i >= 0) {
    if (/[\s/>]/.test(src[i + 9] ?? '>')) {
      let depth = 0
      let j = i + 1
      while (j < src.length) {
        const c = src[j]
        if (c === '{') depth++
        else if (c === '}') depth--
        else if (c === '>' && depth === 0) break
        j++
      }
      tags.push(src.slice(i, j + 1))
    }
    i = src.indexOf('<Dropdown', i + 1)
  }
  return tags
}

/** count of <Dropdown> elements without an explicit ariaLabel */
function unlabeledCount(src: string): number {
  return dropdownTags(src).filter((tag) => !tag.includes('ariaLabel=')).length
}

describe('dropdowns name themselves by field, not current value (UX-1206)', () => {
  const files = tsxFiles(SRC)

  it('finds the renderer tree (scanner sanity)', () => {
    expect(files.length).toBeGreaterThan(20)
  })

  it('finds dropdowns to check (scanner sanity, >= 25)', () => {
    const total = files.reduce(
      (n, file) => n + dropdownTags(readFileSync(join(SRC, file), 'utf8')).length,
      0,
    )
    expect(total).toBeGreaterThanOrEqual(25)
  })

  it.each(files)('%s keeps its unlabeled dropdown count at the allowance', (file) => {
    const src = readFileSync(join(SRC, file), 'utf8')
    // equality (not <=) keeps ALLOWANCE self-tightening: fix a dropdown and
    // its allowance must shrink in the same commit
    expect(unlabeledCount(src)).toBe(ALLOWANCE[file] ?? 0)
  })

  it('lists only files that still render unlabeled dropdowns', () => {
    for (const [file, allowance] of Object.entries(ALLOWANCE)) {
      const src = readFileSync(join(SRC, file), 'utf8')
      expect(unlabeledCount(src), `${file} entry is dead`).toBeGreaterThan(0)
      expect(allowance, `${file} allowance must be positive`).toBeGreaterThan(0)
    }
  })
})
