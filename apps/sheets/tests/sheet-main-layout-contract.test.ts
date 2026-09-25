// PAR-206 fix round: the save-compat banner is a conditional 4th child of the
// .sheet-main grid. Regression under test (PR #202 e2e
// sheets-csv-inplace-save): an extra in-flow child in a 3-row
// `auto minmax(0, 1fr) auto` template shifts the flexible row onto the
// workbook and the canvas collapses below the e2e height gate. The contract:
// the template reserves a dedicated auto row for the banner (collapses to 0
// when absent) and EVERY in-flow child pins its grid-row, so sparse
// auto-placement never decides anything regardless of whether the banner is
// rendered. jsdom cannot compute grid layout, so this pins the stylesheet
// source itself plus the ExcelShell child order.
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const rendererDir = join(__dirname, '../src/renderer')
const css = readFileSync(join(rendererDir, 'styles.css'), 'utf8')
const shell = readFileSync(join(rendererDir, 'ExcelShell.tsx'), 'utf8')

/** Extract the declaration block of a top-level `.selector { ... }` rule. */
function ruleBlock(cssText: string, selector: string): string | null {
  const match = cssText.match(new RegExp(`(?:^|\\n)\\${selector}\\s*\\{([^}]*)\\}`))
  return match ? (match[1] ?? null) : null
}

function declarations(block: string): Map<string, string> {
  const map = new Map<string, string>()
  const stripped = block.replace(/\/\*[\s\S]*?\*\//g, '')
  for (const line of stripped.split(';')) {
    const idx = line.indexOf(':')
    if (idx > 0) map.set(line.slice(0, idx).trim(), line.slice(idx + 1).trim())
  }
  return map
}

/** Children of .sheet-main in document order, as (marker, inFlow) pairs. */
const SHEET_MAIN_CHILDREN: ReadonlyArray<{ marker: string; inFlow: boolean; row: number }> = [
  { marker: 'className="name-box-bar"', inFlow: true, row: 1 },
  { marker: '{saveCompatBanner}', inFlow: true, row: 2 },
  { marker: 'className="workbook-area"', inFlow: true, row: 3 },
  { marker: 'className="status-bar"', inFlow: true, row: 4 },
]

describe('sheet-main layout contract (PAR-206 banner regression)', () => {
  it('reserves a dedicated banner row; the canvas row stays the only flexible one', () => {
    const block = ruleBlock(css, '.sheet-main')
    expect(block, '.sheet-main rule must exist in styles.css').not.toBeNull()
    const template = declarations(block!).get('grid-template-rows')
    expect(template, 'grid-template-rows must be declared').toBeTruthy()
    const tracks = template!
      .replace(/,\s+/g, ',')
      .split(/\s+/)
      .filter((t) => t.length > 0)
    expect(tracks, 'template: banner row / canvas row / status row + name box').toEqual([
      'auto',
      'auto',
      'minmax(0,1fr)',
      'auto',
    ])
    // The canvas row is the single flexible track; the banner row is auto so
    // it collapses to 0 height when the banner is not rendered.
    expect(tracks.filter((t) => t.includes('1fr'))).toEqual(['minmax(0,1fr)'])
  })

  it('pins grid-row of every in-flow child so auto-placement never shifts rows', () => {
    const selectors = ['.name-box-bar', '.save-compat-banner', '.workbook-area', '.status-bar']
    selectors.forEach((selector, i) => {
      const block = ruleBlock(css, selector)
      expect(block, `${selector} rule must exist`).not.toBeNull()
      const pinned = declarations(block!).get('grid-row')
      expect(pinned, `${selector} must pin grid-row (banner must not reflow the grid)`).toBe(
        String(i + 1),
      )
    })
  })

  it('ExcelShell renders the contract children of .sheet-main in pinned order', () => {
    const start = shell.indexOf('<div className="sheet-main">')
    expect(start, 'sheet-main container must exist in ExcelShell').toBeGreaterThanOrEqual(0)
    const end = shell.indexOf('</footer>', start)
    expect(end, 'status-bar footer must close the sheet-main region').toBeGreaterThan(start)
    const region = shell.slice(start, end)
    let cursor = 0
    for (const { marker, row } of SHEET_MAIN_CHILDREN) {
      const at = region.indexOf(marker, cursor)
      expect(at, `${marker} must appear in row-${row} order inside .sheet-main`).toBeGreaterThan(-1)
      cursor = at + marker.length
    }
  })
})
