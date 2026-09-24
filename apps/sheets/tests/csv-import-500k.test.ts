/**
 * PERF-1663 scale + correctness guard for the CSV import path: a 500k-row
 * synthetic file must survive the shipped conversion (parseCsv -> worksheet
 * XML -> zip container) with every row accounted for and the boundary rows
 * intact, and — when the sidecar binary is available — open in the engine
 * with head/tail values round-tripping. Kept slim on purpose: a row counter
 * plus first/last-row probes, not a 4M-cell sweep. The corpus is generated
 * programmatically (no binary fixtures) and covers the edge shapes: quoted
 * commas, escaped quotes, cyrillic, emoji, empty trailing cells.
 */
import { existsSync } from 'node:fs'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { z } from 'zod'

import { csvToXlsxBuffer } from '../src/gateway/csv-import'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { workbookRangeResultSchema } from '../src/shared/desktop-api'

const DATA_ROWS = 500_000
const TOTAL_ROWS = DATA_ROWS + 1 // header

const WORDS = ['alpha', 'beta', 'gamma', 'delta', 'epsilon']

/** same shape as the PERF-1663 bench corpus: mixed types, edge-heavy notes */
function buildCsv(): string {
  const parts: string[] = ['id,name,qty,note,tags,price,city,extra\n']
  let buffer = ''
  for (let i = 0; i < DATA_ROWS; i += 1) {
    const first = WORDS[i % WORDS.length]
    const second = WORDS[(i * 7) % WORDS.length]
    let note = `${first} ${second}`
    if (i % 17 === 0) note = `"${first}, ${second}"` // quoted, embedded delimiter
    if (i % 97 === 0) note = `"say ""hi"" ${first}"` // escaped quotes
    if (i % 211 === 0) note = `"строка «ёлки» ${first}"` // cyrillic + guillemets
    if (i % 409 === 0) note = `emoji 😀 ${first}` // astral chars (surrogate pair)
    const extra = i % 5 === 0 ? '' : first // empty trailing cell
    const row =
      `${i + 1},${first}-${second},${i % 1000},${note},${second},${(i % 997) + 0.5},` +
      `City${i % 50},${extra}\n`
    buffer += row
    if (buffer.length > 1 << 22) {
      parts.push(buffer)
      buffer = ''
    }
  }
  parts.push(buffer)
  return parts.join('')
}

describe('csvToXlsxBuffer at 500k rows (PERF-1663)', () => {
  it(
    'converts with the full row count and intact boundary rows',
    { timeout: 120_000 },
    async () => {
      const zip = await JSZip.loadAsync(await csvToXlsxBuffer(buildCsv()))
      const sheet = (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''

      // boolean-gated probes: a bare toContain would print the ~100MB sheet on
      // failure; only the verdict (plus the short needle) may hit the log
      const contains = (needle: string): void => {
        expect(sheet.includes(needle), `worksheet missing ${JSON.stringify(needle)}`).toBe(true)
      }

      // dimension carries the counter: header + DATA_ROWS
      contains(`<dimension ref="A1:H${TOTAL_ROWS}"/>`)
      // first data row (i=0 hits every note edge): numeric id, emoji note intact
      contains('<row r="2"><c r="A2"><v>1</v></c>')
      contains('<t xml:space="preserve">emoji 😀 alpha</t>')
      // edge shapes land where the generator puts them (i % period === 0 -> row i + 2;
      // first/second words follow WORDS[i % 5] / WORDS[(i * 7) % 5])
      contains('<t xml:space="preserve">gamma, epsilon</t>') // i=17: quoted delimiter
      contains('<t xml:space="preserve">say &quot;hi&quot; gamma</t>') // i=97: escaped quotes
      contains('<t xml:space="preserve">строка «ёлки» beta</t>') // i=211: cyrillic
      // empty trailing cells stay absent: row 7 (i=5, i%5===0) must end at G7,
      // no H7 — compared exactly against the extracted row (a bare `H7` needle
      // would false-positive on H700)
      const row7 = /<row r="7">.*?<\/row>/.exec(sheet)?.[0]
      expect(row7).toBe(
        '<row r="7">' +
          '<c r="A7"><v>6</v></c>' +
          '<c r="B7" t="inlineStr"><is><t xml:space="preserve">alpha-alpha</t></is></c>' +
          '<c r="C7"><v>5</v></c>' +
          '<c r="D7" t="inlineStr"><is><t xml:space="preserve">alpha alpha</t></is></c>' +
          '<c r="E7" t="inlineStr"><is><t xml:space="preserve">alpha</t></is></c>' +
          '<c r="F7"><v>5.5</v></c>' +
          '<c r="G7" t="inlineStr"><is><t xml:space="preserve">City5</t></is></c>' +
          '</row>',
      )
      // last data row: numeric id matches the row counter
      contains(`<row r="${TOTAL_ROWS}"><c r="A${TOTAL_ROWS}"><v>${DATA_ROWS}</v></c>`)
    },
  )

  it(
    'opens in the sidecar with head and tail values round-tripping',
    { timeout: 120_000, skip: !sidecarBinaryPath() || !existsSync(sidecarBinaryPath()) },
    async () => {
      const directory = await mkdtemp(join(tmpdir(), 'csv-500k-'))
      const path = join(directory, 'csv_500k.xlsx')
      await writeFile(path, await csvToXlsxBuffer(buildCsv()))
      const client = new XlsxSidecarClient(sidecarBinaryPath())
      const openedResultSchema = z.object({
        sessionId: z.string().uuid(),
        sheets: z.array(
          z.object({ id: z.string(), rowCount: z.number(), columnCount: z.number() }),
        ),
      })
      let sessionId: string | null = null
      try {
        const opened = openedResultSchema.parse(await client.open(path))
        sessionId = opened.sessionId
        const sheet = opened.sheets[0]
        if (!sheet) throw new Error('sidecar returned no sheets')
        expect(sheet).toMatchObject({ rowCount: TOTAL_ROWS, columnCount: 8 })
        const sheetId = sheet.id

        // tail: far rows index in the background — poll until complete
        let tailCells: { row: number; column: number; value: unknown }[] = []
        for (let attempt = 0; attempt < 600 && tailCells.length === 0; attempt += 1) {
          const tail = workbookRangeResultSchema.parse(
            await client.readRange({
              sessionId,
              sheetId,
              range: {
                startRow: TOTAL_ROWS - 1,
                endRow: TOTAL_ROWS - 1,
                startColumn: 0,
                endColumn: 1,
              },
            }),
          )
          if (tail.indexingComplete && tail.cells.length > 0) tailCells = tail.cells
          else await new Promise((resolve) => setTimeout(resolve, 50))
        }
        const head = workbookRangeResultSchema.parse(
          await client.readRange({
            sessionId,
            sheetId,
            range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 },
          }),
        )
        const cellAt = (cells: typeof tailCells, column: number): unknown =>
          cells.find((cell) => cell.column === column)?.value
        expect(cellAt(head.cells, 0)).toBe(1) // first data row id, typed numeric
        expect(cellAt(tailCells, 0)).toBe(DATA_ROWS) // last data row id
      } finally {
        if (sessionId) await client.close(sessionId)
        client.stop()
        await rm(directory, { recursive: true, force: true })
      }
    },
  )
})

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL('../native/xlsx-engine/target/release/' + executable, import.meta.url),
  )
}
