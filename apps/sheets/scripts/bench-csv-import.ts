// One-off benchmark for the CSV import pipeline (PERF-1663, not part of gate).
// Times the same phases the main process runs when a .csv is opened
// (resolveOpenTarget: read -> decode -> convert -> temp file -> sidecar open
// -> indexed deep reads), without any Electron/GUI — the engine-level
// stand-in for the UI "fullyLoaded" metric, which cannot be measured while
// Electron windows are degraded on this machine (OBS-1665).
//
// Usage: tsx scripts/bench-csv-import.ts <csv-path> [reps]
//   SHEETS_SIDECAR_BIN overrides the sidecar binary (defaults to the
//   workspace-local release build; may point at any equivalent build).
//   BENCH_REFERENCE=1 additionally times the superseded internals
//   (parseCsv / worksheetXml standalone, JSZip deflate) as profile context;
//   those numbers never enter the total.
import { readFileSync } from 'node:fs'
import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'

import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import {
  buildWorksheetXml,
  csvToXlsxBuffer,
  decodeCsvBuffer,
  parseCsv,
} from '../src/gateway/csv-import'

const REPS = Number(process.argv[3] ?? 3)
const REFERENCE = process.env.BENCH_REFERENCE === '1'

async function sidecarBinary(): Promise<string> {
  const override = process.env.SHEETS_SIDECAR_BIN
  if (override) return resolve(override)
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}

function peakRssKilobytes(pid: number): number {
  try {
    // /proc status VmHWM is the kernel-maintained high-water mark, no sampling race
    const status = readFileSync(`/proc/${pid}/status`, 'utf8')
    const match = /VmHWM:\s+(\d+)/.exec(status)
    return match ? Number(match[1]) : 0
  } catch {
    return 0
  }
}

interface Phase {
  readonly name: string
  readonly ms: number
}

/// Reference-only profile lines (superseded internals); excluded from total.
async function referencePhases(text: string): Promise<Phase[]> {
  const phases: Phase[] = []
  let start = performance.now()
  const rows = parseCsv(text)
  phases.push({ name: '[ref] parseCsv', ms: performance.now() - start })
  start = performance.now()
  const xml = buildWorksheetXml(rows)
  phases.push({ name: '[ref] worksheetXml', ms: performance.now() - start })
  start = performance.now()
  const zip = new JSZip()
  zip.file('xl/worksheets/sheet1.xml', xml)
  await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  phases.push({ name: '[ref] JSZip deflate', ms: performance.now() - start })
  return phases
}

async function timePhases(
  csvPath: string,
  directory: string,
  client: XlsxSidecarClient,
): Promise<{ phases: Phase[]; totalMs: number; rows: number; verified: boolean }> {
  const phases: Phase[] = []
  const mark = (name: string, start: number): void => {
    phases.push({ name, ms: performance.now() - start })
  }

  let start = performance.now()
  const bytes = await readFile(csvPath)
  const text = decodeCsvBuffer(bytes)
  mark('read+decode', start)

  // csvToXlsxBuffer is the shipped conversion: parseCsv + worksheet XML +
  // zip container, exactly what resolveOpenTarget awaits
  start = performance.now()
  const buffer = await csvToXlsxBuffer(text)
  mark('convert (csvToXlsxBuffer)', start)

  const xlsxPath = join(directory, 'imported.xlsx')
  start = performance.now()
  await writeFile(xlsxPath, buffer)
  mark('writeTemp', start)

  start = performance.now()
  const opened = (await client.open(xlsxPath)) as {
    sessionId: string
    sheets: { id: string; rowCount: number; columnCount: number; name: string }[]
  }
  mark('sidecar open', start)
  const sheet = opened.sheets[0]
  if (!sheet) throw new Error('sidecar returned no sheets')

  // deep reads: first and last data rows must round-trip through the engine.
  // Large sheets index in the background (same as benchmark-large-xlsx), so
  // the tail read polls until indexing completes — that wait is part of the
  // engine-level fullyLoaded figure.
  start = performance.now()
  let tail: {
    cells: { row: number; column: number; value: unknown }[]
    indexingComplete?: boolean
  } | null = null
  for (let attempt = 0; attempt < 600; attempt += 1) {
    tail = (await client.readRange({
      sessionId: opened.sessionId,
      sheetId: sheet.id,
      range: {
        // sidecar ranges are inclusive, so the last row is rowCount-1
        startRow: sheet.rowCount - 1,
        endRow: sheet.rowCount - 1,
        startColumn: 0,
        endColumn: 1,
      },
    })) as { cells: { row: number; column: number; value: unknown }[]; indexingComplete?: boolean }
    if (tail.indexingComplete !== false && tail.cells.length > 0) break
    await new Promise((waitResolve) => setTimeout(waitResolve, 50))
  }
  const head = (await client.readRange({
    sessionId: opened.sessionId,
    sheetId: sheet.id,
    range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
  })) as { cells: { row: number; column: number; value: unknown }[] }
  const tailValue = tail?.cells.find((cell) => cell.column === 0)?.value
  const headValue = head.cells.find((cell) => cell.column === 0)?.value
  mark('sidecar deep reads (index wait)', start)
  await client.close(opened.sessionId)

  // slim 500k-correctness: the engine saw every parsed row, head and tail
  // round-trip with values (the counting parse runs untimed, after the marks)
  const expectedRows = parseCsv(text).length
  const verified =
    sheet.rowCount === expectedRows && headValue !== undefined && tailValue !== undefined
  console.log(`    head=${JSON.stringify(headValue)} tail=${JSON.stringify(tailValue)}`)
  const totalMs = phases.reduce((sum, phase) => sum + phase.ms, 0)
  return { phases, totalMs, rows: sheet.rowCount, verified }
}

async function main(): Promise<void> {
  const csvPath = process.argv[2]
  if (!csvPath)
    throw new Error('Pass a CSV path: tsx scripts/bench-csv-import.ts <csv-path> [reps]')
  const { size } = await stat(csvPath)
  console.log(`csv=${csvPath} bytes=${size} reps=${REPS} reference=${REFERENCE}`)
  const benchPid = process.pid

  const client = new XlsxSidecarClient(await sidecarBinary())
  const directory = await mkdtemp(join(tmpdir(), 'bench-csv-'))
  try {
    for (let rep = 1; rep <= REPS; rep += 1) {
      const sidecarPid = client.getProcessId()
      const rssBefore = sidecarPid ? peakRssKilobytes(sidecarPid) : 0
      const run = await timePhases(csvPath, directory, client)
      const rssAfter = sidecarPid ? peakRssKilobytes(sidecarPid) : 0
      if (REFERENCE) {
        const bytes = await readFile(csvPath)
        for (const phase of await referencePhases(decodeCsvBuffer(bytes))) run.phases.push(phase)
      }
      console.log(
        `rep ${rep}: total ${run.totalMs.toFixed(0)}ms rows=${run.rows} verified=${run.verified} ` +
          `sidecarPeakRssMB=${(Math.max(rssBefore, rssAfter) / 1024).toFixed(0)} ` +
          `benchPeakRssMB=${(peakRssKilobytes(benchPid) / 1024).toFixed(0)}`,
      )
      for (const phase of run.phases) {
        console.log(`    ${phase.name.padEnd(32)} ${phase.ms.toFixed(0).padStart(7)}ms`)
      }
    }
  } finally {
    client.stop()
    await rm(directory, { recursive: true, force: true })
  }
}

void main()
