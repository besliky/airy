/// Renderer bundle size report (PERF-502 split, PERF-501 metrics hook,
/// PERF-1104 entry gate).
///
/// Reads apps/sheets/out/renderer after `electron-vite build` and prints the
/// chunk table: which chunks exist, which load eagerly at startup (the entry
/// plus everything index.html modulepreloads), which stay lazy (Univer
/// locale tables, the EMF/WMF metafile decoder), and whether budgets hold.
/// Two budgets: any single chunk over 5 MB (the shared PERF-502 ceiling)
/// and the entry chunk over 3,000 kB — the PERF-1104 gate. A source-map
/// audit of the entry chunk at 2,685 kB found ~407 kB of eagerly-bundled
/// en-US Univer locale packs (App.tsx imports them statically because
/// createUniver needs them synchronously), led by
/// @univerjs/sheets-formula's en-US tables alone at ~354 kB — the lazy
/// candidate PERF-1104 flags for a follow-up (deferring it means an async
/// createUniver bootstrap); the rest is first-party app code (univer-sync,
/// ExcelShell, App). 3,000 kB keeps ~11% headroom over 2,685 kB, matching
/// the docs gate's headroom, while still catching a regression toward the
/// pre-split monolith. Exits non-zero on a violation so CI can gate the
/// build before e2e minutes are spent.
///
/// Usage: npm run bundle-report [-w @airy-office/sheets] [--json]
import { readdir, readFile, stat } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

/// Largest allowed single chunk (bytes) — the PERF-502 goal was splitting
/// the 18.56 MB renderer monolith into logical chunks below this size.
const CHUNK_BUDGET_BYTES = 5 * 1024 * 1024

/// Largest allowed entry chunk (bytes) — PERF-1104: the entry was 2,679.5 kB
/// at v0.12.0 and 2,685 kB at the audit; see the header for its composition
/// and the flagged lazy candidate. 3,000 kB holds ~11% headroom.
const ENTRY_BUDGET_BYTES = 3_000 * 1024

const AS_JSON = process.argv.includes('--json')

const appRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const outDir = join(appRoot, 'out', 'renderer')

function kilobytes(bytes: number): string {
  return (bytes / 1024).toLocaleString('en-US', { maximumFractionDigits: 1 })
}

async function startupChunks(): Promise<{ eager: Set<string>; entry: string | null }> {
  // The entry script plus modulepreload links are the startup set; every
  // other asset only loads through a dynamic import.
  const html = await readFile(join(outDir, 'index.html'), 'utf8')
  const refs = [...html.matchAll(/(?:src|href)="\.\/(assets\/[^"]+)"/g)].flatMap((m) =>
    m[1]?.endsWith('.js') ? [m[1]] : [],
  )
  const entryMatch = /<script\b[^>]*\bsrc="\.\/(assets\/[^"]+?\.js)"/.exec(html)
  return { eager: new Set(refs), entry: entryMatch?.[1] ?? null }
}

interface ChunkRow {
  name: string
  bytes: number
  eager: boolean
  entry: boolean
}

async function collectJsChunks(eager: Set<string>, entry: string | null): Promise<ChunkRow[]> {
  const assetsDir = join(outDir, 'assets')
  const rows: ChunkRow[] = []
  for (const entryName of await readdir(assetsDir)) {
    if (!entryName.endsWith('.js')) continue
    const bytes = (await stat(join(assetsDir, entryName))).size
    rows.push({
      name: entryName,
      bytes,
      eager: eager.has(`assets/${entryName}`),
      entry: `assets/${entryName}` === entry,
    })
  }
  rows.sort((a, b) => b.bytes - a.bytes || a.name.localeCompare(b.name))
  return rows
}

async function main(): Promise<number> {
  const { eager, entry } = await startupChunks()
  const rows = await collectJsChunks(eager, entry)
  if (rows.length === 0) {
    console.error(`No JS chunks under ${outDir} — run \`npm run build\` first.`)
    return 2
  }
  if (entry === null) {
    console.error(
      `No entry script found in ${join(outDir, 'index.html')} — unexpected build layout.`,
    )
    return 2
  }
  const eagerBytes = rows.filter((r) => r.eager).reduce((sum, r) => sum + r.bytes, 0)
  const lazyBytes = rows.filter((r) => !r.eager).reduce((sum, r) => sum + r.bytes, 0)
  const overBudget = rows.filter((r) => r.bytes > CHUNK_BUDGET_BYTES)
  const entryRow = rows.find((r) => r.entry)!
  const entryOverBudget = entryRow.bytes > ENTRY_BUDGET_BYTES

  if (AS_JSON) {
    console.log(
      JSON.stringify(
        {
          budgetBytes: CHUNK_BUDGET_BYTES,
          entryBudgetBytes: ENTRY_BUDGET_BYTES,
          entryChunk: entryRow.name,
          entryBytes: entryRow.bytes,
          entryOverBudget,
          eagerBytes,
          lazyBytes,
          overBudget: overBudget.map((r) => r.name),
          chunks: rows,
        },
        null,
        2,
      ),
    )
    return overBudget.length > 0 || entryOverBudget ? 1 : 0
  }

  console.log(
    `Renderer chunks under ${outDir} (budget ${kilobytes(CHUNK_BUDGET_BYTES)} kB/chunk, entry ≤ ${kilobytes(ENTRY_BUDGET_BYTES)} kB):`,
  )
  for (const row of rows) {
    const flag =
      row.bytes > CHUNK_BUDGET_BYTES
        ? '  <-- OVER BUDGET'
        : row.entry && row.bytes > ENTRY_BUDGET_BYTES
          ? '  <-- ENTRY OVER BUDGET'
          : ''
    console.log(
      `${kilobytes(row.bytes).padStart(10)} kB  ${row.eager ? 'eager' : 'lazy '}  ${row.entry ? 'entry ' : '      '}  ${row.name}${flag}`,
    )
  }
  const eagerCount = rows.filter((r) => r.eager).length
  console.log('')
  console.log(`eager (entry + modulepreload): ${eagerCount} chunks, ${kilobytes(eagerBytes)} kB`)
  console.log(
    `lazy  (dynamic import):         ${rows.length - eagerCount} chunks, ${kilobytes(lazyBytes)} kB`,
  )
  if (overBudget.length > 0) {
    console.error(
      `FAIL: ${overBudget.length} chunk(s) exceed the ${kilobytes(CHUNK_BUDGET_BYTES)} kB budget`,
    )
  }
  if (entryOverBudget) {
    console.error(
      `FAIL: the entry chunk is ${kilobytes(entryRow.bytes)} kB, over the ${kilobytes(ENTRY_BUDGET_BYTES)} kB entry budget`,
    )
  }
  if (overBudget.length > 0 || entryOverBudget) return 1
  console.log('OK: every chunk is within the budgets.')
  return 0
}

void main().then(
  (code) => process.exit(code),
  (error) => {
    console.error(error)
    process.exit(2)
  },
)
