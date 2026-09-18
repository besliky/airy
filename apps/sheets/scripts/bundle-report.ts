/// Renderer bundle size report (PERF-502 split, PERF-501 metrics hook).
///
/// Reads apps/sheets/out/renderer after `electron-vite build` and prints the
/// chunk table: which chunks exist, which load eagerly at startup (the entry
/// plus everything index.html modulepreloads), which stay lazy (Univer
/// locale tables, the EMF/WMF metafile decoder), and whether any single
/// chunk crosses the 5 MB budget from PERF-502. Exits non-zero on a budget
/// violation so it can gate a build in CI later.
///
/// Usage: npm run bundle-report [-w @airy-office/sheets] [--json]
import { readdir, readFile, stat } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

/// Largest allowed single chunk (bytes) — the PERF-502 goal was splitting
/// the 18.56 MB renderer monolith into logical chunks below this size.
const CHUNK_BUDGET_BYTES = 5 * 1024 * 1024

const AS_JSON = process.argv.includes('--json')

const appRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const outDir = join(appRoot, 'out', 'renderer')

function kilobytes(bytes: number): string {
  return (bytes / 1024).toLocaleString('en-US', { maximumFractionDigits: 1 })
}

async function eagerChunks(): Promise<Set<string>> {
  // The entry script plus modulepreload links are the startup set; every
  // other asset only loads through a dynamic import.
  const html = await readFile(join(outDir, 'index.html'), 'utf8')
  const refs = [...html.matchAll(/(?:src|href)="\.\/(assets\/[^"]+)"/g)].flatMap((m) =>
    m[1]?.endsWith('.js') ? [m[1]] : [],
  )
  return new Set(refs)
}

interface ChunkRow {
  name: string
  bytes: number
  eager: boolean
}

async function collectJsChunks(eager: Set<string>): Promise<ChunkRow[]> {
  const assetsDir = join(outDir, 'assets')
  const rows: ChunkRow[] = []
  for (const entry of await readdir(assetsDir)) {
    if (!entry.endsWith('.js')) continue
    const bytes = (await stat(join(assetsDir, entry))).size
    rows.push({ name: entry, bytes, eager: eager.has(`assets/${entry}`) })
  }
  rows.sort((a, b) => b.bytes - a.bytes || a.name.localeCompare(b.name))
  return rows
}

async function main(): Promise<number> {
  const eager = await eagerChunks()
  const rows = await collectJsChunks(eager)
  if (rows.length === 0) {
    console.error(`No JS chunks under ${outDir} — run \`npm run build\` first.`)
    return 2
  }
  const eagerBytes = rows.filter((r) => r.eager).reduce((sum, r) => sum + r.bytes, 0)
  const lazyBytes = rows.filter((r) => !r.eager).reduce((sum, r) => sum + r.bytes, 0)
  const overBudget = rows.filter((r) => r.bytes > CHUNK_BUDGET_BYTES)

  if (AS_JSON) {
    console.log(
      JSON.stringify(
        {
          budgetBytes: CHUNK_BUDGET_BYTES,
          eagerBytes,
          lazyBytes,
          overBudget: overBudget.map((r) => r.name),
          chunks: rows,
        },
        null,
        2,
      ),
    )
    return overBudget.length > 0 ? 1 : 0
  }

  console.log(`Renderer chunks under ${outDir} (budget ${kilobytes(CHUNK_BUDGET_BYTES)} kB/chunk):`)
  for (const row of rows) {
    const flag = row.bytes > CHUNK_BUDGET_BYTES ? '  <-- OVER BUDGET' : ''
    console.log(
      `${kilobytes(row.bytes).padStart(10)} kB  ${row.eager ? 'eager' : 'lazy '}  ${row.name}${flag}`,
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
    return 1
  }
  console.log('OK: every chunk is within the budget.')
  return 0
}

void main().then(
  (code) => process.exit(code),
  (error) => {
    console.error(error)
    process.exit(2)
  },
)
