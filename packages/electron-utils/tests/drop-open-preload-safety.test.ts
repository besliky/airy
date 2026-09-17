import { readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

/**
 * `@airy-office/electron-utils/drop-open` is imported by every app's
 * sandboxed preload. Sandboxed preloads cannot require Node built-ins — a
 * single transitive `node:*` import breaks the whole preload at load time
 * ("module not found: node:path"), the contextBridge API never appears, and
 * every renderer in the app boots to a blank page. That failure mode only
 * reproduces in a real Electron preload (unit tests run in plain Node, where
 * `node:path` resolves fine), so this test walks the static import graph of
 * drop-open.ts and fails on any Node built-in (bare specifier or `node:`)
 * reachable from it.
 */

const SRC = resolve(__dirname, '../src')
const IMPORT_RE = /(?:^|\n)\s*import\s+(?:[^'"]*?\sfrom\s+)?['"]([^'"]+)['"]/g

function relativeImportsOf(file: string): string[] {
  const specifiers: string[] = []
  const source = readFileSync(file, 'utf8')
  for (const match of source.matchAll(IMPORT_RE)) specifiers.push(match[1])
  return specifiers.filter((specifier) => specifier.startsWith('.'))
}

function resolveSource(file: string): string {
  try {
    readFileSync(file)
    return file
  } catch {
    return `${file}.ts`
  }
}

describe('drop-open preload safety', () => {
  it('pulls no node built-ins through its static import graph', () => {
    const seen = new Set<string>()
    const queue = [join(SRC, 'drop-open.ts')]
    while (queue.length > 0) {
      const file = resolveSource(queue.shift() as string)
      if (seen.has(file)) continue
      seen.add(file)
      const source = readFileSync(file, 'utf8')
      for (const match of source.matchAll(IMPORT_RE)) {
        const specifier = match[1]
        const bare = !specifier.startsWith('.') && !specifier.startsWith('@')
        const allowedBare = specifier === 'electron'
        expect(
          specifier.startsWith('node:') || (bare && !allowedBare),
          `${file} imports "${specifier}" — sandboxed preloads cannot load Node built-ins`,
        ).toBe(false)
      }
      for (const relative of relativeImportsOf(file)) {
        queue.push(resolve(dirname(file), relative))
      }
    }
    // the graph is non-trivial (electron + at least the shared channel)
    expect(seen.size).toBeGreaterThan(1)
  })
})
