// Locates the xlsx sidecar binary for the headless MCP server (Phase 3).
//
// Resolution order (SA4 section 4.3, adapted for a package instead of the
// Electron shell):
//   1. AIRY_XLSX_SIDECAR env var — explicit override, wins unconditionally
//      (tests point it at fixture/stub binaries too).
//   2. A repo checkout: walk up from this module looking for
//      apps/sheets/native/xlsx-engine/Cargo.toml, then take
//      target/release/xlsx-sidecar[.exe] (the `npm run native:build` output
//      and the macOS universal lipo target — same path).
//   3. An installed Airy app (or a legacy GenOffice install): process.resourcesPath/native/<exe>
//      (electron-builder extraResources layout).
//
// Returns null (never throws) when nothing is found so open_document can
// produce one clear, actionable error instead of a spawn crash.
import { existsSync } from 'node:fs'
import { access, constants } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const SIDECAR_ENV = 'AIRY_XLSX_SIDECAR'

const EXE = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
const REPO_MARKER = join('apps', 'sheets', 'native', 'xlsx-engine', 'Cargo.toml')

async function isFile(path: string): Promise<boolean> {
  try {
    await access(path, constants.F_OK)
    return true
  } catch {
    return false
  }
}

/**
 * Closest ancestor directory of `start` (exclusive) that contains the
 * sheets-engine Cargo manifest, i.e. the repo root in a checkout. Works from
 * raw TS (src/xlsx/…) and from the esbuild bundle (dist/index.js) alike,
 * since both live inside packages/mcp-server/.
 */
export function repoRootOf(start: string, maxDepth = 6): string | null {
  let current = start
  for (let depth = 0; depth < maxDepth; depth += 1) {
    const parent = dirname(current)
    if (current === parent) break
    current = parent
    if (isFileSync(join(current, REPO_MARKER))) return current
  }
  return null
}

function isFileSync(path: string): boolean {
  try {
    return existsSync(path)
  } catch {
    return false
  }
}

/** Candidate paths in priority order; the first existing one wins. */
export async function sidecarCandidates(fromDir?: string): Promise<string[]> {
  const candidates: string[] = []
  const fromEnv = process.env[SIDECAR_ENV]
  if (fromEnv && fromEnv.trim() !== '') candidates.push(fromEnv.trim())

  const start = fromDir ?? dirname(fileURLToPath(import.meta.url))
  const root = repoRootOf(start)
  if (root) {
    candidates.push(join(root, 'apps', 'sheets', 'native', 'xlsx-engine', 'target', 'release', EXE))
  }

  // Electron-only (extraResources layout); absent in plain Node processes
  const resources = (process as NodeJS.Process & { resourcesPath?: string }).resourcesPath
  if (resources) candidates.push(join(resources, 'native', EXE))
  return candidates
}

/** Absolute path of the sidecar binary, or null when unavailable. */
export async function findSidecarBinary(fromDir?: string): Promise<string | null> {
  for (const candidate of await sidecarCandidates(fromDir)) {
    if (await isFile(candidate)) return candidate
  }
  return null
}

/** Agent-facing error for the no-binary case, with build instructions. */
export function sidecarMissingError(): Error {
  return new Error(
    'The xlsx sidecar binary is not available. Build it with ' +
      '`npm run native:build -w @airy-office/sheets` (requires cargo/rust), or point the ' +
      `${SIDECAR_ENV} env var at an existing xlsx-sidecar binary.`,
  )
}
