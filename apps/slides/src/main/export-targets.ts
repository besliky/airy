/**
 * Export target validation for slides:export-images / slides:export-pdf.
 *
 * Both channels write renderer-supplied bytes to a renderer-named
 * destination, so a compromised renderer could otherwise overwrite
 * arbitrary files (path traversal via the base name, or any path for the
 * PDF). The main process remembers the directory / file each webContents
 * last picked through the dedicated pick channels (slides:pick-export-dir
 * / slides:pick-export-pdf-path) and confines the export to that pick.
 * The helpers here are pure so the validation can be unit-tested.
 */
import { join, posix, win32 } from 'node:path'

import { isPathInsideDir } from '@airy-office/electron-utils'

/** Normalize for platform-aware equality (win32 compares case-insensitively). */
function normalizeForCompare(path: string, platform: NodeJS.Platform): string {
  const resolved = platform === 'win32' ? win32.resolve(path) : posix.resolve(path)
  return platform === 'win32' ? resolved.toLowerCase() : resolved
}

/** True when both paths resolve to the same file on the given platform. */
export function isSameExportFile(
  a: string,
  b: string,
  platform: NodeJS.Platform = process.platform,
): boolean {
  const av = a.trim()
  const bv = b.trim()
  if (!av || !bv) return false
  return normalizeForCompare(av, platform) === normalizeForCompare(bv, platform)
}

/**
 * Export base name must be a single path segment: no separators, no `.` /
 * `..` traversal, no control characters. Returns the trimmed name, or null
 * when the name could escape the target directory.
 */
export function sanitizeExportBaseName(baseName: string): string | null {
  const name = baseName.trim()
  if (!name || name === '.' || name === '..') return null
  if (name.includes('/') || name.includes('\\')) return null
  // eslint-disable-next-line no-control-regex
  if (/[\u0000-\u001f\u007f]/.test(name)) return null
  return name
}

/** Zero-padding width follows the total page count (3 digits for >= 100 pages). */
function padWidth(pageCount: number): number {
  return pageCount >= 100 ? 3 : 2
}

/**
 * Resolve the per-page PNG paths of an images export inside `dir`:
 * `<name>-01.png`, `<name>-02.png`, ... Null when the base name is unsafe
 * or a resolved target escapes `dir` (defense in depth — the sanitized
 * name alone already prevents traversal).
 */
export function resolveExportImagePaths(
  dir: string,
  baseName: string,
  pageCount: number,
  platform: NodeJS.Platform = process.platform,
): string[] | null {
  const name = sanitizeExportBaseName(baseName)
  if (!name || !Number.isFinite(pageCount) || pageCount < 1 || pageCount > 10_000) return null
  const pad = padWidth(pageCount)
  const paths: string[] = []
  for (let i = 1; i <= pageCount; i++) {
    const p = join(dir, `${name}-${String(i).padStart(pad, '0')}.png`)
    if (!isPathInsideDir(dir, p, platform)) return null
    paths.push(p)
  }
  return paths
}
