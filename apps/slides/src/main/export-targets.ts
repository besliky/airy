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
import { basename, dirname, join, posix, win32 } from 'node:path'

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
 * `..` traversal, no control characters. Windows-reserved characters are
 * rejected too — `:` would create an NTFS alternate data stream
 * (`name.png:hidden`) and `*?"< >|` are illegal in Windows filenames
 * (hygiene: the export would fail there anyway). Returns the trimmed name,
 * or null when the name could escape the target directory or misbehave.
 */
export function sanitizeExportBaseName(baseName: string): string | null {
  const name = baseName.trim()
  if (!name || name === '.' || name === '..') return null
  if (name.includes('/') || name.includes('\\')) return null
  // eslint-disable-next-line no-control-regex
  if (/[\u0000-\u001f\u007f]/.test(name)) return null
  if (/[:*?"<>|]/.test(name)) return null
  return name
}

/** Zero-padding width follows the total page count (3 digits for >= 100 pages). */
function padWidth(pageCount: number): number {
  return pageCount >= 100 ? 3 : 2
}

/** Supported export image formats: PNG (default) and JPEG. */
export type ExportImageExt = 'png' | 'jpg'

/**
 * Resolve the per-page image paths of an images export inside `dir`:
 * `<name>-01.png`, `<name>-02.png`, … (`ext='jpg'` → `.jpg`). Null when the
 * base name is unsafe, the extension is unsupported, or a resolved target
 * escapes `dir` (defense in depth — the sanitized name alone already prevents
 * traversal).
 */
export function resolveExportImagePaths(
  dir: string,
  baseName: string,
  pageCount: number,
  platform: NodeJS.Platform = process.platform,
  ext: ExportImageExt = 'png',
): string[] | null {
  if (ext !== 'png' && ext !== 'jpg') return null
  const name = sanitizeExportBaseName(baseName)
  if (!name || !Number.isFinite(pageCount) || pageCount < 1 || pageCount > 10_000) return null
  const pad = padWidth(pageCount)
  const paths: string[] = []
  for (let i = 1; i <= pageCount; i++) {
    const p = join(dir, `${name}-${String(i).padStart(pad, '0')}.${ext}`)
    if (!isPathInsideDir(dir, p, platform)) return null
    paths.push(p)
  }
  return paths
}

// ── physical containment ──────────────────────────────────────────────────

/** throws when the path (or its deepest existing ancestor) cannot be resolved */
export type RealpathFn = (path: string) => string

/**
 * The realpath when it exists; for a non-existent path, the physical path of
 * its deepest EXISTING ancestor with the missing tail components appended —
 * so a not-yet-created export subdirectory still resolves to the physical
 * folder it would be created in (mirrors mcp-server's docx/paths.ts). A
 * symlink swapped into the path after the pick therefore resolves to its
 * real target and fails the containment check below.
 */
export function realPathOrDeepestExisting(path: string, realpath: RealpathFn): string {
  try {
    return realpath(path)
  } catch {
    const tail: string[] = []
    let current = path
    for (;;) {
      const parent = dirname(current)
      // dirname(x) === x only at the filesystem root — nothing exists to
      // resolve against; the lexical spelling is the best answer left
      if (parent === current) return path
      tail.unshift(basename(current))
      current = parent
      try {
        return join(realpath(current), ...tail)
      } catch {
        // this ancestor does not exist either — keep climbing
      }
    }
  }
}

/**
 * Containment of a renderer-named export directory against the REALPATH of
 * the directory the user picked (stored at pick time): the target is
 * re-resolved physically (deepest existing ancestor handles fresh
 * subdirectories) before the prefix check, so a later symlink swap of the
 * picked dir — or of any intermediate component — cannot pass by spelling.
 */
export function exportDirInsidePick(
  pickedRealDir: string,
  dir: string,
  realpath: RealpathFn,
  platform: NodeJS.Platform = process.platform,
): boolean {
  const physical = realPathOrDeepestExisting(dir, realpath)
  return isPathInsideDir(pickedRealDir, physical, platform)
}

/** Same physical check for the picked PDF file path (a fresh target resolves
 *  through its deepest existing ancestor, so saving a new file still works). */
export function exportFileMatchesPick(
  pickedRealFile: string,
  filePath: string,
  realpath: RealpathFn,
  platform: NodeJS.Platform = process.platform,
): boolean {
  const physical = realPathOrDeepestExisting(filePath, realpath)
  return normalizeForCompare(physical, platform) === normalizeForCompare(pickedRealFile, platform)
}
