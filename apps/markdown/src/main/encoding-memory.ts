/// Remembered per-file encodings (UX-1653). The plain-text editors decode
/// user files with a scoring detector (packages/file-parse/src/text.ts) whose
/// guess can be wrong (BUG-1646 / BUG-1651); when the user overrides it with
/// a manual pick, that pick must outlive the tab — reopening the same path
/// would otherwise run the same detector and produce the same mojibake.
///
/// The store is workspace-level: one `fileEncodings` list (path → charset,
/// most-recently-used first, capped) inside the shared userData/
/// app-settings.json that the shell and every editor already write through
/// the single-writer queue (packages/electron-utils/src/app-settings-file.ts).
/// Writes go exclusively through queueAppSettingsUpdate, so a pick landing
/// while other settings writes are in flight can neither drop their keys nor
/// be dropped by them (OBS-1532 discipline). Every open re-reads the file
/// fresh, so a pick is visible immediately — to other windows and to the
/// other editor apps sharing the userData — with no in-memory cache to
/// invalidate.
import { queueAppSettingsUpdate, readAppSettingsFile } from '@airy-office/electron-utils'

import { SELECTABLE_ENCODINGS } from '../shared/ipc'
import type { SelectableEncoding } from '../shared/ipc'

export { SELECTABLE_ENCODINGS }
export type { SelectableEncoding }

const SELECTABLE_ENCODING_SET: ReadonlySet<string> = new Set(SELECTABLE_ENCODINGS)

export function isSelectableEncoding(value: unknown): value is SelectableEncoding {
  return typeof value === 'string' && SELECTABLE_ENCODING_SET.has(value)
}

/** one remembered file encoding; the persisted list is most-recently-used first */
export interface FileEncodingEntry {
  path: string
  encoding: string
}

const FILE_ENCODINGS_KEY = 'fileEncodings'
const FILE_ENCODINGS_CAP = 200

/**
 * Record a pick: move the path's entry to the front (or add it) and cap the
 * list. Pure — unit-tested; the persisted JSON keeps this order.
 */
export function recordFileEncoding(
  entries: FileEncodingEntry[],
  filePath: string,
  encoding: string,
  cap = FILE_ENCODINGS_CAP,
): FileEncodingEntry[] {
  const rest = entries.filter((entry) => entry.path !== filePath)
  return [{ path: filePath, encoding }, ...rest].slice(0, cap)
}

/** parse the persisted LRU; malformed values are dropped entry-by-entry */
export function parseFileEncodings(raw: unknown): FileEncodingEntry[] {
  if (!Array.isArray(raw)) return []
  const entries: FileEncodingEntry[] = []
  for (const item of raw) {
    if (
      item &&
      typeof item === 'object' &&
      !Array.isArray(item) &&
      typeof (item as FileEncodingEntry).path === 'string' &&
      (item as FileEncodingEntry).path &&
      isSelectableEncoding((item as FileEncodingEntry).encoding)
    ) {
      entries.push({
        path: (item as FileEncodingEntry).path,
        encoding: (item as FileEncodingEntry).encoding,
      })
    }
  }
  return entries
}

/** the remembered charset for a path, or undefined when auto-detection applies */
export function readRememberedFileEncoding(
  settingsPath: string,
  filePath: string,
): string | undefined {
  return parseFileEncodings(readAppSettingsFile(settingsPath)[FILE_ENCODINGS_KEY]).find(
    (entry) => entry.path === filePath,
  )?.encoding
}

/**
 * Persist a pick through the shared single-writer queue: the read-modify-write
 * runs as one queued section and re-reads the on-disk state inside it, so a
 * pick computed earlier can never clobber keys written in between. Rejects on
 * I/O errors; callers treat persistence as best-effort — the open still
 * works, auto-detected.
 */
export function rememberFileEncoding(
  settingsPath: string,
  filePath: string,
  encoding: string,
): Promise<void> {
  return queueAppSettingsUpdate(settingsPath, (current) => ({
    ...current,
    [FILE_ENCODINGS_KEY]: recordFileEncoding(
      parseFileEncodings(current[FILE_ENCODINGS_KEY]),
      filePath,
      encoding,
    ),
  }))
}

/** drop a path's pick so the file falls back to auto-detection (the "Auto" pick). Pure — unit-tested */
export function removeFileEncoding(
  entries: FileEncodingEntry[],
  filePath: string,
): FileEncodingEntry[] {
  return entries.filter((entry) => entry.path !== filePath)
}

/** Forget a pick through the shared single-writer queue (same discipline as rememberFileEncoding). */
export function forgetFileEncoding(settingsPath: string, filePath: string): Promise<void> {
  return queueAppSettingsUpdate(settingsPath, (current) => ({
    ...current,
    [FILE_ENCODINGS_KEY]: removeFileEncoding(
      parseFileEncodings(current[FILE_ENCODINGS_KEY]),
      filePath,
    ),
  }))
}

function tryDecode(bytes: Uint8Array, charset: string, keepBom: boolean): string | null {
  try {
    return new TextDecoder(charset, { ignoreBOM: keepBom }).decode(bytes)
  } catch {
    return null
  }
}

/**
 * Decode bytes AS the remembered charset. A byte-order mark still wins — an
 * explicit byte signature outranks even a manual pick, and a file with a BOM
 * never needs one (it decodes correctly anyway). Everything else is decoded
 * exactly as chosen: unlike the detector's `preferred` hint, which only
 * breaks score ties, a remembered pick is authoritative, so neither the
 * strict-UTF-8 probe nor a declared HTML charset runs first. The kept-BOM
 * decode matches what the editors' open path always produced, so an
 * untouched open→save still round-trips byte-identically.
 */
export function decodeBytesAsEncoding(bytes: Uint8Array, encoding: string): string {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return tryDecode(bytes, 'utf-8', true) ?? ''
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return tryDecode(bytes, 'utf-16le', true) ?? ''
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return tryDecode(bytes, 'utf-16be', true) ?? ''
  return tryDecode(bytes, encoding, true) ?? ''
}
