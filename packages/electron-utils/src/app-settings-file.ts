/// Single-writer discipline for userData/app-settings.json — a flat JSON
/// object shared by the shell and every editor main (language, theme,
/// onboarding flag, dialog-dir LRU, …). Its correctness rule is simple but
/// easy to break: every write must be a read-merge-write that keeps keys it
/// did not touch.
///
/// OBS-1532: an audit twice observed a seeded `onboardingSeen` key vanish
/// mid-run while only {starPrompt, lastRunVersion} remained (self-healed by
/// the next merge-write). A static audit found no wholesale writer, but the
/// merge-write existed as two independent copies (shell app-settings.ts and
/// dialog-memory.ts) and the one async-context writer had no serialization.
/// This module closes the whole class:
///   1. ONE merge implementation — nothing else may write the file, so a
///      wholesale JSON.stringify writer cannot (re)appear unnoticed.
///   2. Async read-modify-write goes through a per-path promise queue: one
///      queued write in flight at a time, later writes chain behind it, and
///      the re-read happens inside the queued section, so no caller can
///      persist a snapshot taken before another write landed.
///   3. Synchronous merge writes stay synchronous (Electron's single main
///      thread already serializes them) and share the same code path.
/// The residual risk is a second PROCESS sharing the file (two instances on
/// one userData); the shell's single-instance lock covers that.
///
/// BUG-1771 (SET-26-1, family UX-1681): every unreadable byte in the file is
/// a TOTAL silent loss of ALL settings — readAppSettingsFile used to answer
/// {} and the next merge-write persisted that, overwriting the corrupt bytes.
/// Now a corrupt file is first preserved verbatim as `app-settings.json.bak`
/// (stable name, once per distinct corrupt state, before ANY rewrite) so a
/// user can forensically recover, and its content is salvaged as far as
/// strict JSON parsing allows (BOM prefix / garbage after a complete object /
/// truncated tail). A wrong-type root (array, number), empty or
/// whitespace-only content has no object data to salvage: honest defaults.
import { randomUUID } from 'node:crypto'
import {
  copyFileSync,
  readFileSync,
  renameSync,
  statSync,
  unlinkSync,
  utimesSync,
  writeFileSync,
} from 'node:fs'

/** app-settings.json is always a plain flat object of JSON values */
export type AppSettings = Record<string, unknown>

/**
 * Forensic copy kept beside the settings file whenever corrupt bytes are
 * about to be replaced (BUG-1771). A stable name is deliberate: one
 * discoverable file with unambiguous recovery instructions ("rename it back
 * to app-settings.json"), no unbounded `.corrupt-<timestamp>` litter in
 * userData; each corruption event refreshes it with the newest corrupt bytes
 * and the copy carries the corrupt file's own timestamps (the mtime answers
 * "when did this break?" after the original is overwritten).
 */
export const APP_SETTINGS_BACKUP_SUFFIX = '.bak'

/** UTF-8 byte-order mark — JSON.parse rejects it, "smart" external tools write it */
const BOM = '\uFEFF'

/**
 * A queued update: either a patch object merged over the fresh on-disk
 * state, or a reducer that receives the fresh state — read INSIDE its
 * queued section, never a snapshot captured earlier — and returns the next
 * full object.
 */
export type AppSettingsUpdate = AppSettings | ((current: AppSettings) => AppSettings)

function isSettingsObject(value: unknown): value is AppSettings {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

/** Strict parse to a plain settings object; anything else (or a throw) is null. */
function parseSettingsText(text: string): AppSettings | null {
  try {
    const parsed: unknown = JSON.parse(text)
    return isSettingsObject(parsed) ? parsed : null
  } catch {
    return null
  }
}

/** How corrupt bytes were salvaged, for the main log */
export type CorruptRecoveryMode = 'bom' | 'trailing' | 'truncated'

export interface CorruptAppSettingsRecovery {
  settings: AppSettings
  mode: CorruptRecoveryMode
}

/**
 * Structural map of a candidate text from one O(n) pass. Feeds the repair
 * tiers; every candidate it produces is still validated by JSON.parse, so a
 * scanner imprecision can only MISS a repair, never invent invalid JSON.
 */
interface TextScan {
  /** index of the `}` that closes the root object, or -1 (truncated/garbage) */
  rootCloseIndex: number
  /** last `,` outside strings at bracket depth >= 1, oldest-first, capped */
  commaCuts: Array<{ index: number; depth: number }>
  /** brackets still open at end of input, outermost first */
  openBrackets: string[]
  /** true when the input ends inside a string literal */
  endsInString: boolean
}

/** Repair tries at most this many member-boundary cuts, newest first */
const MAX_REPAIR_CUTS = 3

function scanText(text: string): TextScan {
  const scan: TextScan = {
    rootCloseIndex: -1,
    commaCuts: [],
    openBrackets: [],
    endsInString: false,
  }
  let inString = false
  let escaped = false
  for (let i = 0; i < text.length; i++) {
    const char = text[i]
    if (inString) {
      if (escaped) {
        escaped = false
      } else if (char === '\\') {
        escaped = true
      } else if (char === '"') {
        inString = false
      }
      continue
    }
    if (char === '"') {
      inString = true
    } else if (char === '{' || char === '[') {
      scan.openBrackets.push(char)
    } else if (char === '}' || char === ']') {
      const open = scan.openBrackets.pop()
      const matched = (char === '}' && open === '{') || (char === ']' && open === '[')
      if (!matched) {
        // mismatched closer: not structural — restore the depth and move on
        if (open !== undefined) scan.openBrackets.push(open)
      } else if (char === '}' && scan.openBrackets.length === 0 && scan.rootCloseIndex < 0) {
        scan.rootCloseIndex = i
      }
    } else if (char === ',' && scan.openBrackets.length >= 1) {
      // a member boundary: the text before it is structurally complete as far
      // as the scanner can tell, so it is a cut candidate for repair
      if (scan.commaCuts.length >= MAX_REPAIR_CUTS) scan.commaCuts.shift()
      scan.commaCuts.push({ index: i, depth: scan.openBrackets.length })
    }
  }
  scan.endsInString = inString
  return scan
}

/** Closers that would balance the brackets still open, innermost last. */
function closersFor(openBrackets: string[]): string {
  let closers = ''
  for (let i = openBrackets.length - 1; i >= 0; i--) {
    closers += openBrackets[i] === '{' ? '}' : ']'
  }
  return closers
}

/**
 * Salvage a plain settings object from corrupt file bytes (BUG-1771). Tiered,
 * strictly additive: every candidate is validated by a real JSON.parse, so the
 * function returns either parseable JSON or null — it never invents data, it
 * only keeps what already parses.
 *  - BOM prefix: strip it and an otherwise intact file parses fully ('bom').
 *  - Garbage after a complete root object: keep the object ('trailing').
 *  - Truncated tail: complete the open structure (keeping every member the
 *    scanner saw finish), or cut back to the newest member boundary that
 *    yields parseable JSON ('truncated').
 * A wrong-type root (array, number, string, null), empty or whitespace-only
 * text has no object content to salvage: null (honest defaults).
 */
export function recoverCorruptAppSettings(rawText: string): CorruptAppSettingsRecovery | null {
  const hasBom = rawText.startsWith(BOM)
  const text = hasBom ? rawText.slice(1) : rawText
  if (hasBom) {
    const intact = parseSettingsText(text)
    if (intact) return { settings: intact, mode: 'bom' }
  }
  // only a root object can hold settings; every other root is a dead end
  if (!/^\s*\{/.test(text)) return null
  const scan = scanText(text)
  if (scan.rootCloseIndex >= 0) {
    const prefix = parseSettingsText(text.slice(0, scan.rootCloseIndex + 1))
    if (prefix) return { settings: prefix, mode: 'trailing' }
  }
  // full completion first: it keeps strictly more members than any cut
  const closers = closersFor(scan.openBrackets)
  const completed = parseSettingsText(
    scan.endsInString ? `${text}"${closers}` : `${text}${closers}`,
  )
  if (completed) return { settings: completed, mode: 'truncated' }
  for (let i = scan.commaCuts.length - 1; i >= 0; i--) {
    const cut = scan.commaCuts[i]
    if (!cut) continue
    const head = text.slice(0, cut.index)
    const repaired = parseSettingsText(`${head}${closersFor(scanText(head).openBrackets)}`)
    if (repaired) return { settings: repaired, mode: 'truncated' }
  }
  return null
}

/**
 * Once-per-corrupt-state forensic bookkeeping: path -> fingerprint of the
 * corrupt bytes already preserved. readAppSettingsFile runs on every getter,
 * so without this dedup each read would re-copy and re-log the same corrupt
 * file; a valid read re-arms the path so a FUTURE corruption is preserved
 * again. Bounded by the number of settings paths the process touches.
 */
const preservedCorruptStates = new Map<string, string>()

/**
 * Preserve the corrupt bytes verbatim before anything can overwrite them.
 * Returns true when this corrupt state was seen for the first time (the one
 * occasion to log). Never throws: forensics must not fail the read.
 */
function preserveCorruptFile(settingsPath: string): boolean {
  try {
    const stat = statSync(settingsPath)
    const fingerprint = `${stat.size}:${stat.mtimeMs}`
    const fresh = preservedCorruptStates.get(settingsPath) !== fingerprint
    preservedCorruptStates.set(settingsPath, fingerprint)
    if (!fresh) return false
    const backupPath = `${settingsPath}${APP_SETTINGS_BACKUP_SUFFIX}`
    try {
      copyFileSync(settingsPath, backupPath)
      try {
        // keep the corrupt bytes' own timestamps on the copy: after the
        // original is overwritten, the .bak must still answer "when?"
        utimesSync(backupPath, stat.atime, stat.mtime)
      } catch {
        // cosmetic only — an untimestamped copy still preserves the bytes
      }
      console.warn(
        `[app-settings] corrupt settings file; preserved ${stat.size} bytes as ${backupPath} before any rewrite`,
      )
    } catch (error) {
      console.warn(
        `[app-settings] could not preserve corrupt settings file as ${backupPath}: ${String(error)}`,
      )
    }
    return true
  } catch {
    // stat failed (file vanished mid-read?): nothing left to preserve
    return false
  }
}

/**
 * String spellings of the two booleans that typed-but-sloppy external
 * editors write into app-settings.json (BUG-1773). Matched case- and
 * whitespace-insensitively; "0"/"1" included because number-in-a-string is
 * the same class of mistake.
 */
const BOOLEAN_FALSE_STRINGS = new Set(['', 'no', 'false', 'off', '0'])
const BOOLEAN_TRUE_STRINGS = new Set(['yes', 'true', 'on', '1'])

/**
 * Coerce a stored settings value to its schema type (boolean) BEFORE the
 * caller decides anything (BUG-1773, SET-26-3): readers used to answer
 * `value !== false`, so a string `"no"` written by an external editor — or
 * any non-false junk — silently counted as "enabled". Recognized boolean
 * spellings win; real booleans pass through; every other value (typed junk
 * like `42`, `{}` or `"banana"`, or an absent key) falls back to the key's
 * schema default passed as `fallback`.
 */
export function normalizeBooleanSetting(value: unknown, fallback: boolean): boolean {
  if (typeof value === 'boolean') return value
  if (typeof value === 'string') {
    const text = value.trim().toLowerCase()
    if (BOOLEAN_FALSE_STRINGS.has(text)) return false
    if (BOOLEAN_TRUE_STRINGS.has(text)) return true
  }
  return fallback
}

const RECOVERY_MODE_LABELS: Record<CorruptRecoveryMode, string> = {
  bom: 'BOM-prefixed',
  trailing: 'garbage-trailed',
  truncated: 'truncated',
}

export function readAppSettingsFile(settingsPath: string): AppSettings {
  let raw: string
  try {
    raw = readFileSync(settingsPath, 'utf8')
  } catch {
    // missing or unreadable file: treat as empty settings
    return {}
  }
  const direct = parseSettingsText(raw)
  if (direct) {
    // healthy file: re-arm the forensic backup for a possible future corruption
    preservedCorruptStates.delete(settingsPath)
    return direct
  }
  // BUG-1771: these bytes would be destroyed by the very next merge-write.
  // Preserve them once, then answer with whatever strict parsing can salvage
  // so the merge-write restores recoverable keys instead of dropping them.
  const fresh = preserveCorruptFile(settingsPath)
  const recovery = recoverCorruptAppSettings(raw)
  if (recovery) {
    if (fresh) {
      console.warn(
        `[app-settings] recovered ${Object.keys(recovery.settings).length} settings key(s) from a ${RECOVERY_MODE_LABELS[recovery.mode]} file; they persist on the next write`,
      )
    }
    return recovery.settings
  }
  if (fresh) {
    console.warn('[app-settings] settings file corrupt beyond repair; defaults apply')
  }
  return {}
}

/**
 * The canonical write: merge `updates` into a FRESH read and replace the
 * file atomically. The temporary file lives beside the destination so the
 * rename never crosses filesystems, and a crash mid-write can never
 * truncate settings — readers see either the old or the new object.
 * Wholesale rewrites are impossible by construction: this is the only
 * writer, and it always merges. When the on-disk file is corrupt, the fresh
 * read has already preserved those bytes as `.bak` (BUG-1771), so the merge
 * can only ever replace preserved, recoverable content.
 */
export function writeAppSettingsFile(settingsPath: string, updates: AppSettings): void {
  const settings = { ...readAppSettingsFile(settingsPath), ...updates }
  const tempPath = `${settingsPath}.${process.pid}.${randomUUID()}.tmp`
  try {
    writeFileSync(tempPath, JSON.stringify(settings, null, 2), {
      encoding: 'utf8',
      flag: 'wx',
      flush: true,
    })
    renameSync(tempPath, settingsPath)
  } catch (error) {
    try {
      unlinkSync(tempPath)
    } catch {
      // The write may have failed before the temporary file was created.
    }
    throw error
  }
}

/** per-settings-file write queues: at most one queued write in flight each */
const writeQueues = new Map<string, Promise<void>>()

/**
 * Serialized read-modify-write for async callers. Queued writes for the
 * same path run strictly one at a time (promise chain); a reducer's
 * `current` is read inside its turn, so updates computed earlier can never
 * clobber keys written in between. A failing write rejects its own promise
 * but never poisons the chain — the next queued write still runs.
 */
export function queueAppSettingsUpdate(
  settingsPath: string,
  update: AppSettingsUpdate,
): Promise<void> {
  const previous = writeQueues.get(settingsPath) ?? Promise.resolve()
  const run = previous.then(() => {
    if (typeof update === 'function') {
      const current = readAppSettingsFile(settingsPath)
      writeAppSettingsFile(settingsPath, update(current))
    } else {
      writeAppSettingsFile(settingsPath, update)
    }
  })
  // the chain continues past failures; only `run`'s caller sees the error
  const tail = run.catch(() => {})
  writeQueues.set(settingsPath, tail)
  // drop the queue entry once idle so a long-lived main leaks nothing
  void tail.then(() => {
    if (writeQueues.get(settingsPath) === tail) writeQueues.delete(settingsPath)
  })
  return run
}
