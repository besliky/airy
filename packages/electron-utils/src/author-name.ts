/// The user's display name for authorship metadata: new comments, tracked-
/// change revision marks, and the slides comment author table. Stored under
/// `authorName` in userData/app-settings.json (set from the shell's
/// Settings > General); every editor main resolves through here so they all
/// honor the same setting. An unset/empty value means "use the caller's
/// fallback" (localized default author or system username).
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

/** the subset of Electron's `app` needed here (kept structural: this package has no Electron dependency) */
export interface AuthorNamePathProvider {
  getPath(name: 'userData'): string
}

export const AUTHOR_NAME_KEY = 'authorName'

/** display names are short free text; PowerPoint/Word UI effectively caps around this length */
export const AUTHOR_NAME_MAX = 60

/**
 * Normalize a raw author name: drop control characters, collapse whitespace
 * runs, trim, and cap the length. Returns '' for anything unusable.
 */
export function sanitizeAuthorName(raw: unknown): string {
  if (typeof raw !== 'string') return ''
  // C0 controls and DEL become spaces (kept as separators, not deleted, so a
  // name like "Ada\u0007Lovelace" still reads as two words)
  const stripped = Array.from(raw, (ch) => {
    const code = ch.codePointAt(0) ?? 0
    return code < 0x20 || code === 0x7f ? ' ' : ch
  }).join('')
  return stripped.replace(/\s+/g, ' ').trim().slice(0, AUTHOR_NAME_MAX)
}

/** the configured author name from app-settings.json, or '' when unset/unreadable */
export function readAuthorNameSetting(settingsPath: string): string {
  try {
    const raw: unknown = JSON.parse(readFileSync(settingsPath, 'utf8'))
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return ''
    return sanitizeAuthorName((raw as Record<string, unknown>)[AUTHOR_NAME_KEY])
  } catch {
    return ''
  }
}

/** convenience for the Electron mains: settings lookup in one call ('' = unset) */
export function configuredAuthorName(app: AuthorNamePathProvider): string {
  return readAuthorNameSetting(join(app.getPath('userData'), 'app-settings.json'))
}
