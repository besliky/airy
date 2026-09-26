/**
 * userData/app-settings.json access for the shell main process: a flat JSON
 * object holding cross-module app preferences (UI language, first-run
 * onboarding flag). The read/merge/atomic-rename implementation and the
 * async write queue live in @airy-office/electron-utils (single writer
 * shared with the editor modules — see app-settings-file.ts there, and
 * OBS-1532 for why the duplication was removed); these wrappers keep the
 * shell's synchronous call sites unchanged. Editor modules read the same
 * file, so the shape must stay a plain object.
 *
 * Corrupt files (BUG-1771): the shared implementation preserves the corrupt
 * bytes as `app-settings.json.bak` before any rewrite and salvages what
 * strict JSON parsing allows (BOM / trailing garbage / truncated tail), so a
 * merge-write restores recoverable keys instead of silently dropping every
 * setting. These wrappers inherit all of it unchanged.
 */
import {
  readAppSettingsFile,
  writeAppSettingsFile,
  type AppSettings,
} from '@airy-office/electron-utils'

export type { AppSettings }

export const readAppSettings = readAppSettingsFile
export const writeAppSettings = writeAppSettingsFile

export function writeAppSetting(path: string, key: string, value: unknown): void {
  writeAppSettings(path, { [key]: value })
}
