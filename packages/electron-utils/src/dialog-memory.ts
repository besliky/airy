/// Electron 43 pins file dialogs without an explicit `defaultPath` to the
/// user's Downloads folder and no longer lets the OS restore the last-used
/// directory between invocations (electron/electron#49868). These wrappers
/// restore the pre-43 behavior the way the Electron breaking-changes guide
/// recommends: remember the directory of the last confirmed pick and thread
/// it into the next dialog's `defaultPath`. The session memory is a WeakMap
/// keyed by the dialog module (tests with fake dialogs stay isolated; each
/// app process has its own), seeded from and written back to a persisted
/// LRU in userData/app-settings.json (`lastDialogDirs`) so the memory also
/// survives relaunches.
import { randomUUID } from 'node:crypto'
import { readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs'
import { basename, dirname, join } from 'node:path'

import { grantRendererDir, grantRendererFileAccess } from './renderer-file-access'

import type {
  BrowserWindow,
  Dialog,
  OpenDialogOptions,
  OpenDialogReturnValue,
  SaveDialogOptions,
  SaveDialogReturnValue,
} from 'electron'

// Keyed by the dialog module so tests with fake dialogs stay isolated. Each
// app bundles its own copy of this module, so the memory is scoped per
// editor — close enough to the per-app directory tracking the OS did before.
const lastUsedDirectoryByDialog = new WeakMap<Dialog, string>()

// ── persisted last-dialog directories (app-settings.json `lastDialogDirs`) ──

/** one remembered directory per app scope, most recent first */
export interface DialogDirEntry {
  scope: string
  dir: string
}

const LAST_DIALOG_DIRS_KEY = 'lastDialogDirs'
const LAST_DIALOG_DIRS_CAP = 10

/**
 * Record a pick: move the scope's entry to the front, cap the list. Pure —
 * unit-tested; the persisted JSON keeps this order.
 */
export function recordDialogDir(
  entries: DialogDirEntry[],
  scope: string,
  dir: string,
  cap = LAST_DIALOG_DIRS_CAP,
): DialogDirEntry[] {
  const rest = entries.filter((entry) => entry.scope !== scope)
  return [{ scope, dir }, ...rest].slice(0, cap)
}

function readAppSettings(settingsPath: string): Record<string, unknown> {
  try {
    const raw: unknown = JSON.parse(readFileSync(settingsPath, 'utf8'))
    if (raw && typeof raw === 'object' && !Array.isArray(raw)) {
      return raw as Record<string, unknown>
    }
  } catch {
    // missing or corrupt file: treat as empty settings
  }
  return {}
}

/** parse the persisted LRU; malformed values yield an empty list */
export function readLastDialogDirs(settingsPath: string): DialogDirEntry[] {
  const raw = readAppSettings(settingsPath)[LAST_DIALOG_DIRS_KEY]
  if (!Array.isArray(raw)) return []
  const entries: DialogDirEntry[] = []
  for (const item of raw) {
    if (
      item &&
      typeof item === 'object' &&
      !Array.isArray(item) &&
      typeof (item as DialogDirEntry).scope === 'string' &&
      typeof (item as DialogDirEntry).dir === 'string' &&
      (item as DialogDirEntry).scope &&
      (item as DialogDirEntry).dir
    ) {
      entries.push({ scope: (item as DialogDirEntry).scope, dir: (item as DialogDirEntry).dir })
    }
  }
  return entries
}

/** read-merge-write with the same atomic temp+rename the shell uses */
export function writeLastDialogDir(settingsPath: string, scope: string, dir: string): void {
  const settings = readAppSettings(settingsPath)
  settings[LAST_DIALOG_DIRS_KEY] = recordDialogDir(readLastDialogDirs(settingsPath), scope, dir)
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
    // persistence is best-effort; the session WeakMap still works
  }
}

/** minimal electron app surface the persistence glue needs */
interface ElectronAppLike {
  getPath?(name: string): string
  getAppPath?(): string
}

let electronApp: ElectronAppLike | null | undefined

/** lazily resolve electron's app; null outside a real main process (tests) */
async function resolveElectronApp(): Promise<ElectronAppLike | null> {
  if (electronApp !== undefined) return electronApp
  try {
    const electron = (await import('electron')) as { app?: ElectronAppLike }
    electronApp = electron.app ?? null
  } catch {
    electronApp = null
  }
  return electronApp
}

/**
 * App scope for the persisted entry: the running app's directory name
 * (shell / docs / sheets…). In a packaged build every editor shares the
 * shell's single dialog module anyway, so one bucket matches reality.
 */
async function dialogScope(): Promise<string> {
  const app = await resolveElectronApp()
  try {
    return basename(app?.getAppPath?.() ?? 'app')
  } catch {
    return 'app'
  }
}

/** seed the session WeakMap from the persisted LRU (once per dialog) */
async function seedPersistedDirectory(dialog: Dialog): Promise<void> {
  if (lastUsedDirectoryByDialog.has(dialog)) return
  lastUsedDirectoryByDialog.set(dialog, '') // reserve while the read is in flight
  const app = await resolveElectronApp()
  if (!app?.getPath) {
    lastUsedDirectoryByDialog.delete(dialog)
    return
  }
  const settingsPath = join(app.getPath('userData'), 'app-settings.json')
  const scope = await dialogScope()
  const found = readLastDialogDirs(settingsPath).find((entry) => entry.scope === scope)
  if (found) lastUsedDirectoryByDialog.set(dialog, found.dir)
  else lastUsedDirectoryByDialog.delete(dialog)
}

/** write a confirmed pick back to the persisted LRU (best-effort, async) */
async function persistPickedDirectory(dir: string): Promise<void> {
  const app = await resolveElectronApp()
  if (!app?.getPath) return
  writeLastDialogDir(join(app.getPath('userData'), 'app-settings.json'), await dialogScope(), dir)
}

function withRememberedDirectory<T extends OpenDialogOptions | SaveDialogOptions>(
  dialog: Dialog,
  options: T,
  fallbackDir?: string,
): T {
  // No pick confirmed yet this session: anchor in the caller's fallback (the
  // configurable default save folder) instead of Electron's Downloads pin.
  const lastDir = lastUsedDirectoryByDialog.get(dialog) ?? fallbackDir
  if (!lastDir) return options
  if (options.defaultPath === undefined) return { ...options, defaultPath: lastDir }
  // A bare file name is a Save As suggestion: anchor it in the remembered
  // directory instead of letting Electron relocate it to Downloads.
  if (basename(options.defaultPath) === options.defaultPath) {
    return { ...options, defaultPath: join(lastDir, options.defaultPath) }
  }
  return options
}

export async function showOpenDialogWithMemory(
  dialog: Dialog,
  parent: BrowserWindow | null | undefined,
  options: OpenDialogOptions,
  fallbackDir?: string,
): Promise<OpenDialogReturnValue> {
  await seedPersistedDirectory(dialog)
  const withDir = withRememberedDirectory(dialog, options, fallbackDir)
  const result = parent
    ? await dialog.showOpenDialog(parent, withDir)
    : await dialog.showOpenDialog(withDir)
  const picked = result.filePaths[0]
  if (!result.canceled && picked) {
    const pickedDir = options.properties?.includes('openDirectory') ? picked : dirname(picked)
    lastUsedDirectoryByDialog.set(dialog, pickedDir)
    void persistPickedDirectory(pickedDir)
    // Every dialog pick is a user-driven folder choice: allow the renderers
    // to read what the user just picked (files and chosen directories alike).
    for (const filePath of result.filePaths) grantRendererFileAccess(filePath)
    if (options.properties?.includes('openDirectory')) grantRendererDir(picked)
  }
  return result
}

/**
 * Save As suggestion for a document that was opened from `sourcePath`: the
 * same folder with the suggested name (Word parity — Save As starts where the
 * document lives, not in the last-used or default folder). Falls back to the
 * bare name, which `withRememberedDirectory` then anchors, when the document
 * has never been on disk.
 */
export function saveAsSuggestion(
  sourcePath: string | null | undefined,
  defaultName: string,
): string {
  return sourcePath ? join(dirname(sourcePath), defaultName) : defaultName
}

export async function showSaveDialogWithMemory(
  dialog: Dialog,
  parent: BrowserWindow | null | undefined,
  options: SaveDialogOptions,
  fallbackDir?: string,
): Promise<SaveDialogReturnValue> {
  await seedPersistedDirectory(dialog)
  const withDir = withRememberedDirectory(dialog, options, fallbackDir)
  const result = parent
    ? await dialog.showSaveDialog(parent, withDir)
    : await dialog.showSaveDialog(withDir)
  if (!result.canceled && result.filePath) {
    lastUsedDirectoryByDialog.set(dialog, dirname(result.filePath))
    void persistPickedDirectory(dirname(result.filePath))
    // the user just chose this folder for a save: reading there is fine too
    grantRendererFileAccess(result.filePath)
  }
  return result
}
