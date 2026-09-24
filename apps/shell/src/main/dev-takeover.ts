/**
 * Helpers for the dev-mode single-instance takeover in the shell entry
 * (see the '!hasLock && !app.isPackaged' branch in src/main/index.ts).
 *
 * BUG-1680: the zombie check used `cmd.includes('Electron')`, which is both
 * case-sensitive and matches anywhere in the command line. A linux dev
 * checkout runs the lowercase binary `node_modules/electron/dist/electron`,
 * so a wedged previous instance kept the SingletonLock, the takeover refused
 * to kill it (the pid-recycling guard read the miss as "not Electron"), and
 * every later dev launch on that userData died with "Target closed".
 */

/**
 * True when a `ps -o command=` line belongs to an Electron executable.
 *
 * The decision is made on the EXECUTABLE only — the first whitespace-separated
 * token — by comparing its basename (either `/` or `\` separators, so windows
 * command lines work too) against `electron`, case-insensitively, after
 * stripping a windows-style `.exe` suffix. This matches every dev layout:
 * linux `…/dist/electron`, macOS `…/Electron.app/Contents/MacOS/Electron`,
 * windows `…\dist\Electron.exe` — and deliberately rejects foreign processes
 * that merely mention "electron" in an argument (e.g. `vim electron-notes.md`),
 * so pid recycling can never talk us into killing an innocent process.
 */
export function isElectronProcessCommand(commandLine: string): boolean {
  const executable = commandLine.trim().split(/\s+/)[0] ?? ''
  if (!executable) return false
  const basename = executable.split(/[\\/]/).pop() ?? ''
  const stem = basename.toLowerCase().replace(/\.exe$/, '')
  return stem === 'electron'
}
