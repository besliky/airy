import { describe, expect, it } from 'vitest'
import { isElectronProcessCommand } from '../src/main/dev-takeover'

/**
 * Regression tests for the dev-takeover pid-recycling guard (BUG-1680): the
 * old check `cmd.includes('Electron')` was case-sensitive and scanned the
 * whole command line, so the lowercase linux dev binary `…/dist/electron`
 * was never recognized as Electron — a wedged previous instance kept the
 * SingletonLock and every later dev launch died with "Target closed". The
 * matcher must accept an Electron executable by basename on every platform
 * layout, and must reject foreign processes that merely mention "electron"
 * in their arguments (pid recycling must not kill innocents).
 */

describe('isElectronProcessCommand', () => {
  it('matches the lowercase linux dev binary (BUG-1680)', () => {
    expect(isElectronProcessCommand('/home/dev/airy/node_modules/electron/dist/electron')).toBe(
      true,
    )
    expect(
      isElectronProcessCommand('/home/dev/airy/node_modules/electron/dist/electron --no-sandbox .'),
    ).toBe(true)
  })

  it('matches the macOS Electron.app binary', () => {
    expect(
      isElectronProcessCommand(
        '/Users/dev/airy/node_modules/electron/dist/Electron.app/Contents/MacOS/Electron',
      ),
    ).toBe(true)
    expect(
      isElectronProcessCommand(
        '/Users/dev/airy/node_modules/electron/dist/Electron.app/Contents/MacOS/Electron .',
      ),
    ).toBe(true)
  })

  it('matches the windows Electron.exe binary (backslash separators)', () => {
    expect(
      isElectronProcessCommand('C:\\dev\\airy\\node_modules\\electron\\dist\\Electron.exe'),
    ).toBe(true)
    expect(
      isElectronProcessCommand('C:\\dev\\airy\\node_modules\\electron\\dist\\Electron.exe .'),
    ).toBe(true)
    // lowercase .exe suffix and mixed-case binary name stay case-insensitive
    expect(isElectronProcessCommand('D:\\Apps\\airy-shell\\ELECTRON.exe')).toBe(true)
  })

  it('matches mixed-case paths and a bare executable name', () => {
    expect(isElectronProcessCommand('/opt/ElectronDist/electron')).toBe(true)
    expect(isElectronProcessCommand('/opt/ElectronDist/ELECTRON --type=zygote')).toBe(true)
    expect(isElectronProcessCommand('electron')).toBe(true)
    // ps -o command= pads its output; leading whitespace must not break the match
    expect(isElectronProcessCommand('  /home/dev/airy/node_modules/electron/dist/electron .')).toBe(
      true,
    )
  })

  it('rejects foreign processes, including ones mentioning electron in arguments', () => {
    expect(isElectronProcessCommand('node server.js')).toBe(false)
    expect(isElectronProcessCommand('/usr/bin/node /home/dev/airy/scripts/dev.mjs')).toBe(false)
    expect(isElectronProcessCommand('bash')).toBe(false)
    // an editor or pager opening a file named "electron*" is not an Electron process
    expect(isElectronProcessCommand('/usr/bin/vim electron-notes.md')).toBe(false)
    expect(isElectronProcessCommand('tail -f /tmp/electron.log')).toBe(false)
    expect(isElectronProcessCommand('grep -rn Electron src/main')).toBe(false)
    expect(isElectronProcessCommand('code /home/dev/electron')).toBe(false)
    // a shell wrapper launching electron is not itself the Electron binary
    expect(isElectronProcessCommand('sh -c electron .')).toBe(false)
  })

  it('rejects empty and blank command lines', () => {
    expect(isElectronProcessCommand('')).toBe(false)
    expect(isElectronProcessCommand('   ')).toBe(false)
    expect(isElectronProcessCommand('\n')).toBe(false)
  })
})
