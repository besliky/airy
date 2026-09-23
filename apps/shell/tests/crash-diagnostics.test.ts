import { EventEmitter } from 'node:events'
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { MainLog } from '../src/main/main-log'

/**
 * Local-only crash diagnostics (src/main/crash-diagnostics.ts, UX-1628):
 * the reporter must never be given a submitURL or uploads (minidumps stay
 * under userData/crash-dumps), the fatal process events must land in the
 * ring log, and none of the wiring may throw — it runs before the heavy
 * startup phases and sits on the crash path itself.
 */

const setPath = vi.fn()
const crashStart = vi.fn()
const fakeApp = Object.assign(new EventEmitter(), {
  setPath,
  getPath: vi.fn(() => '/fake-user-data'),
  getVersion: vi.fn(() => '0.15.0-test'),
})

vi.mock('electron', () => ({
  app: fakeApp,
  crashReporter: { start: crashStart },
}))

function capturingLog(): MainLog & { events: string[] } {
  const events: string[] = []
  const noop = (): void => {}
  return {
    line: noop,
    event: (name, fields) => {
      events.push(`event ${name} ${JSON.stringify(fields ?? {})}`)
    },
    debug: noop,
    info: noop,
    warn: noop,
    error: noop,
    recentLines: () => [],
    paths: () => ({ dir: '', activeFile: '' }),
    stats: () => ({ ringSize: 0, activeFileBytes: 0, rotations: 0, droppedWrites: 0 }),
    events,
  }
}

function startOptions(): Record<string, unknown> {
  expect(crashStart).toHaveBeenCalledTimes(1)
  return crashStart.mock.calls[0][0] as Record<string, unknown>
}

let userDataDir: string

beforeEach(() => {
  vi.resetModules()
  userDataDir = mkdtempSync(join(tmpdir(), 'airy-crash-diag-'))
  setPath.mockClear()
  crashStart.mockReset()
})

afterEach(() => {
  rmSync(userDataDir, { recursive: true, force: true })
  fakeApp.removeAllListeners()
})

describe('startLocalCrashReporter', () => {
  it('starts the reporter with uploads disabled and no submitURL', async () => {
    const { startLocalCrashReporter } = await import('../src/main/crash-diagnostics')
    startLocalCrashReporter({
      app: fakeApp,
      crashReporter: { start: crashStart },
      userDataDir,
    })
    const options = startOptions()
    expect(options.uploadToServer).toBe(false)
    // a submitURL would be the only way reports could leave the machine
    expect('submitURL' in options).toBe(false)
  })

  it('pins crashDumps under userData and creates the dumps directory', async () => {
    const { startLocalCrashReporter, CRASH_DUMPS_DIR_NAME } =
      await import('../src/main/crash-diagnostics')
    const { dumpsDir } = startLocalCrashReporter({
      app: fakeApp,
      crashReporter: { start: crashStart },
      userDataDir,
    })
    expect(setPath).toHaveBeenCalledWith('crashDumps', join(userDataDir, CRASH_DUMPS_DIR_NAME))
    expect(dumpsDir).toBe(join(userDataDir, 'crash-dumps'))
    expect(existsSync(dumpsDir)).toBe(true)
  })

  it('survives a throwing crashReporter.start and a throwing setPath', async () => {
    const { startLocalCrashReporter } = await import('../src/main/crash-diagnostics')
    setPath.mockImplementation(() => {
      throw new Error('frozen paths')
    })
    crashStart.mockImplementation(() => {
      throw new Error('no crashpad here')
    })
    expect(() =>
      startLocalCrashReporter({
        app: fakeApp,
        crashReporter: { start: crashStart },
        userDataDir,
      }),
    ).not.toThrow()
  })
})

describe('installFatalEventLogging', () => {
  it('logs render-process-gone from a mocked app event into the log', async () => {
    const { installFatalEventLogging } = await import('../src/main/crash-diagnostics')
    const log = capturingLog()
    installFatalEventLogging({ app: fakeApp, log })
    // mock event + mock webContents, as the app-level event delivers them
    fakeApp.emit(
      'render-process-gone',
      { preventDefault: () => {} },
      { id: 7 },
      {
        reason: 'oom',
        exitCode: -1,
      },
    )
    expect(log.events).toEqual([expect.stringContaining('"wc":7')])
    expect(log.events[0]).toContain('event render-process-gone')
    expect(log.events[0]).toContain('"reason":"oom"')
    expect(log.events[0]).toContain('"exitCode":-1')
  })

  it('logs child-process-gone with type, reason and exit code', async () => {
    const { installFatalEventLogging } = await import('../src/main/crash-diagnostics')
    const log = capturingLog()
    installFatalEventLogging({ app: fakeApp, log })
    fakeApp.emit(
      'child-process-gone',
      { preventDefault: () => {} },
      { type: 'GPU', name: 'GPU Process', reason: 'crashed', exitCode: 11 },
    )
    expect(log.events).toHaveLength(1)
    expect(log.events[0]).toContain('event child-process-gone')
    expect(log.events[0]).toContain('"type":"GPU"')
    expect(log.events[0]).toContain('"reason":"crashed"')
  })

  it('marks the graceful quit boundary with an app-quit event', async () => {
    const { installFatalEventLogging } = await import('../src/main/crash-diagnostics')
    const log = capturingLog()
    installFatalEventLogging({ app: fakeApp, log })
    fakeApp.emit('will-quit')
    expect(log.events).toEqual(['event app-quit {}'])
  })

  it('tolerates missing event details (undefined webContents/details)', async () => {
    const { installFatalEventLogging } = await import('../src/main/crash-diagnostics')
    const log = capturingLog()
    installFatalEventLogging({ app: fakeApp, log })
    fakeApp.emit('render-process-gone')
    // all fields undefined → JSON.stringify drops them, the event survives
    expect(log.events).toEqual(['event render-process-gone {}'])
  })
})

describe('initShellDiagnostics', () => {
  it('wires log dir + crash reporter + fatal events in one guarded call', async () => {
    const { initShellDiagnostics, CRASH_DUMPS_DIR_NAME, LOGS_DIR_NAME } =
      await import('../src/main/crash-diagnostics')
    const diagnostics = initShellDiagnostics({ userDataDir })
    expect(diagnostics.dumpsDir).toBe(join(userDataDir, CRASH_DUMPS_DIR_NAME))
    expect(existsSync(join(userDataDir, LOGS_DIR_NAME))).toBe(true)
    // the startup line records the local-only contract for the user's report
    diagnostics.log.info('startup probe')
    const content = readFileSync(join(userDataDir, LOGS_DIR_NAME, 'main.log'), 'utf-8')
    expect(content).toContain('event crash-diagnostics-start')
    expect(content).toContain('uploadToServer=false')
    expect(content).toContain('startup probe')
  })

  it('is idempotent: a second init returns the same diagnostics', async () => {
    const { initShellDiagnostics } = await import('../src/main/crash-diagnostics')
    const first = initShellDiagnostics({ userDataDir })
    const second = initShellDiagnostics({ userDataDir })
    expect(second).toBe(first)
    expect(crashStart).toHaveBeenCalledTimes(1)
  })
})
