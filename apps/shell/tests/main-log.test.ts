import { mkdtempSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { inspect } from 'node:util'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { createMainLog, teeConsoleInto, type MainLog } from '../src/main/main-log'

/**
 * Main-process ring log (src/main/main-log.ts, UX-1628): bounded memory,
 * size-capped rotated files under a caller-chosen directory, and a hard
 * never-throw contract — the logger sits on the app's crash path, so a
 * failing filesystem (read-only volume, vanished directory) may degrade
 * persistence but must never break the caller.
 */

let dir: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'airy-main-log-'))
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

function activeFile(): string {
  return join(dir, 'main.log')
}

describe('createMainLog — writing', () => {
  it('persists one line per entry with timestamp, level and message', () => {
    const log = createMainLog({ dir })
    log.info('hello world')
    const content = readFileSync(activeFile(), 'utf-8')
    expect(content).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}.\d{3}Z INFO hello world\n$/)
  })

  it('creates the log directory when missing', () => {
    const nested = join(dir, 'logs')
    const log = createMainLog({ dir: nested })
    log.warn('late line')
    expect(readFileSync(join(nested, 'main.log'), 'utf-8')).toContain('late line')
  })

  it('flattens multi-line messages into a single line', () => {
    const log = createMainLog({ dir })
    log.error('line one\nline two\r\nline three')
    const lines = readFileSync(activeFile(), 'utf-8').trim().split('\n')
    expect(lines).toHaveLength(1)
    expect(lines[0]).toContain('line one line two line three')
  })

  it('truncates lines beyond maxLineLength but keeps the head', () => {
    const log = createMainLog({ dir, maxLineLength: 64 })
    log.info('x'.repeat(300))
    const line = readFileSync(activeFile(), 'utf-8').trim()
    expect(line.length).toBeLessThan(120)
    expect(line).toContain('xxxx')
    expect(line).toContain('(+')
  })

  it('event() renders a structured single line with JSON fields', () => {
    const log = createMainLog({ dir })
    log.event('render-process-gone', { wc: 7, reason: 'oom', exitCode: -1 })
    const line = readFileSync(activeFile(), 'utf-8').trim()
    expect(line).toContain(' WARN event render-process-gone')
    expect(line).toContain('wc=7')
    expect(line).toContain('reason="oom"')
    expect(line).toContain('exitCode=-1')
  })

  it('event() tolerates unserializable and circular field values', () => {
    const log = createMainLog({ dir })
    const circular: Record<string, unknown> = {}
    circular.self = circular
    log.event('uncaught-exception', { fn: () => 1, circular })
    const line = readFileSync(activeFile(), 'utf-8').trim()
    expect(line).toContain('[unserializable]')
  })

  it('the ring keeps only the newest ringCapacity lines, oldest first', () => {
    const log = createMainLog({ dir, ringCapacity: 3 })
    for (let i = 1; i <= 5; i++) log.info(`line ${i}`)
    const recent = log.recentLines()
    expect(recent).toHaveLength(3)
    expect(recent[0]).toContain('line 3')
    expect(recent[2]).toContain('line 5')
    expect(log.stats().ringSize).toBe(3)
  })
})

describe('createMainLog — rotation and size cap', () => {
  it('rotates the active file to main.log.1 when the cap is exceeded', () => {
    const log = createMainLog({ dir, maxFileBytes: 200, maxBackups: 2 })
    for (let i = 0; i < 20; i++) log.info(`rotation probe line ${i}`)
    const backup = readFileSync(join(dir, 'main.log.1'), 'utf-8')
    expect(backup).toContain('rotation probe line')
    // the active file never exceeds the cap by more than its newest line
    expect(statSync(activeFile()).size).toBeLessThanOrEqual(200 + 100)
    expect(log.stats().rotations).toBeGreaterThanOrEqual(1)
  })

  it('keeps at most maxBackups rotated files plus the active one', () => {
    const log = createMainLog({ dir, maxFileBytes: 120, maxBackups: 2 })
    for (let i = 0; i < 60; i++) log.info(`cap probe line ${i}`)
    const files = readdirSync(dir).filter((f) => f.startsWith('main.log'))
    expect(files).toHaveLength(3)
    expect(files).toContain('main.log')
    expect(files).toContain('main.log.1')
    expect(files).toContain('main.log.2')
  })

  it('bounds total on-disk size by (maxBackups + 1) * maxFileBytes', () => {
    const maxFileBytes = 150
    const maxBackups = 2
    const log = createMainLog({ dir, maxFileBytes, maxBackups })
    for (let i = 0; i < 200; i++) log.info(`disk bound probe ${i}`)
    const total = readdirSync(dir)
      .filter((f) => f.startsWith('main.log'))
      .reduce((sum, f) => sum + statSync(join(dir, f)).size, 0)
    // each line here fits the cap, so rotation fires before the write and no
    // file can exceed maxFileBytes at all
    expect(total).toBeLessThanOrEqual((maxBackups + 1) * maxFileBytes)
    expect(log.stats().droppedWrites).toBe(0)
  })

  it('rotates an already oversized file on init, before the first write', () => {
    writeFileSync(activeFile(), 'z'.repeat(500))
    const log = createMainLog({ dir, maxFileBytes: 100, maxBackups: 2 })
    log.info('fresh start')
    expect(readFileSync(join(dir, 'main.log.1'), 'utf-8')).toBe('z'.repeat(500))
    expect(readFileSync(activeFile(), 'utf-8')).toContain('fresh start')
    expect(log.stats().rotations).toBe(1)
  })
})

describe('createMainLog — failure containment', () => {
  it('never throws when the log directory cannot exist (parent is a file)', () => {
    const blocker = join(dir, 'blocker')
    writeFileSync(blocker, 'not a directory')
    const log = createMainLog({ dir: join(blocker, 'logs') })
    expect(() => {
      log.info('into the void')
      log.event('fatal-event', {})
    }).not.toThrow()
    expect(log.stats().droppedWrites).toBeGreaterThan(0)
    // the in-memory ring still works for diagnostics
    expect(log.recentLines()).toHaveLength(2)
  })

  it('never throws when a custom inspector throws on an argument', () => {
    const log = createMainLog({ dir })
    const evil = {
      [inspect.custom]: () => {
        throw new Error('no')
      },
    }
    expect(() => log.info('evil:', evil)).not.toThrow()
    expect(readFileSync(activeFile(), 'utf-8')).toContain('evil:')
  })

  it('a throwing log.line inside the console tee cannot recurse forever', () => {
    let consoleErrorCalls = 0
    const errorSpy = vi.spyOn(console, 'error').mockImplementation(() => {
      consoleErrorCalls++
    })
    let lineCalls = 0
    const hostile = {
      line: () => {
        lineCalls++
        if (lineCalls < 10) console.error('echo')
      },
    } as unknown as MainLog
    const restore = teeConsoleInto(hostile)
    expect(() => console.error('outer')).not.toThrow()
    // outer call + exactly one mirrored echo; the guard stops the loop (an
    // unguarded tee would ping-pong until the stack blows)
    expect(consoleErrorCalls).toBe(2)
    expect(lineCalls).toBe(1)
    restore()
    errorSpy.mockRestore()
  })
})

describe('teeConsoleInto', () => {
  it('mirrors console output into the log while stderr keeps working', () => {
    const log = createMainLog({ dir })
    const printed: string[] = []
    const spies = [
      vi
        .spyOn(console, 'log')
        .mockImplementation((...args: unknown[]) => printed.push(args.join(' '))),
      vi
        .spyOn(console, 'warn')
        .mockImplementation((...args: unknown[]) => printed.push(args.join(' '))),
      vi
        .spyOn(console, 'error')
        .mockImplementation((...args: unknown[]) => printed.push(args.join(' '))),
    ]
    const restore = teeConsoleInto(log)
    try {
      console.log('info line', 42)
      console.warn('warn line')
      console.error('error line')
    } finally {
      restore()
      for (const spy of spies) spy.mockRestore()
    }
    const persisted = readFileSync(activeFile(), 'utf-8')
    expect(persisted).toContain('INFO info line 42')
    expect(persisted).toContain('WARN warn line')
    expect(persisted).toContain('ERROR error line')
    // stderr parity: every call still reached the original sink
    expect(printed).toHaveLength(3)
    // after restore, console calls no longer land in the log
    console.error('post-restore')
    expect(readFileSync(activeFile(), 'utf-8')).not.toContain('post-restore')
  })
})
