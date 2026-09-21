import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

/**
 * Session persistence (src/main/session-state.ts): serialize the open-tab
 * set window by window (window-ordered, so a tab moved to a second window
 * restores both), parse/validate persisted JSON (including the legacy
 * single-window shape), prune entries whose files vanished, and round-trip
 * the session file atomically.
 */

let S: typeof import('../src/main/session-state')
let scratch: string

beforeEach(async () => {
  S = await import('../src/main/session-state')
  scratch = mkdtempSync(join(tmpdir(), 'airy-session-state-'))
})

afterEach(() => {
  rmSync(scratch, { recursive: true, force: true })
})

describe('serializeSessionWindow', () => {
  it('keeps file-backed editor tabs in strip order and skips untitled/home/present tabs', () => {
    const state = S.serializeSessionWindow(
      [
        { id: 'home', kind: 'home' as const },
        { id: 't1', kind: 'docs' as const, filePath: '/docs/a.docx' },
        { id: 't2', kind: 'sheets' as const }, // untitled — no backing file
        { id: 't3', kind: 'html' as const }, // present tab — no file
        { id: 't4', kind: 'pdf' as const, filePath: '/pdf/b.pdf' },
      ],
      't4',
    )
    expect(state).toEqual({
      tabs: [
        { kind: 'docs', path: '/docs/a.docx' },
        { kind: 'pdf', path: '/pdf/b.pdf' },
      ],
      activePath: '/pdf/b.pdf',
    })
  })

  it('drops the activePath when the active tab has no file', () => {
    const state = S.serializeSessionWindow([{ id: 't1', kind: 'markdown' as const }], 't1')
    expect(state.activePath).toBeNull()
  })

  it('records no activePath when no tab matches the active id', () => {
    const state = S.serializeSessionWindow(
      [{ id: 't1', kind: 'docs' as const, filePath: '/a.docx' }],
      't9',
    )
    expect(state.activePath).toBeNull()
  })
})

describe('serializeSession (window-ordered)', () => {
  it('serializes two windows in order with their own active tabs', () => {
    const state = S.serializeSession(
      [
        {
          tabs: [
            { id: 'home', kind: 'home' as const },
            { id: 't1', kind: 'docs' as const, filePath: '/a.docx' },
            { id: 't2', kind: 'pdf' as const, filePath: '/b.pdf' },
          ],
          activeId: 't1',
        },
        {
          tabs: [{ id: 't1', kind: 'sheets' as const, filePath: '/c.xlsx' }],
          activeId: 't1',
        },
      ],
      1,
    )
    expect(state).toEqual({
      windows: [
        {
          tabs: [
            { kind: 'docs', path: '/a.docx' },
            { kind: 'pdf', path: '/b.pdf' },
          ],
          activePath: '/a.docx',
        },
        { tabs: [{ kind: 'sheets', path: '/c.xlsx' }], activePath: '/c.xlsx' },
      ],
      focusedWindow: 1,
    })
  })

  it('clamps an out-of-range focused window index to a valid one', () => {
    expect(S.serializeSession([{ tabs: [], activeId: null }], 3).focusedWindow).toBe(0)
    expect(S.serializeSession([{ tabs: [], activeId: null }], -1).focusedWindow).toBe(0)
  })
})

describe('parseSession', () => {
  it('parses a saved multi-window session', () => {
    const state = S.parseSession({
      windows: [
        {
          tabs: [
            { kind: 'docs', path: '/a.docx' },
            { kind: 'markdown', path: '/b.md' },
          ],
          activePath: '/b.md',
        },
        { tabs: [{ kind: 'pdf', path: '/c.pdf' }], activePath: null },
      ],
      focusedWindow: 1,
    })
    expect(state.windows).toHaveLength(2)
    expect(state.windows[0]!.activePath).toBe('/b.md')
    expect(state.focusedWindow).toBe(1)
  })

  it('migrates the legacy single-window shape into one focused window', () => {
    const state = S.parseSession({
      tabs: [
        { kind: 'docs', path: '/a.docx' },
        { kind: 'markdown', path: '/b.md' },
      ],
      activePath: '/b.md',
    })
    expect(state).toEqual({
      windows: [
        {
          tabs: [
            { kind: 'docs', path: '/a.docx' },
            { kind: 'markdown', path: '/b.md' },
          ],
          activePath: '/b.md',
        },
      ],
      focusedWindow: 0,
    })
  })

  it('yields an empty state for malformed payloads', () => {
    for (const bad of [
      null,
      undefined,
      'x',
      42,
      [],
      {},
      { windows: 'nope' },
      { windows: [1, 2] },
    ]) {
      expect(S.parseSession(bad)).toEqual({ windows: [], focusedWindow: 0 })
    }
  })

  it('skips invalid entries without aborting the rest', () => {
    const state = S.parseSession({
      windows: [
        {
          tabs: [
            null,
            { kind: 'home', path: '/x' }, // home is not restorable
            { kind: 'docs' }, // missing path
            { kind: 'unknown-kind', path: '/y' },
            { kind: 'pdf', path: '' },
            { kind: 'sheets', path: '/ok.xlsx' },
          ],
        },
      ],
    })
    expect(state.windows).toEqual([
      { tabs: [{ kind: 'sheets', path: '/ok.xlsx' }], activePath: null },
    ])
  })

  it('drops windows left without a single valid tab', () => {
    const state = S.parseSession({
      windows: [
        { tabs: [{ kind: 'docs', path: '/a.docx' }] },
        { tabs: [{ kind: 'home', path: '/x' }] }, // nothing restorable
        'garbage',
        { tabs: [{ kind: 'pdf', path: '/c.pdf' }] },
      ],
      focusedWindow: 2,
    })
    expect(state.windows.map((w: { tabs: unknown[] }) => w.tabs)).toHaveLength(2)
    // focus slid onto a surviving window
    expect(state.focusedWindow).toBe(1)
  })

  it('keeps the same file in two windows (per-window dedup, BUG-1108)', () => {
    const state = S.parseSession({
      windows: [
        { tabs: [{ kind: 'docs', path: '/a.docx' }] },
        {
          tabs: [
            { kind: 'docs', path: '/a.docx' }, // the same file, second window — a supported layout
            { kind: 'pdf', path: '/c.pdf' },
          ],
          activePath: '/a.docx',
        },
      ],
    })
    expect(state.windows[0]!.tabs).toEqual([{ kind: 'docs', path: '/a.docx' }])
    expect(state.windows[1]!.tabs).toEqual([
      { kind: 'docs', path: '/a.docx' },
      { kind: 'pdf', path: '/c.pdf' },
    ])
    // the second window's own activePath still validates against its tabs
    expect(state.windows[1]!.activePath).toBe('/a.docx')
  })

  it('still drops duplicate paths within one window', () => {
    const state = S.parseSession({
      windows: [
        {
          tabs: [
            { kind: 'docs', path: '/a.docx' },
            { kind: 'docs', path: '/a.docx' }, // same window: would just re-activate
            { kind: 'pdf', path: '/c.pdf' },
          ],
        },
      ],
    })
    expect(state.windows[0]!.tabs).toEqual([
      { kind: 'docs', path: '/a.docx' },
      { kind: 'pdf', path: '/c.pdf' },
    ])
  })

  it('an activePath repeated in another window stays valid for its own window', () => {
    const state = S.parseSession({
      windows: [
        { tabs: [{ kind: 'docs', path: '/a.docx' }], activePath: '/a.docx' },
        { tabs: [{ kind: 'docs', path: '/a.docx' }], activePath: '/a.docx' },
      ],
    })
    expect(state.windows[0]!.activePath).toBe('/a.docx')
    expect(state.windows[1]!.activePath).toBe('/a.docx')
  })

  it('ignores an activePath that no listed tab owns', () => {
    const state = S.parseSession({
      windows: [{ tabs: [{ kind: 'docs', path: '/a.docx' }], activePath: '/gone.docx' }],
    })
    expect(state.windows[0]!.activePath).toBeNull()
  })

  it('prefers the multi-window shape when both keys are present (BUG-1224)', () => {
    // a corrupt or hand-made file carrying both `tabs` and `windows`: the
    // legacy branch used to run first and silently discard the multi-window
    // layout (the v2 writer never emits `tabs`, so `windows` is the richer
    // reading of the file)
    const state = S.parseSession({
      tabs: [{ kind: 'docs', path: '/legacy.docx' }],
      activePath: '/legacy.docx',
      windows: [
        { tabs: [{ kind: 'docs', path: '/one.docx' }], activePath: '/one.docx' },
        { tabs: [{ kind: 'pdf', path: '/two.pdf' }], activePath: null },
      ],
      focusedWindow: 1,
    })
    expect(state.windows.map((w) => w.tabs.map((t) => t.path))).toEqual([
      ['/one.docx'],
      ['/two.pdf'],
    ])
    expect(state.windows[0]!.activePath).toBe('/one.docx')
    expect(state.focusedWindow).toBe(1)
  })

  it('caps runaway tab lists across all windows', () => {
    const tabs = Array.from({ length: 200 }, (_, i) => ({ kind: 'docs', path: `/f${i}.docx` }))
    const twoWindows = S.parseSession({ windows: [{ tabs }, { tabs }] })
    expect(twoWindows.windows[0]!.tabs).toHaveLength(64)
    expect(twoWindows.windows).toHaveLength(1) // the second window hit the shared cap
  })

  it('caps the GLOBAL total at 64, not 127 (BUG-1109 audit tail)', () => {
    // a first window with 63 valid tabs used to let the second add 64 of
    // its own — the per-window parser did not know the global budget, so
    // the declared cap of 64 actually admitted 127
    const w1 = Array.from({ length: 63 }, (_, i) => ({ kind: 'docs', path: `/a${i}.docx` }))
    const w2 = Array.from({ length: 80 }, (_, i) => ({ kind: 'pdf', path: `/b${i}.pdf` }))
    const state = S.parseSession({ windows: [{ tabs: w1 }, { tabs: w2 }] })
    const total = state.windows.reduce((n, w) => n + w.tabs.length, 0)
    expect(total).toBe(64)
    expect(state.windows[1]!.tabs).toHaveLength(1) // 64 - 63 leftover budget
  })
})

describe('pruneSession', () => {
  it('drops entries whose file no longer exists, window by window', () => {
    const state = {
      windows: [
        {
          tabs: [
            { kind: 'docs' as const, path: '/keep.docx' },
            { kind: 'pdf' as const, path: '/gone.pdf' },
          ],
          activePath: '/keep.docx',
        },
        { tabs: [{ kind: 'pdf' as const, path: '/gone2.pdf' }], activePath: '/gone2.pdf' },
      ],
      focusedWindow: 1,
    }
    const pruned = S.pruneSession(state, (p) => p === '/keep.docx')
    expect(pruned.windows).toEqual([
      { tabs: [{ kind: 'docs', path: '/keep.docx' }], activePath: '/keep.docx' },
    ])
    // the focused window vanished → focus falls to the last survivor
    expect(pruned.focusedWindow).toBe(0)
  })

  it('clears a window activePath whose file vanished but keeps the window', () => {
    const state = {
      windows: [
        {
          tabs: [
            { kind: 'docs' as const, path: '/keep.docx' },
            { kind: 'pdf' as const, path: '/gone.pdf' },
          ],
          activePath: '/gone.pdf',
        },
      ],
      focusedWindow: 0,
    }
    expect(S.pruneSession(state, (p) => p === '/keep.docx').windows[0]!.activePath).toBeNull()
  })
})

describe('session file round-trip', () => {
  it('writes atomically and reads back the same multi-window state', () => {
    const path = join(scratch, 'session.json')
    const state = {
      windows: [
        {
          tabs: [
            { kind: 'docs', path: '/a.docx' },
            { kind: 'html', path: '/b.html' },
          ],
          activePath: '/a.docx',
        },
        { tabs: [{ kind: 'pdf', path: '/c.pdf' }], activePath: '/c.pdf' },
      ],
      focusedWindow: 1,
    }
    S.writeSessionState(path, state)
    expect(JSON.parse(readFileSync(path, 'utf8'))).toEqual(state)
    expect(S.readSessionState(path)).toEqual(state)
  })

  it('reads a missing or corrupt file as null', () => {
    expect(S.readSessionState(join(scratch, 'absent.json'))).toBeNull()
    const corrupt = join(scratch, 'session.json')
    writeFileSync(corrupt, '{ nope', 'utf8')
    expect(S.readSessionState(corrupt)).toBeNull()
  })
})
