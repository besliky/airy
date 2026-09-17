import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

/**
 * Session persistence (src/main/session-state.ts): serialize the open-tab
 * set, parse/validate persisted JSON, prune entries whose files vanished,
 * and round-trip the session file atomically.
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

describe('serializeSession', () => {
  it('keeps file-backed editor tabs in strip order and skips untitled/home/present tabs', () => {
    const state = S.serializeSession(
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
    const state = S.serializeSession([{ id: 't1', kind: 'markdown' as const }], 't1')
    expect(state.activePath).toBeNull()
  })

  it('records no activePath when no tab matches the active id', () => {
    const state = S.serializeSession(
      [{ id: 't1', kind: 'docs' as const, filePath: '/a.docx' }],
      't9',
    )
    expect(state.activePath).toBeNull()
  })
})

describe('parseSession', () => {
  it('parses a saved session', () => {
    const state = S.parseSession({
      tabs: [
        { kind: 'docs', path: '/a.docx' },
        { kind: 'markdown', path: '/b.md' },
      ],
      activePath: '/b.md',
    })
    expect(state.tabs).toHaveLength(2)
    expect(state.activePath).toBe('/b.md')
  })

  it('yields an empty state for malformed payloads', () => {
    for (const bad of [null, undefined, 'x', 42, [], {}, { tabs: 'nope' }, { tabs: [1, 2] }]) {
      expect(S.parseSession(bad)).toEqual({ tabs: [], activePath: null })
    }
  })

  it('skips invalid entries without aborting the rest', () => {
    const state = S.parseSession({
      tabs: [
        null,
        { kind: 'home', path: '/x' }, // home is not restorable
        { kind: 'docs' }, // missing path
        { kind: 'unknown-kind', path: '/y' },
        { kind: 'pdf', path: '' },
        { kind: 'sheets', path: '/ok.xlsx' },
      ],
    })
    expect(state.tabs).toEqual([{ kind: 'sheets', path: '/ok.xlsx' }])
  })

  it('drops duplicate paths', () => {
    const state = S.parseSession({
      tabs: [
        { kind: 'docs', path: '/a.docx' },
        { kind: 'docs', path: '/a.docx' },
      ],
      activePath: '/a.docx',
    })
    expect(state.tabs).toHaveLength(1)
  })

  it('ignores an activePath that no listed tab owns', () => {
    const state = S.parseSession({
      tabs: [{ kind: 'docs', path: '/a.docx' }],
      activePath: '/gone.docx',
    })
    expect(state.activePath).toBeNull()
  })

  it('caps runaway tab lists', () => {
    const tabs = Array.from({ length: 200 }, (_, i) => ({ kind: 'docs', path: `/f${i}.docx` }))
    expect(S.parseSession({ tabs }).tabs).toHaveLength(64)
  })
})

describe('pruneSession', () => {
  it('drops entries whose file no longer exists', () => {
    const state = {
      tabs: [
        { kind: 'docs' as const, path: '/keep.docx' },
        { kind: 'pdf' as const, path: '/gone.pdf' },
      ],
      activePath: '/keep.docx',
    }
    expect(S.pruneSession(state, (p) => p === '/keep.docx').tabs).toEqual([
      { kind: 'docs', path: '/keep.docx' },
    ])
  })

  it('clears the activePath when the active file vanished', () => {
    const state = {
      tabs: [
        { kind: 'docs' as const, path: '/keep.docx' },
        { kind: 'pdf' as const, path: '/gone.pdf' },
      ],
      activePath: '/gone.pdf',
    }
    expect(S.pruneSession(state, (p) => p === '/keep.docx').activePath).toBeNull()
  })
})

describe('session file round-trip', () => {
  it('writes atomically and reads back the same state', () => {
    const path = join(scratch, 'session.json')
    const state = {
      tabs: [
        { kind: 'docs', path: '/a.docx' },
        { kind: 'html', path: '/b.html' },
      ],
      activePath: '/a.docx',
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
