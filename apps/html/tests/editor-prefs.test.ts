/**
 * UX-1704: the source editor's view preferences (word wrap, split-view
 * scroll-sync) persist workspace-wide in the shared app-settings.json under
 * `htmlEditorPrefs`, written exclusively through the single-writer queue
 * (PR #108 / OBS-1532 discipline). Parsing/merging is pure and unit-tested;
 * persistence is behavior-tested against a real temp settings file, the same
 * shape as encoding-memory.test.ts. The renderer→preload→main wiring is
 * pinned in editor-prefs-wiring.test.ts.
 */
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'

import { DEFAULT_EDITOR_PREFS } from '../src/shared/ipc'
import {
  EDITOR_PREFS_KEY,
  mergeEditorPrefs,
  parseEditorPrefs,
  readEditorPrefs,
  writeEditorPrefs,
} from '../src/main/editor-prefs'

const cleanup: string[] = []

function settingsPathIn(name: string): string {
  const dir = mkdtempSync(join(tmpdir(), 'html-editor-prefs-'))
  cleanup.push(dir)
  return join(dir, name)
}

afterEach(() => {
  for (const dir of cleanup.splice(0)) rmSync(dir, { recursive: true, force: true })
})

describe('parseEditorPrefs', () => {
  it('falls back to the defaults for anything that is not an object', () => {
    expect(parseEditorPrefs(undefined)).toEqual({ wordWrap: true, scrollSync: false })
    expect(parseEditorPrefs('nope')).toEqual(DEFAULT_EDITOR_PREFS)
    expect(parseEditorPrefs([true, false])).toEqual(DEFAULT_EDITOR_PREFS)
  })

  it('drops malformed fields individually — one bad key never disables the other', () => {
    expect(parseEditorPrefs({ wordWrap: 'yes', scrollSync: true })).toEqual({
      wordWrap: true,
      scrollSync: true,
    })
    expect(parseEditorPrefs({ wordWrap: false, scrollSync: 0 })).toEqual({
      wordWrap: false,
      scrollSync: false,
    })
  })

  it('reads back stored booleans verbatim', () => {
    expect(parseEditorPrefs({ wordWrap: false, scrollSync: true })).toEqual({
      wordWrap: false,
      scrollSync: true,
    })
  })
})

describe('mergeEditorPrefs', () => {
  it('applies only the fields the patch names', () => {
    expect(mergeEditorPrefs({ wordWrap: true, scrollSync: false }, { wordWrap: false })).toEqual({
      wordWrap: false,
      scrollSync: false,
    })
    expect(mergeEditorPrefs({ wordWrap: false, scrollSync: false }, { scrollSync: true })).toEqual({
      wordWrap: false,
      scrollSync: true,
    })
    expect(mergeEditorPrefs({ wordWrap: true, scrollSync: true }, {})).toEqual({
      wordWrap: true,
      scrollSync: true,
    })
  })

  it('ignores non-object patches and wrong-typed values', () => {
    const current = { wordWrap: true, scrollSync: false }
    expect(mergeEditorPrefs(current, null)).toBe(current)
    expect(mergeEditorPrefs(current, 'scrollSync')).toBe(current)
    expect(mergeEditorPrefs(current, { scrollSync: 'on' })).toEqual(current)
    expect(mergeEditorPrefs(current, [123])).toBe(current)
  })
})

describe('editor prefs persistence', () => {
  it('round-trips a patch and preserves unrelated settings keys', async () => {
    const settingsPath = settingsPathIn('app-settings.json')
    writeFileSync(
      settingsPath,
      JSON.stringify({ language: 'ru', fileEncodings: [{ path: '/a.html', encoding: 'big5' }] }),
    )
    const stored = await writeEditorPrefs(settingsPath, { scrollSync: true })
    expect(stored).toEqual({ wordWrap: true, scrollSync: true })
    expect(readEditorPrefs(settingsPath)).toEqual({ wordWrap: true, scrollSync: true })
    const settings = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
    // the merge-write kept keys this feature never touches (OBS-1532 discipline)
    expect(settings.language).toBe('ru')
    expect(settings.fileEncodings).toEqual([{ path: '/a.html', encoding: 'big5' }])
    expect(settings[EDITOR_PREFS_KEY]).toEqual({ wordWrap: true, scrollSync: true })
  })

  it('merges over the stored prefs instead of replacing them', async () => {
    const settingsPath = settingsPathIn('app-settings.json')
    writeFileSync(settingsPath, JSON.stringify({ [EDITOR_PREFS_KEY]: { wordWrap: false } }))
    await writeEditorPrefs(settingsPath, { scrollSync: true })
    expect(readEditorPrefs(settingsPath)).toEqual({ wordWrap: false, scrollSync: true })
  })

  it('returns the defaults for a missing or corrupt file', () => {
    expect(readEditorPrefs(settingsPathIn('never-written.json'))).toEqual(DEFAULT_EDITOR_PREFS)
    const corrupt = settingsPathIn('app-settings.json')
    writeFileSync(corrupt, '{not json')
    expect(readEditorPrefs(corrupt)).toEqual(DEFAULT_EDITOR_PREFS)
  })

  it('serialized writes keep every patch (the queue discipline)', async () => {
    const settingsPath = settingsPathIn('app-settings.json')
    await Promise.all([
      writeEditorPrefs(settingsPath, { wordWrap: false }),
      writeEditorPrefs(settingsPath, { scrollSync: true }),
      writeEditorPrefs(settingsPath, { wordWrap: true }),
    ])
    expect(readEditorPrefs(settingsPath)).toEqual({ wordWrap: true, scrollSync: true })
  })
})
