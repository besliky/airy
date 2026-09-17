import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it, vi } from 'vitest'

import {
  recordDialogDir,
  readLastDialogDirs,
  saveAsSuggestion,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
  writeLastDialogDir,
} from '../src/index'

import type { Dialog } from 'electron'

function fakeDialog(overrides: Partial<Dialog> = {}): Dialog {
  return {
    showOpenDialog: vi.fn().mockResolvedValue({ canceled: true, filePaths: [] }),
    showSaveDialog: vi.fn().mockResolvedValue({ canceled: true, filePath: '' }),
    ...overrides,
  } as unknown as Dialog
}

const pickedOpen = (paths: string[]) =>
  vi.fn().mockResolvedValue({ canceled: false, filePaths: paths })
const pickedSave = (path: string) => vi.fn().mockResolvedValue({ canceled: false, filePath: path })

describe('showOpenDialogWithMemory', () => {
  it('passes options through unchanged before any pick was made', async () => {
    const dialog = fakeDialog()
    await showOpenDialogWithMemory(dialog, undefined, { properties: ['openFile'] })
    expect(dialog.showOpenDialog).toHaveBeenCalledWith({ properties: ['openFile'] })
  })

  it('remembers the picked directory and injects it as defaultPath next time', async () => {
    const dialog = fakeDialog({ showOpenDialog: pickedOpen([join('/work', 'report.docx')]) })
    await showOpenDialogWithMemory(dialog, undefined, {})
    await showOpenDialogWithMemory(dialog, undefined, {})
    expect(dialog.showOpenDialog).toHaveBeenLastCalledWith({ defaultPath: '/work' })
  })

  it('forwards the parent window when given', async () => {
    const dialog = fakeDialog()
    const parent = { id: 1 } as never
    await showOpenDialogWithMemory(dialog, parent, {})
    expect(dialog.showOpenDialog).toHaveBeenCalledWith(parent, {})
  })

  it('remembers the selected directory itself for openDirectory pickers', async () => {
    const dialog = fakeDialog({ showOpenDialog: pickedOpen(['/exports/images']) })
    await showOpenDialogWithMemory(dialog, undefined, { properties: ['openDirectory'] })
    await showOpenDialogWithMemory(dialog, undefined, {})
    expect(dialog.showOpenDialog).toHaveBeenLastCalledWith({ defaultPath: '/exports/images' })
  })

  it('keeps an explicit absolute defaultPath untouched', async () => {
    const dialog = fakeDialog({ showOpenDialog: pickedOpen([join('/work', 'a.docx')]) })
    await showOpenDialogWithMemory(dialog, undefined, {})
    await showOpenDialogWithMemory(dialog, undefined, { defaultPath: '/elsewhere/b.docx' })
    expect(dialog.showOpenDialog).toHaveBeenLastCalledWith({ defaultPath: '/elsewhere/b.docx' })
  })

  it('does not remember canceled picks', async () => {
    const dialog = fakeDialog()
    await showOpenDialogWithMemory(dialog, undefined, {})
    await showOpenDialogWithMemory(dialog, undefined, {})
    expect(dialog.showOpenDialog).toHaveBeenLastCalledWith({})
  })
})

describe('showSaveDialogWithMemory', () => {
  it('anchors a bare file-name suggestion in the remembered directory', async () => {
    const dialog = fakeDialog({ showSaveDialog: pickedSave(join('/work', 'deck.pptx')) })
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'deck.pptx' })
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'deck2.pptx' })
    expect(dialog.showSaveDialog).toHaveBeenLastCalledWith({
      defaultPath: join('/work', 'deck2.pptx'),
    })
  })

  it('shares the remembered directory between save and open dialogs', async () => {
    const dialog = fakeDialog({ showSaveDialog: pickedSave(join('/work', 'deck.pptx')) })
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'deck.pptx' })
    await showOpenDialogWithMemory(dialog, undefined, {})
    expect(dialog.showOpenDialog).toHaveBeenCalledWith({ defaultPath: '/work' })
  })

  it('keeps an explicit absolute defaultPath untouched', async () => {
    const dialog = fakeDialog({ showSaveDialog: pickedSave(join('/work', 'a.pdf')) })
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: '/docs/tab.pdf' })
    expect(dialog.showSaveDialog).toHaveBeenCalledWith({ defaultPath: '/docs/tab.pdf' })
  })

  it('isolates memory between different dialog instances', async () => {
    const first = fakeDialog({ showOpenDialog: pickedOpen([join('/work', 'a.docx')]) })
    const second = fakeDialog()
    await showOpenDialogWithMemory(first, undefined, {})
    await showOpenDialogWithMemory(second, undefined, {})
    expect(second.showOpenDialog).toHaveBeenCalledWith({})
  })

  it('anchors a bare file-name suggestion in the fallback dir when nothing is remembered', async () => {
    const dialog = fakeDialog()
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'deck.pptx' }, '/default/save')
    expect(dialog.showSaveDialog).toHaveBeenCalledWith({
      defaultPath: join('/default/save', 'deck.pptx'),
    })
  })

  it('prefers the remembered directory over the fallback dir', async () => {
    const dialog = fakeDialog({ showSaveDialog: pickedSave(join('/work', 'deck.pptx')) })
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'deck.pptx' })
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'b.pptx' }, '/default/save')
    expect(dialog.showSaveDialog).toHaveBeenLastCalledWith({
      defaultPath: join('/work', 'b.pptx'),
    })
  })

  it('keeps an explicit absolute defaultPath untouched even with a fallback dir', async () => {
    const dialog = fakeDialog()
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: '/docs/tab.pdf' }, '/default')
    expect(dialog.showSaveDialog).toHaveBeenCalledWith({ defaultPath: '/docs/tab.pdf' })
  })
})

describe('saveAsSuggestion', () => {
  it('suggests the source document folder with the new name (Word parity)', () => {
    expect(saveAsSuggestion(join('/work', 'report.docx'), 'report.docx')).toBe(
      join('/work', 'report.docx'),
    )
    expect(saveAsSuggestion(join('/work', 'report.docx'), 'copy.docx')).toBe(
      join('/work', 'copy.docx'),
    )
  })

  it('falls back to the bare name for a document that was never on disk', () => {
    expect(saveAsSuggestion(null, 'Untitled.docx')).toBe('Untitled.docx')
    expect(saveAsSuggestion(undefined, 'Untitled.docx')).toBe('Untitled.docx')
    expect(saveAsSuggestion('', 'Untitled.pptx')).toBe('Untitled.pptx')
  })

  it('beats the remembered directory when threaded through the save dialog', async () => {
    const dialog = fakeDialog({ showSaveDialog: pickedSave(join('/elsewhere', 'x.docx')) })
    // a pick in /elsewhere seeds the remembered directory…
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'x.docx' })
    // …but Save As of a document living in /work still opens in /work
    await showSaveDialogWithMemory(dialog, undefined, {
      defaultPath: saveAsSuggestion(join('/work', 'report.docx'), 'report.docx'),
    })
    expect(dialog.showSaveDialog).toHaveBeenLastCalledWith({
      defaultPath: join('/work', 'report.docx'),
    })
  })
})

describe('persisted dialog directories (lastDialogDirs LRU)', () => {
  it('moves the scope to the front and keeps other scopes', () => {
    const base = [
      { scope: 'a', dir: '/a' },
      { scope: 'b', dir: '/b' },
    ]
    expect(recordDialogDir(base, 'a', '/new')).toEqual([
      { scope: 'a', dir: '/new' },
      { scope: 'b', dir: '/b' },
    ])
    expect(recordDialogDir(base, 'c', '/c')).toEqual([
      { scope: 'c', dir: '/c' },
      { scope: 'a', dir: '/a' },
      { scope: 'b', dir: '/b' },
    ])
  })

  it('caps the list at 10 entries, evicting the least recently used', () => {
    let entries: Array<{ scope: string; dir: string }> = []
    for (let i = 0; i < 14; i++) entries = recordDialogDir(entries, `s${i}`, `/d${i}`)
    expect(entries).toHaveLength(10)
    expect(entries[0]).toEqual({ scope: 's13', dir: '/d13' })
    expect(entries.some((e) => e.scope === 's3')).toBe(false)
    expect(entries.some((e) => e.scope === 's4')).toBe(true)
  })

  it('round-trips through app-settings.json atomically', async () => {
    const scratch = mkdtempSync(join(tmpdir(), 'airy-dialog-dirs-'))
    try {
      const settingsPath = join(scratch, 'app-settings.json')
      expect(readLastDialogDirs(settingsPath)).toEqual([]) // absent file
      writeLastDialogDir(settingsPath, 'shell', '/work')
      writeLastDialogDir(settingsPath, 'docs', '/docs-dir')
      // an unrelated setting survives the read-merge-write
      const raw = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
      expect(raw.lastDialogDirs).toEqual([
        { scope: 'docs', dir: '/docs-dir' },
        { scope: 'shell', dir: '/work' },
      ])
      expect(readLastDialogDirs(settingsPath)).toEqual([
        { scope: 'docs', dir: '/docs-dir' },
        { scope: 'shell', dir: '/work' },
      ])
    } finally {
      rmSync(scratch, { recursive: true, force: true })
    }
  })

  it('tolerates malformed persisted values', () => {
    const scratch = mkdtempSync(join(tmpdir(), 'airy-dialog-dirs-bad-'))
    try {
      const settingsPath = join(scratch, 'app-settings.json')
      writeFileSync(
        settingsPath,
        JSON.stringify({
          lastDialogDirs: ['nope', 42, { scope: '', dir: '/x' }, { scope: 'ok', dir: '/ok' }],
        }),
        'utf8',
      )
      expect(readLastDialogDirs(settingsPath)).toEqual([{ scope: 'ok', dir: '/ok' }])
    } finally {
      rmSync(scratch, { recursive: true, force: true })
    }
  })
})
