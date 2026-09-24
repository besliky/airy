import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import {
  MAX_PAINTED_HITS,
  REPLACE_ALL_SYNC_LIMIT,
  findMatches,
  pickHighlightedRanges,
  tiptapFindTarget,
  type FindStatus,
} from '../src/renderer/editor/findTarget'
import { searchPluginKey } from '../src/renderer/editor/searchHighlight'

function createEditor(markdown: string): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: markdown,
    contentType: 'markdown',
  })
}

/** poll an async condition (chunked replaceAll settles over macrotasks) */
async function until(fn: () => boolean, ms = 5000): Promise<void> {
  const start = Date.now()
  while (!fn()) {
    if (Date.now() - start > ms) throw new Error('condition not reached')
    await new Promise((r) => setTimeout(r, 10))
  }
}

const OPTS = { matchCase: false, wholeWord: false }

/** a one-paragraph doc with `n` copies of the word */
const repeatedParagraph = (n: number, word = 'alpha') => `${Array(n).fill(word).join(' ')}\n`

afterEach(() => {
  vi.useRealTimers()
})

describe('pickHighlightedRanges (decoration cap)', () => {
  const range = (from: number) => ({ from, to: from + 5 })
  const full = Array.from({ length: 2000 }, (_, i) => range(i * 10))

  it('returns the full list when it fits under the cap', () => {
    const small = full.slice(0, 10)
    expect(pickHighlightedRanges(small, 3, undefined)).toBe(small)
  })

  it('caps the paint, always keeps the active match and stays honest: painted < total', () => {
    const painted = pickHighlightedRanges(full, 1999, undefined)
    expect(painted.length).toBe(MAX_PAINTED_HITS)
    expect(painted.some((r) => r.from === full[1999]!.from)).toBe(true)
    // overflow the host is told about
    const unpainted = full.length - painted.length
    expect(unpainted).toBe(full.length - MAX_PAINTED_HITS)
  })

  it('prefers the matches flagged as visible over the earliest ones', () => {
    // viewport window reported by the host: matches #1500..#1519, active elsewhere
    const visibleIndexes = Array.from({ length: 20 }, (_, k) => 1500 + k)
    const painted = pickHighlightedRanges(full, 0, visibleIndexes)
    const fromVisible = painted.filter(
      (r) => r.from >= full[1500]!.from && r.from <= full[1519]!.from,
    )
    expect(fromVisible.length).toBe(visibleIndexes.length)
    // with the viewport satisfied, the remaining slots still go to the earliest matches
    expect(painted.some((r) => r.from === full[0]!.from)).toBe(true)
  })

  it('never exceeds the cap even when the visible set is huge', () => {
    const visibleIndexes = Array.from({ length: 2000 }, (_, k) => k)
    expect(pickHighlightedRanges(full, 0, visibleIndexes).length).toBe(MAX_PAINTED_HITS)
  })

  it('keeps the result sorted by document position', () => {
    const painted = pickHighlightedRanges(full, 777, [{ from: 15000, to: 16000 }])
    for (let i = 1; i < painted.length; i++) {
      expect(painted[i]!.from).toBeGreaterThanOrEqual(painted[i - 1]!.from)
    }
  })
})

describe('giant-document find honesty (BUG-1694)', () => {
  it('search returns the full count while painting at most the cap', () => {
    const n = 1200
    const editor = createEditor(repeatedParagraph(n))
    const target = tiptapFindTarget(editor)
    const statuses: FindStatus[] = []
    target.onStatus((s) => statuses.push(s))
    expect(target.search('alpha', OPTS, 0)).toBe(n)
    const painted = searchPluginKey.getState(editor.state)!.find().length
    expect(painted).toBe(MAX_PAINTED_HITS)
    expect(
      statuses.some((s) => s.kind === 'overflow' && s.unpainted === n - MAX_PAINTED_HITS),
    ).toBe(true)
    target.clear()
    editor.destroy()
  })

  it('every match stays navigable/replacable even when not painted', () => {
    const editor = createEditor(repeatedParagraph(700))
    const target = tiptapFindTarget(editor)
    target.search('alpha', OPTS, 0)
    // painted is capped, but the last match (index 699) is still replaceable
    target.replaceOne(699, 'ZZZ')
    expect(editor.getText()).toContain('ZZZ')
    expect(target.search('alpha', OPTS, 0)).toBe(699)
    target.clear()
    editor.destroy()
  })
})

describe('replaceAll over capped matches', () => {
  it('sync path: replaces all matches in one transaction and one undo step', () => {
    const n = Math.min(REPLACE_ALL_SYNC_LIMIT, 600)
    const editor = createEditor(repeatedParagraph(n))
    const target = tiptapFindTarget(editor)
    expect(target.search('alpha', OPTS, 0)).toBe(n)
    target.replaceAll('QQ')
    expect(editor.getText()).not.toContain('alpha')
    expect(editor.getText().match(/\bQQ\b/g)?.length).toBe(n)
    // single undo restores everything (one transaction)
    editor.commands.undo()
    expect(editor.getText()).toContain('alpha')
    expect(target.search('alpha', OPTS, 0)).toBe(n)
    target.clear()
    editor.destroy()
  })

  it('large replace: kicked to the next tick and completes beyond the sync limit', async () => {
    const n = REPLACE_ALL_SYNC_LIMIT + 200
    const editor = createEditor(repeatedParagraph(n))
    const target = tiptapFindTarget(editor)
    expect(target.search('alpha', OPTS, 0)).toBe(n)
    target.replaceAll('QQ')
    // not applied synchronously (the click handler returns first)
    expect(editor.getText()).toContain('alpha')
    await until(() => !editor.getText().includes('alpha'))
    expect(editor.getText().match(/QQ/g)?.length).toBe(n)
    target.clear()
    editor.destroy()
  })

  it('large replace: emits replace status up front and idle when done, single undo step', async () => {
    const n = REPLACE_ALL_SYNC_LIMIT + 200
    const editor = createEditor(repeatedParagraph(n))
    const target = tiptapFindTarget(editor)
    const statuses: FindStatus[] = []
    target.onStatus((s) => statuses.push(s))
    target.search('alpha', OPTS, 0)
    target.replaceAll('QQ')
    await until(() => !editor.getText().includes('alpha'))
    await until(() => statuses.some((s) => s.kind === 'idle'))
    expect(statuses.some((s) => s.kind === 'replace' && s.remaining === n && s.total === n)).toBe(
      true,
    )
    // one transaction despite thousands of matches — a single undo restores all
    editor.commands.undo()
    expect(editor.getText().match(/alpha/g)?.length).toBe(n)
    target.clear()
    editor.destroy()
  })

  it('clear() before the deferred dispatch aborts a large replace and keeps the target usable', async () => {
    const n = REPLACE_ALL_SYNC_LIMIT + 200
    const editor = createEditor(repeatedParagraph(n))
    const target = tiptapFindTarget(editor)
    target.search('alpha', OPTS, 0)
    target.replaceAll('QQ')
    target.clear() // cancels before the next-tick dispatch
    await new Promise((r) => setTimeout(r, 30))
    expect(editor.getText().match(/alpha/g)?.length).toBe(n)
    // target must be usable again: a fresh search + replace goes through
    expect(target.search('alpha', OPTS, 0)).toBe(n)
    target.replaceAll('WW')
    await until(() => !editor.getText().includes('alpha'))
    target.clear()
    editor.destroy()
  })
})

describe('replaceAll preserves structure (slab rebuild)', () => {
  it('keeps marks on untouched text and inherits match-start marks for replacements', () => {
    const editor = createEditor('one **two** one two\n')
    const target = tiptapFindTarget(editor)
    target.search('two', OPTS, 0)
    target.replaceAll('X')
    const text = editor.getText()
    expect(text).toBe('one X one X')
    const para = editor.state.doc.firstChild!
    const boldChild: string[] = []
    para.forEach((child) => {
      if (child.marks.some((m) => m.type.name === 'bold')) boldChild.push(child.text ?? '')
    })
    // same semantics as the old insertText loop: the replacement inherits the
    // marks at the match start; the first 'two' was the bold one, so 'X' is bold
    expect(boldChild.join('')).toBe('X')
    target.clear()
    editor.destroy()
  })

  it('merges same-mark runs so a 43k-match replace does not explode the block into tiny nodes', () => {
    const editor = createEditor(repeatedParagraph(300))
    const target = tiptapFindTarget(editor)
    target.search('alpha', OPTS, 0)
    target.replaceAll('ZZ')
    const para = editor.state.doc.firstChild!
    // one flat text run, not 300 replacement fragments
    expect(para.childCount).toBe(1)
    expect(para.firstChild!.isText).toBe(true)
    target.clear()
    editor.destroy()
  })

  it('replaces matches spanning mark boundaries inside one word', () => {
    const editor = createEditor('fo**o**bar baz\n')
    const ranges = findMatches(editor, 'foobar', OPTS)
    expect(ranges).toHaveLength(1)
    const target = tiptapFindTarget(editor)
    target.search('foobar', OPTS, 0)
    target.replaceOne(0, 'QUUX')
    expect(editor.getText()).toBe('QUUX baz')
    target.clear()
    editor.destroy()
  })

  it('empty replacement deletes every match', () => {
    const editor = createEditor('alpha beta alpha beta\n')
    const target = tiptapFindTarget(editor)
    target.search('alpha ', OPTS, 0)
    target.replaceAll('')
    expect(editor.getText()).toBe('beta beta')
    target.clear()
    editor.destroy()
  })
})
