import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { splitBodyForHydration } from '../src/renderer/markdown/segments'
import {
  HYDRATION_META,
  hydrateSegments,
  mountFirstSegment,
} from '../src/renderer/markdown/hydration'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom
// teardown ("document is not defined" unhandled error) — destroy everything
// once at the end of the file.
const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

function createEditor(): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  editors.push(editor)
  return editor
}

/** mixed top-level blocks; long enough that maxLines forces several segments */
function buildCorpus(blocks: number): string {
  const parts: string[] = []
  for (let i = 0; i < blocks; i++) {
    switch (i % 10) {
      case 0:
        parts.push(`## Heading ${i}`)
        break
      case 1:
        parts.push(`Paragraph ${i} with **bold** and *italic* text.`)
        break
      case 2:
        parts.push('- alpha\n- beta\n- gamma')
        break
      case 3:
        parts.push('1. first\n2. second\n3. third')
        break
      case 4:
        parts.push('> quoted line\n> another quoted line')
        break
      case 5:
        parts.push('```\nfenced code\n\nwith a blank line\n```')
        break
      case 6:
        parts.push('| a | b |\n| --- | --- |\n| 1 | 2 |')
        break
      case 7:
        parts.push('Setext heading\n--------------')
        break
      case 8:
        parts.push('- loose one\n\n- loose two\n\n- loose three')
        break
      default:
        parts.push(`Plain paragraph ${i} with text.`)
    }
  }
  return parts.join('\n\n') + '\n'
}

/** instrument markdown.parse: records every input it is called with */
function spyOnParse(editor: Editor): string[] {
  const calls: string[] = []
  const manager = editor.markdown
  const original = manager.parse.bind(manager)
  manager.parse = ((md: string) => {
    calls.push(md)
    return original(md)
  }) as typeof manager.parse
  return calls
}

const immediateYield = (): Promise<void> => Promise.resolve()

describe('segmented hydration', () => {
  it('parses each fragment once and never re-parses the document on edits', async () => {
    const body = buildCorpus(120)
    const segments = splitBodyForHydration(body, { maxLines: 24, maxChars: 4096 })
    expect(segments.length).toBeGreaterThan(4)

    const editor = createEditor()
    const calls = spyOnParse(editor)
    mountFirstSegment(editor, segments[0])
    const ok = await hydrateSegments(editor, segments, { yieldFn: immediateYield })
    expect(ok).toBe(true)

    // exactly one parse per segment — each a small fragment, never the whole body
    expect(calls.length).toBe(segments.length)
    expect(Math.max(...calls.map((c) => c.length))).toBeLessThan(body.length)
    expect(calls.map((c) => c.length).reduce((a, b) => a + b, 0)).toBeLessThanOrEqual(body.length)

    const blocksAfterHydration = editor.state.doc.childCount
    expect(blocksAfterHydration).toBeGreaterThanOrEqual(120)

    // an edit in the middle of the document must not re-parse anything:
    // the incremental path computes the change from the transaction alone
    const callsBefore = calls.length
    const mid = Math.floor(editor.state.doc.content.size / 2)
    editor.commands.insertContentAt(mid, 'EDIT', { updateSelection: false })
    expect(calls.length).toBe(callsBefore)
    expect(editor.state.doc.textContent).toContain('EDIT')
  })

  it('hydration transactions are not undoable and carry the hydration meta', async () => {
    const body = buildCorpus(60)
    const segments = splitBodyForHydration(body, { maxLines: 12, maxChars: 2048 })
    expect(segments.length).toBeGreaterThan(2)

    const editor = createEditor()
    const metas: unknown[] = []
    editor.on('transaction', ({ transaction }) => {
      metas.push(transaction.getMeta(HYDRATION_META))
    })
    mountFirstSegment(editor, segments[0])
    await hydrateSegments(editor, segments, { yieldFn: immediateYield })

    // hydration re-assembles the file just read from disk — nothing to undo
    expect(editor.can().undo()).toBe(false)
    expect(metas.length).toBeGreaterThan(0)
    expect(metas.every((meta) => meta === true)).toBe(true)

    // a user edit after hydration is undoable, and undo lands exactly on the
    // fully hydrated state (not one step into it)
    const hydratedJson = editor.getJSON()
    editor.commands.insertContentAt(1, 'X')
    expect(editor.can().undo()).toBe(true)
    editor.commands.undo()
    expect(editor.getJSON()).toEqual(hydratedJson)
  })

  it('matches a monolithic full parse bit for bit, before and after edits', async () => {
    const body = buildCorpus(120)
    const segments = splitBodyForHydration(body, { maxLines: 24, maxChars: 4096 })
    expect(segments.length).toBeGreaterThan(4)

    // incremental path: first segment mounted, the rest appended
    const incremental = createEditor()
    mountFirstSegment(incremental, segments[0])
    await hydrateSegments(incremental, segments, { yieldFn: immediateYield })

    // baseline: one full-document parse
    const monolithic = createEditor()
    monolithic.commands.setContent(monolithic.markdown.parse(body), {})

    expect(incremental.getJSON()).toEqual(monolithic.getJSON())

    // identical edit series on both documents (same starting positions)
    const edits: Array<'insert' | 'delete' | 'lineBreak'> = [
      'insert',
      'delete',
      'lineBreak',
      'insert',
    ]
    let at = 40
    for (const kind of edits) {
      if (kind === 'insert') {
        incremental.commands.insertContentAt(at, 'EDIT', { updateSelection: false })
        monolithic.commands.insertContentAt(at, 'EDIT', { updateSelection: false })
        at += 4
      } else if (kind === 'delete') {
        incremental.commands.deleteRange({ from: at, to: at + 6 })
        monolithic.commands.deleteRange({ from: at, to: at + 6 })
      } else {
        // Enter is selection-dependent — pin both editors to the same spot
        incremental.commands.setTextSelection(at)
        monolithic.commands.setTextSelection(at)
        incremental.commands.keyboardShortcut('Enter')
        monolithic.commands.keyboardShortcut('Enter')
      }
    }

    // the incremental result must serialize byte-identically to the full parse
    expect(incremental.getMarkdown()).toBe(monolithic.getMarkdown())
  })

  it('joins lists that become adjacent across a chunk boundary', async () => {
    const editor = createEditor()
    // first segment already in place as a list; the appended segment starts
    // with the same kind of list — the monolithic parse renders one list
    mountFirstSegment(editor, '- alpha\n')
    const ok = await hydrateSegments(editor, ['- alpha\n', '- beta\n- gamma\n'], {
      yieldFn: immediateYield,
    })
    expect(ok).toBe(true)

    const monolithic = createEditor()
    monolithic.commands.setContent(monolithic.markdown.parse('- alpha\n- beta\n- gamma\n'), {})

    expect(editor.getJSON()).toEqual(monolithic.getJSON())
    expect(editor.getMarkdown()).toBe(monolithic.getMarkdown())
  })

  it('reports progress and stops cleanly when the editor is destroyed', async () => {
    const body = buildCorpus(80)
    const segments = splitBodyForHydration(body, { maxLines: 16, maxChars: 4096 })
    expect(segments.length).toBeGreaterThan(3)

    const editor = createEditor()
    const progress: number[] = []
    editor.commands.setContent(segments[0], { contentType: 'markdown' })
    const run = hydrateSegments(editor, segments, {
      yieldFn: immediateYield,
      onProgress: (p) => progress.push(p.done),
    })
    editor.destroy()
    const ok = await run
    expect(ok).toBe(false)
  })
})
