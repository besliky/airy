import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { splitBodyForHydration } from '../src/renderer/markdown/segments'
import { hydrateSegments, mountFirstSegment } from '../src/renderer/markdown/hydration'

// BUG-1685 / BUG-1693: saving a .md file silently destroyed every raw-HTML
// element the GFM schema had no node for (`<div class="callout">` 298B file
// shrank to 155B without a single edit). The RawHtml node captures unclaimed
// markup verbatim; these tests pin the save-path contract: open → save keeps
// the HTML, edits around it keep it, and the segmented hydration path (#150)
// produces the same document as the monolithic parse.

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

/** parse → serialize → parse → serialize; the second pass must be a fixed point */
function saveCycle(editor: Editor, md: string): { first: string; reopen: string } {
  const manager = editor.markdown!
  const first = manager.serialize(manager.parse(md))
  const reopen = manager.serialize(manager.parse(first))
  return { first, reopen }
}

// audit patho-nested-html.md: block div/script/style, inline script, script in
// a table cell (the audit's 298B → 155B loss)
const AUDIT_NESTED = [
  '# T',
  '',
  '<div class="callout">',
  'callout body',
  '</div>',
  '',
  '<script>window.__scriptRan = true</script>',
  '',
  'para with <script>alert(1)</script> inline',
  '',
  '<style>.x { color: red }</style>',
  '',
  '| s | <script>2</script> |',
  '| --- | --- |',
  '| a | b |',
  '',
].join('\n')

// audit patho-common-html.md: kbd/mark/details/span
const AUDIT_COMMON = [
  'Press <kbd>Ctrl</kbd> and <mark>mark</mark>.',
  '',
  '<details><summary>More</summary>hidden body</details>',
  '',
  '<span style="color:red">red span</span>',
  '',
].join('\n')

describe('raw HTML survives save (BUG-1685 audit vectors)', () => {
  it('keeps div/script/style blocks, inline script and script in a table cell', () => {
    const { first, reopen } = saveCycle(createEditor(), AUDIT_NESTED)
    expect(first).toContain('<div class="callout">')
    expect(first).toContain('callout body')
    expect(first).toContain('</div>')
    expect(first).toContain('<script>window.__scriptRan = true</script>')
    expect(first).toContain('<script>alert(1)</script>')
    expect(first).toContain('<style>.x { color: red }</style>')
    // the table cell keeps its raw script (cell padding is the serializer's
    // normal column alignment, not a loss)
    expect(first).toContain('<script>2</script>')
    // reopen: the saved file parses back to the same document and serializes
    // identically — no further churn
    expect(reopen).toBe(first)
  })

  it('keeps kbd, mark, details and styled span', () => {
    const { first, reopen } = saveCycle(createEditor(), AUDIT_COMMON)
    expect(first).toContain('<kbd>Ctrl</kbd>')
    expect(first).toContain('<mark>mark</mark>')
    expect(first).toContain('<details><summary>More</summary>hidden body</details>')
    expect(first).toContain('<span style="color:red">red span</span>')
    expect(reopen).toBe(first)
  })

  it('identity save: untouched HTML round-trips to the identical model', () => {
    const editor = createEditor()
    const manager = editor.markdown!
    const once = manager.parse(AUDIT_NESTED)
    const twice = manager.parse(manager.serialize(once))
    // a no-edit save followed by a reopen must not mutate the document
    expect(JSON.stringify(twice)).toBe(JSON.stringify(once))
  })

  it('editing a neighbor keeps the raw HTML (save after a real edit)', () => {
    const editor = createEditor()
    const manager = editor.markdown!
    editor.commands.setContent(manager.parse(AUDIT_COMMON), {})
    // type inside the first paragraph and at the very end of the document
    editor.commands.insertContentAt(1, 'EDITED ')
    const end = editor.state.doc.content.size
    editor.commands.insertContentAt(end, ' TAIL')
    const out = manager.serialize(manager.parse(editor.getMarkdown()))
    expect(out).toContain('EDITED Press')
    expect(out).toContain('<kbd>Ctrl</kbd>')
    expect(out).toContain('<mark>mark</mark>')
    expect(out).toContain('<details><summary>More</summary>hidden body</details>')
    expect(out).toContain('<span style="color:red">red span</span>')
    expect(out).toContain('TAIL')
  })
})

describe('raw HTML in inline contexts (own vectors)', () => {
  it('survives inside a list item', () => {
    const { first, reopen } = saveCycle(
      createEditor(),
      '- item with <sub>sub</sub> text\n- plain item',
    )
    expect(first).toContain('- item with <sub>sub</sub> text')
    expect(reopen).toBe(first)
  })

  it('survives inside a heading', () => {
    const { first, reopen } = saveCycle(createEditor(), '## Title <kbd>H</kbd> tail')
    expect(first).toContain('## Title <kbd>H</kbd> tail')
    expect(reopen).toBe(first)
  })

  it('survives inside a blockquote', () => {
    const { first, reopen } = saveCycle(
      createEditor(),
      '> quoted <span style="color:blue">span</span> text',
    )
    expect(first).toContain('<span style="color:blue">span</span>')
    expect(first).toMatch(/^> /)
    expect(reopen).toBe(first)
  })

  it('survives inside a table cell', () => {
    const { first, reopen } = saveCycle(
      createEditor(),
      '| a | b |\n| --- | --- |\n| <sup>x</sup> | y |',
    )
    expect(first).toContain('<sup>x</sup>')
    expect(first).toContain('| y ')
    expect(reopen).toBe(first)
  })

  it('keeps several raw elements in a row, back to back', () => {
    const { first, reopen } = saveCycle(
      createEditor(),
      'start <kbd>1</kbd><mark>2</mark><abbr title="ok">a</abbr> end',
    )
    expect(first).toContain('<kbd>1</kbd><mark>2</mark><abbr title="ok">a</abbr>')
    expect(reopen).toBe(first)
  })

  it('keeps a block-shaped element merged into an inline run', () => {
    // CommonMark lets block tags appear mid-paragraph; the merged raw token
    // must not lose its content either
    const { first, reopen } = saveCycle(
      createEditor(),
      'before <div class="d">inner text</div> after',
    )
    expect(first).toContain('<div class="d">inner text</div>')
    expect(first).toContain('before ')
    expect(first).toContain(' after')
    expect(reopen).toBe(first)
  })
})

describe('dedicated schema rules keep winning (no overcapture)', () => {
  it('GFM formatting still parses to its normal model', () => {
    const editor = createEditor()
    const json = editor.markdown!.parse('a **bold** *it* `code` [l](https://e.c) <br> end')
    const text = JSON.stringify(json)
    expect(text).toContain('"bold"') // strong mark, not a chip
    expect(text).toContain('"italic"')
    expect(text).toContain('"code"')
    expect(text).toContain('"link"')
    expect(text).toContain('hardBreak') // <br> keeps its legit conversion
    expect(text).not.toContain('rawHtml')
  })

  it('images still parse to image nodes', () => {
    const editor = createEditor()
    const json = editor.markdown!.parse('![pic](assets/p.png)')
    expect(JSON.stringify(json)).toContain('"image"')
    expect(JSON.stringify(json)).not.toContain('rawHtml')
  })
})

describe('stored HTML is inert in the editor', () => {
  it('script markup renders as escaped text, never as an executable element', () => {
    const editor = createEditor()
    editor.commands.setContent(editor.markdown!.parse('<script>alert(1)</script> ok'), {})
    const html = editor.getHTML()
    expect(html).not.toContain('<script')
    expect(html).toContain('&lt;script&gt;alert(1)&lt;/script&gt;')
  })
})

describe('segmented hydration parity (PERF-1647 path)', () => {
  const immediateYield = (): Promise<void> => Promise.resolve()

  /** mixed blocks with raw HTML in several constructs; long enough to split */
  function buildCorpus(blocks: number): string {
    const parts: string[] = []
    for (let i = 0; i < blocks; i++) {
      switch (i % 8) {
        case 0:
          parts.push(`## Head ${i} <kbd>H</kbd>`)
          break
        case 1:
          parts.push(`Paragraph ${i} with <mark>m${i}</mark> inline.`)
          break
        case 2:
          parts.push('- item <sub>s</sub>\n- plain')
          break
        case 3:
          parts.push('> quote <span style="x">q</span>')
          break
        case 4:
          parts.push('<div class="callout">\nbody\n</div>')
          break
        case 5:
          parts.push('| a | b |\n| --- | --- |\n| <sup>x</sup> | y |')
          break
        case 6:
          parts.push('<script>var v = 1</script>')
          break
        default:
          parts.push(`Plain ${i} <abbr title="t">a</abbr> text.`)
      }
    }
    return parts.join('\n\n') + '\n'
  }

  it('monolithic and incremental parses produce identical documents and bytes', async () => {
    const body = buildCorpus(48)
    const segments = splitBodyForHydration(body, { maxLines: 12, maxChars: 2048 })
    expect(segments.length).toBeGreaterThan(3)

    const incremental = createEditor()
    mountFirstSegment(incremental, segments[0])
    const ok = await hydrateSegments(incremental, segments, { yieldFn: immediateYield })
    expect(ok).toBe(true)

    const monolithic = createEditor()
    monolithic.commands.setContent(monolithic.markdown!.parse(body), {})

    expect(incremental.getJSON()).toEqual(monolithic.getJSON())
    // bit-for-bit markdown parity: the segmented save cannot differ from the
    // monolithic save
    expect(incremental.getMarkdown()).toBe(monolithic.getMarkdown())
  })

  it('the audit vector parses identically whichever way the body is fed', async () => {
    const segments = splitBodyForHydration(AUDIT_NESTED, {
      maxLines: 2,
      maxChars: 64,
    })
    const incremental = createEditor()
    mountFirstSegment(incremental, segments[0])
    await hydrateSegments(incremental, segments, { yieldFn: immediateYield })

    const monolithic = createEditor()
    monolithic.commands.setContent(monolithic.markdown!.parse(AUDIT_NESTED), {})
    const saved = incremental.getMarkdown()
    expect(saved).toBe(monolithic.getMarkdown())
    expect(saved).toContain('<div class="callout">')
    expect(saved).toContain('<script>alert(1)</script>')
  })
})
