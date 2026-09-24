import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import {
  liftIndentedCodeAfterLists,
  stripBlankLinePadding,
} from '../src/renderer/markdown/parseContext'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom teardown
// ("document is not defined" unhandled error). Editors here are shared per describe,
// so destroy them once at the end of the file.
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

describe('BUG-1703: indented code block after a list', () => {
  const editor = createEditor()
  const manager = editor.markdown!

  it('lifts the audit fixture block out of the nested list item into a code block', () => {
    // the tabs/spaces audit fixture: a deep-indented code block at the tail of
    // an ordered list was absorbed into "tab ordered nested" as a paragraph
    const md = [
      '# Tabs',
      '',
      '- level one',
      '    - tab indented level two',
      '    - two space indented',
      '        - four space indented',
      '',
      '1. ordered one',
      '    1. tab ordered nested',
      '    ',
      '        plain four-space code block',
      '',
    ].join('\n')
    const doc = manager.parse(md)
    const top = doc.content ?? []
    expect(top.map((n) => n.type)).toEqual(['heading', 'bulletList', 'orderedList', 'codeBlock'])
    // the code is its own node, not list item content
    expect(top[3].content?.[0]?.text).toBe('plain four-space code block')
    const orderedItem = top[2].content?.[0]
    expect(orderedItem?.content?.map((n) => n.type)).toEqual(['paragraph', 'orderedList'])
    const nestedItem = orderedItem?.content?.[1].content?.[0]
    expect(nestedItem?.content?.map((n) => n.type)).toEqual(['paragraph'])
  })

  it('round-trips the lifted model stably (parse → serialize → parse)', () => {
    const md = '1. one\n    1. two\n    \n        code line\n'
    const first = manager.parse(md)
    const once = manager.serialize(first)
    // the serializer has no indented-code form: the block materializes fenced
    expect(once).toBe('1. one\n    1. two\n\n```\ncode line\n```')
    const second = manager.parse(once)
    expect(JSON.stringify(second)).toBe(JSON.stringify(first))
    expect(manager.serialize(second)).toBe(once)
  })

  it('keeps a one-unit indented loose-list paragraph inside its item (CommonMark shape)', () => {
    // "1. one\n\n    more about one" is ordinary nested paragraph content at
    // the serializer's own indentation — it must NOT become a code block
    const doc = manager.parse('1. one\n\n    more about one')
    const item = doc.content?.[0]
    expect(item?.type).toBe('orderedList')
    expect(item?.content?.[0]?.content?.map((n) => n.type)).toEqual(['paragraph', 'paragraph'])
    expect(item?.content?.[0]?.content?.[1].content?.[0]?.text).toBe('more about one')
  })

  it('does not lift when no blank line separates the block from the item', () => {
    const doc = manager.parse('1. one\n2. two\n        glued continuation')
    const item = doc.content?.[0]?.content?.[1]
    expect(item?.content?.map((n) => n.type)).toEqual(['paragraph'])
  })

  it('merges blank-separated deep chunks into one code block', () => {
    const lifted = liftIndentedCodeAfterLists('- a\n    - b\n\n        x\n\n        y\n')
    expect(lifted).toBe('- a\n    - b\n\n```\nx\n\ny\n```\n')
  })

  it('picks a fence no content line can close', () => {
    // the content holds a 4-backtick line, so the synthetic fence uses 5
    const lifted = liftIndentedCodeAfterLists('1. one\n\n        a\n        ````\n        b\n')
    expect(lifted).toBe('1. one\n\n`````\na\n````\nb\n`````\n')
  })

  it('never lifts inside fenced code or raw HTML regions', () => {
    const fenced = '1. one\n\n```\nplain\n\n        still code\n```\n'
    expect(liftIndentedCodeAfterLists(fenced)).toBe(fenced)
    const html = '1. one\n\n<pre>\n\n        still html\n</pre>\n'
    expect(liftIndentedCodeAfterLists(html)).toBe(html)
  })
})

describe('BUG-1703: empty paragraph in a list serializes without trailing spaces', () => {
  const editor = createEditor()
  const manager = editor.markdown!

  it('serializes a nested empty paragraph as a blank line, not "    "', () => {
    const doc = {
      type: 'doc',
      content: [
        {
          type: 'bulletList',
          content: [
            {
              type: 'listItem',
              content: [
                { type: 'paragraph', content: [{ type: 'text', text: 'a' }] },
                { type: 'paragraph' },
                { type: 'paragraph', content: [{ type: 'text', text: 'b' }] },
              ],
            },
          ],
        },
      ],
    }
    const out = manager.serialize(doc as never)
    expect(out).not.toMatch(/^[ \t]+$/m)
    expect(out).toBe('- a\n\n\n\n    b')
    // the strip is parse-neutral: the padded ("    ") and blanked forms of the
    // same document lex identically, and the empty paragraph survives as a
    // real node (marked itself keeps continuation indentation on "b")
    const padded = '- a\n\n    \n\n    b'
    expect(JSON.stringify(manager.parse(out))).toBe(JSON.stringify(manager.parse(padded)))
    const item = manager.parse(out).content?.[0]?.content?.[0]
    expect(item?.content?.map((n) => n.type)).toEqual(['paragraph', 'paragraph', 'paragraph'])
  })

  it('keeps hard-break trailing spaces after text', () => {
    expect(stripBlankLinePadding('a  \nb\n')).toBe('a  \nb\n')
  })

  it('keeps whitespace-only lines inside fenced code and block math', () => {
    const fenced = '```js\nx()\n    \n}\n```\n'
    expect(stripBlankLinePadding(fenced)).toBe(fenced)
    const math = '$$\n\\begin{array}{ll}\na \\\\\n    \n\\end{array}\n$$\n'
    expect(stripBlankLinePadding(math)).toBe(math)
  })
})
