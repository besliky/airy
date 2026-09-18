import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { findMatches } from '../src/renderer/components/FindPanel'

function createEditor(text: string): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text }],
        },
      ],
    },
  })
}

const OFF = { matchCase: false, wholeWord: false, useWildcards: true }

/** matched substrings of the document, in order */
function matchTexts(editor: Editor, query: string, opts = OFF): string[] {
  return findMatches(editor, query, opts).map((m) => editor.state.doc.textBetween(m.from, m.to))
}

describe('wildcards findMatches', () => {
  it('? matches exactly one character', () => {
    const editor = createEditor('test tast toast')
    expect(matchTexts(editor, 't?st')).toEqual(['test', 'tast'])
    editor.destroy()
  })

  it('? matches one astral character (emoji) as a single char', () => {
    const editor = createEditor('🙂🙃')
    expect(matchTexts(editor, '?')).toEqual(['🙂', '🙃'])
    editor.destroy()
  })

  it('* is lazy like Word: s*d finds sad before started', () => {
    const editor = createEditor('sad started')
    expect(matchTexts(editor, 's*d')).toEqual(['sad', 'started'])
    editor.destroy()
  })

  it('* only extends as far as the rest of the pattern forces it to', () => {
    const editor = createEditor('test tasting')
    expect(matchTexts(editor, 't?st*g')).toEqual(['test tasting'])
    editor.destroy()
  })

  it('character ranges and sets match', () => {
    const editor = createEditor('boy box mod toy')
    expect(matchTexts(editor, '[a-m]o[yd]')).toEqual(['boy', 'mod'])
    editor.destroy()
  })

  it('negated sets exclude the listed characters', () => {
    const editor = createEditor('cat bat hat sat')
    expect(matchTexts(editor, '[!abc]at')).toEqual(['hat', 'sat'])
    editor.destroy()
  })

  it('CJK classes and ranges work by code point', () => {
    // あ..う is a genuine code-point range (U+3042..U+3046)
    const editor = createEditor('あいうえ')
    expect(matchTexts(editor, '[あ-う]')).toEqual(['あ', 'い', 'う'])
    expect(matchTexts(editor, '[!あ-う]')).toEqual(['え'])
    editor.destroy()
  })

  it('escapes make wildcard characters literal', () => {
    const editor = createEditor('a*b ?q [w')
    expect(matchTexts(editor, 'a\\*b')).toEqual(['a*b'])
    expect(matchTexts(editor, '[?]')).toEqual(['?'])
    editor.destroy()
  })

  it('unterminated classes and escaped class members fall back to literals', () => {
    const unl = createEditor('x [abc')
    expect(matchTexts(unl, '[abc')).toEqual(['[abc']) // no closing ]
    unl.destroy()
    const lead = createEditor('-y z')
    expect(matchTexts(lead, '[-y]')).toEqual(['-', 'y']) // leading - is literal
    lead.destroy()
    const esc = createEditor('a-b')
    expect(matchTexts(esc, '[a\\-b]')).toEqual(['a', '-', 'b']) // escaped - is literal
    esc.destroy()
  })

  it('reversed ranges match nothing; negated reversed ranges match anything', () => {
    const editor = createEditor('boy')
    expect(matchTexts(editor, '[z-a]oy')).toEqual([])
    expect(matchTexts(editor, '[!z-a]oy')).toEqual(['boy'])
    editor.destroy()
  })

  it('an empty pattern or a bare * finds nothing', () => {
    const editor = createEditor('anything')
    expect(findMatches(editor, '', OFF)).toEqual([])
    expect(matchTexts(editor, '*')).toEqual([])
    editor.destroy()
  })

  it('matchCase applies to wildcards too', () => {
    const editor = createEditor('Test test')
    expect(matchTexts(editor, 'T?st', { ...OFF, matchCase: true })).toEqual(['Test'])
    expect(matchTexts(editor, 'T?st', { ...OFF, matchCase: false })).toEqual(['Test', 'test'])
    editor.destroy()
  })

  it('wholeWord is ignored in wildcards mode', () => {
    const editor = createEditor('scat cat')
    // the first match sits inside a word, so whole-word plain search skips it
    expect(
      findMatches(editor, 'cat', { matchCase: false, wholeWord: true, useWildcards: false }),
    ).toHaveLength(1)
    const wild = { matchCase: false, wholeWord: true, useWildcards: true }
    expect(matchTexts(editor, '?cat', wild)).toEqual(['scat', ' cat'])
    expect(matchTexts(editor, '?cat', { ...wild, wholeWord: false })).toEqual(['scat', ' cat'])
    editor.destroy()
  })

  it('matches across two paragraphs are all collected', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [{ type: 'text', text: 'first test' }],
          },
          {
            type: 'docParagraph',
            attrs: { docxIndex: 1 },
            content: [{ type: 'text', text: 'toast' }],
          },
        ],
      },
    })
    expect(matchTexts(editor, 't?st')).toEqual(['test'])
    expect(matchTexts(editor, 't??st')).toEqual(['toast'])
    editor.destroy()
  })
})
