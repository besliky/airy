import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { findMatches } from '../src/renderer/components/FindPanel'
import { compileWildcards, foldDiacritics } from '../src/renderer/find-wildcards'

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

const OFF = { matchCase: false, wholeWord: false, useWildcards: true, ignoreDiacritics: false }

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
      findMatches(editor, 'cat', {
        matchCase: false,
        wholeWord: true,
        useWildcards: false,
        ignoreDiacritics: false,
      }),
    ).toHaveLength(1)
    const wild = { ...OFF, wholeWord: true }
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

  it('an astral placeholder-free text keeps UTF-16 offsets exact', () => {
    const editor = createEditor('🙂a🙂b')
    expect(matchTexts(editor, 'a?b')).toEqual(['a🙂b'])
    editor.destroy()
  })
})

describe('leaf placeholders are invisible to wildcards (BUG-741)', () => {
  /** paragraph `a` + hardBreak + `b` — the break flattens to `\u0000` */
  const makeWithBreak = () =>
    new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [
              { type: 'text', text: 'a' },
              { type: 'hardBreak' },
              { type: 'text', text: 'b' },
            ],
          },
        ],
      },
    })

  it('? never matches the \\u0000 placeholder of an inline leaf node', () => {
    const editor = makeWithBreak()
    // a hardBreak (same for inline image/math/ruby/note ref) sits between a and b
    expect(findMatches(editor, 'a?b', OFF)).toEqual([])
    // the two real characters still match
    expect(matchTexts(editor, '?')).toEqual(['a', 'b'])
    editor.destroy()
  })

  it('* cannot span the placeholder, so Replace never eats an inline node', () => {
    const editor = makeWithBreak()
    expect(findMatches(editor, 'a*b', OFF)).toEqual([])
    expect(matchTexts(editor, 'a*')).toEqual(['a'])
    editor.destroy()
  })

  it('negated classes refuse the placeholder as well', () => {
    const editor = makeWithBreak()
    expect(findMatches(editor, 'a[!x]b', OFF)).toEqual([])
    expect(matchTexts(editor, '[!a]')).toEqual(['b'])
    // a negated empty/reversed class is "any character" — but still not the placeholder
    expect(matchTexts(editor, '[!z-a]')).toEqual(['a', 'b'])
    editor.destroy()
  })
})

describe('compileWildcards.execAll', () => {
  it('reports UTF-16 offsets, also across astral characters', () => {
    // 🙂 a 🙃 spans 5 UTF-16 units; the next slot (b) cannot host the 3-code-point pattern
    expect(compileWildcards('?a?', false)!.execAll('🙂a🙃b')).toEqual([{ start: 0, end: 5 }])
    expect(compileWildcards('a?b', false)!.execAll('a🙂b')).toEqual([{ start: 0, end: 4 }])
    // leftmost + non-overlapping, like the old gu-flag exec loop
    expect(compileWildcards('aa', false)!.execAll('aaa')).toEqual([{ start: 0, end: 2 }])
    expect(compileWildcards('', false)).toBeNull()
  })
})

describe('wildcard pattern bombs stay linear (BUG-742)', () => {
  // one ~50 KB paragraph — the shape on which chained lazy `*` quantifiers
  // used to backtrack catastrophically and freeze the renderer
  const BIG = 'the quick brown fox jumps over the lazy dog. '.repeat(1119)

  it('completes star-chain bombs over 50 KB in under 100 ms each', () => {
    const editor = createEditor(BIG)
    expect(BIG.length).toBeGreaterThanOrEqual(50_000)
    const bombs = [
      '*a*a*a*a*z', // 'z' occurs ("lazy"): a match exists
      '*a*a*a*a*a*a*a*a*z',
      '*e*e*e*e*e*e*e*e*e*e*Z9', // no Z9 anywhere: full scan, no match
      '*o*o*o*o*o*o*o*o*o*o*o*o*#',
    ]
    for (const bomb of bombs) {
      const t0 = performance.now()
      const found = findMatches(editor, bomb, OFF)
      expect(performance.now() - t0).toBeLessThan(100)
      if (bomb.endsWith('Z9') || bomb.endsWith('#')) expect(found).toEqual([])
      else expect(found.length).toBeGreaterThan(0)
    }
    editor.destroy()
  })

  it('everyday patterns keep their results and speed on the same 50 KB', () => {
    const editor = createEditor(BIG)
    expect(matchTexts(editor, '*quick*')).toHaveLength(1119)
    expect(matchTexts(editor, '?azy')).toHaveLength(1119)
    const t0 = performance.now()
    const spans = findMatches(editor, '*dog*own*', OFF)
    expect(performance.now() - t0).toBeLessThan(100)
    expect(spans.length).toBeGreaterThan(0)
    editor.destroy()
  })
})

describe('foldDiacritics', () => {
  it('strips accents without changing length', () => {
    expect(foldDiacritics('café')).toBe('cafe')
    expect(foldDiacritics('Müller')).toBe('Muller')
    expect(foldDiacritics('naïve')).toBe('naive')
    expect(foldDiacritics('café')).toHaveLength(4)
  })

  it('keeps length for chars it cannot fold, so offsets never shift', () => {
    expect(foldDiacritics('日本語')).toBe('日本語')
    expect(foldDiacritics('가나')).toHaveLength(2) // precomposed Hangul stays
    expect(foldDiacritics('e\u0301')).toHaveLength(2) // lone combining mark stays
    expect(foldDiacritics('🙂é')).toBe('🙂e') // astral chars stay astral
  })
})

describe('ignore-diacritics findMatches', () => {
  const DIA = { matchCase: false, wholeWord: false, useWildcards: false, ignoreDiacritics: true }

  it('cafe matches café and càfe (French corpus)', () => {
    const editor = createEditor('café cafe càfe')
    expect(matchTexts(editor, 'cafe', DIA)).toEqual(['café', 'cafe', 'càfe'])
    editor.destroy()
  })

  it('Muller matches Müller with exact offsets (German corpus)', () => {
    const editor = createEditor('Herr Müller')
    const [m] = findMatches(editor, 'Muller', DIA)
    expect(m).toBeDefined()
    expect(editor.state.doc.textBetween(m.from, m.to)).toBe('Müller')
    editor.destroy()
  })

  it('naive matches naïve and stays case-insensitive', () => {
    const editor = createEditor('CAFÉ naïve café')
    expect(matchTexts(editor, 'cafe', DIA)).toEqual(['CAFÉ', 'café'])
    expect(matchTexts(editor, 'naive', DIA)).toEqual(['naïve'])
    editor.destroy()
  })

  it('accents are distinguished when the option is off or matchCase is on', () => {
    const editor = createEditor('café cafe')
    expect(matchTexts(editor, 'cafe', { ...DIA, ignoreDiacritics: false })).toEqual(['cafe'])
    // case-sensitive search keeps Word behavior: diacritics always differ
    expect(matchTexts(editor, 'cafe', { ...DIA, matchCase: true })).toEqual(['cafe'])
    editor.destroy()
  })

  it('combines with wildcards', () => {
    const editor = createEditor('Müller Muller Mülle')
    const opts = { ...DIA, useWildcards: true }
    expect(matchTexts(editor, 'M?ller', opts)).toEqual(['Müller', 'Muller'])
    expect(matchTexts(editor, 'M?ll*r', opts)).toEqual(['Müller', 'Muller'])
    // the lazy * keeps each match at the M?lle prefix, folded accents and all
    expect(matchTexts(editor, 'M?lle*', opts)).toEqual(['Mülle', 'Mulle', 'Mülle'])
    editor.destroy()
  })
})
