import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { runsToInline, type PmNode } from '../src/renderer/editor/convert'
import {
  docNoteCustomMark,
  docNoteMark,
  formatNoteMark,
  noteMarkerOf,
  setDocNoteNumbering,
  toLetter,
} from '../src/renderer/note-format'
import {
  convertNotes,
  navigateNote,
  noteRefAtSelection,
  refreshNoteMarks,
  type ReviewContext,
} from '../src/renderer/review-actions'
import type { NoteInfo } from '@airy-office/docx-engine'

function createEditor(content: PmNode[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

describe('note numbering formats', () => {
  it('formats 1/a/A/i/I and wraps letters past z', () => {
    expect(formatNoteMark('decimal', 4)).toBe('4')
    expect(formatNoteMark('lowerLetter', 1)).toBe('a')
    expect(formatNoteMark('upperLetter', 2)).toBe('B')
    expect(formatNoteMark('lowerRoman', 14)).toBe('xiv')
    expect(formatNoteMark('upperRoman', 9)).toBe('IX')
    expect(toLetter(26)).toBe('z')
    expect(toLetter(27)).toBe('aa')
    expect(toLetter(53, true)).toBe('BA')
  })

  it('custom mark replaces numbers; numStart shifts the first value', () => {
    expect(noteMarkerOf({ numFmt: 'decimal', customMark: '†' }, 7)).toBe('†')
    expect(noteMarkerOf({ numFmt: 'decimal', numStart: 5 }, 1)).toBe('5')
    expect(noteMarkerOf(undefined, 3)).toBe('3')
  })

  it('module state drives docNoteMark; endnotes default to lowercase roman', () => {
    setDocNoteNumbering(undefined)
    expect(docNoteMark('footnote', 3)).toBe('3')
    expect(docNoteMark('endnote', 3)).toBe('iii')
    setDocNoteNumbering({ footnotes: { numFmt: 'upperLetter' }, endnotes: { numFmt: 'decimal' } })
    expect(docNoteMark('footnote', 3)).toBe('C')
    expect(docNoteMark('endnote', 3)).toBe('3')
    expect(docNoteCustomMark('footnote')).toBeNull()
    setDocNoteNumbering({ footnotes: { numFmt: 'decimal', customMark: '*' } })
    expect(docNoteCustomMark('footnote')).toBe('*')
    setDocNoteNumbering(undefined)
  })

  it('runsToInline formats marks under the current options', () => {
    setDocNoteNumbering({ footnotes: { numFmt: 'lowerRoman' } })
    const inline = runsToInline([{ text: '4', noteRef: { kind: 'footnote', id: '9' } }])
    expect(inline[0]!.attrs).toMatchObject({ num: 4, mark: 'iv' })
    setDocNoteNumbering({ footnotes: { numFmt: 'decimal', customMark: '§' } })
    const custom = runsToInline([{ text: '4', noteRef: { kind: 'footnote', id: '9' } }])
    expect(custom[0]!.attrs).toMatchObject({ customMark: '§', mark: '§' })
    setDocNoteNumbering(undefined)
  })

  it('in-text reference renders the formatted mark', () => {
    setDocNoteNumbering({ footnotes: { numFmt: 'upperRoman' } })
    const editor = createEditor([
      {
        type: 'docParagraph',
        content: [
          { type: 'text', text: 'x' },
          {
            type: 'docNoteRef',
            attrs: { kind: 'footnote', id: '1', num: 2, customMark: null, mark: 'II' },
          },
        ],
      },
    ])
    const sup = editor.view.dom.querySelector('sup[data-note-ref]')
    expect(sup?.textContent).toBe('II')
    editor.destroy()
    setDocNoteNumbering(undefined)
  })
})

function reviewCtx(editor: Editor, footnotes: NoteInfo[], endnotes: NoteInfo[]): ReviewContext {
  return {
    editor,
    doc: null,
    dirtyRef: { current: false },
    setStatus: () => {},
    notePrompt: null,
    setNotePrompt: () => {},
    footnotes,
    endnotes,
    setFootnotes: (next: NoteInfo[] | ((prev: NoteInfo[]) => NoteInfo[])) => {
      const value = typeof next === 'function' ? next(footnotes) : next
      footnotes.splice(0, footnotes.length, ...value)
    },
    setEndnotes: (next: NoteInfo[] | ((prev: NoteInfo[]) => NoteInfo[])) => {
      const value = typeof next === 'function' ? next(endnotes) : next
      endnotes.splice(0, endnotes.length, ...value)
    },
    setNotesDirty: () => {},
    comments: [],
    setComments: () => {},
    setCommentsDirty: () => {},
    setCommentComposing: () => {},
    setShowComments: () => {},
    setInkAnnotations: () => {},
    setInksDirty: () => {},
    setCompareResult: () => {},
    setRevisionDisplay: () => {},
  }
}

describe('note options: conversion + navigation (review-actions)', () => {
  const DOC = (extra?: PmNode[]): PmNode[] => [
    {
      type: 'docParagraph',
      content: [
        { type: 'text', text: 'a' },
        {
          type: 'docNoteRef',
          attrs: { kind: 'footnote', id: '1', num: 1, customMark: null, mark: '1' },
        },
        { type: 'text', text: 'b' },
        {
          type: 'docNoteRef',
          attrs: { kind: 'endnote', id: '2', num: 1, customMark: null, mark: 'i' },
        },
        { type: 'text', text: 'c' },
        {
          type: 'docNoteRef',
          attrs: { kind: 'footnote', id: '3', num: 2, customMark: null, mark: '2' },
        },
      ],
    },
    ...(extra ?? []),
  ]

  const refAttrs = (editor: Editor) =>
    Array.from(editor.view.dom.querySelectorAll('sup[data-note-ref]')).map((s) => ({
      kind: s.getAttribute('data-note-kind'),
      text: s.textContent,
    }))

  it('converts every footnote to an endnote: nodes switch, lists rebuild in doc order', () => {
    const editor = createEditor(DOC())
    const footnotes: NoteInfo[] = [
      { id: '1', text: 'one' },
      { id: '3', text: 'three' },
    ]
    const endnotes: NoteInfo[] = [{ id: '2', text: 'two' }]
    convertNotes(reviewCtx(editor, footnotes, endnotes), 'footnote', 'all')
    // reference order: fn1→en#1, en2→en#2, fn3→en#3
    expect(refAttrs(editor)).toEqual([
      { kind: 'endnote', text: 'i' },
      { kind: 'endnote', text: 'ii' },
      { kind: 'endnote', text: 'iii' },
    ])
    expect(footnotes).toEqual([])
    expect(endnotes.map((n) => n.id)).toEqual(['1', '2', '3'])
    editor.destroy()
  })

  it('converts one note: target numbering merges at its reference position', () => {
    const editor = createEditor(DOC())
    const footnotes: NoteInfo[] = [
      { id: '1', text: 'one' },
      { id: '3', text: 'three' },
    ]
    const endnotes: NoteInfo[] = [{ id: '2', text: 'two' }]
    convertNotes(reviewCtx(editor, footnotes, endnotes), 'footnote', '3')
    expect(refAttrs(editor)).toEqual([
      { kind: 'footnote', text: '1' },
      { kind: 'endnote', text: 'i' },
      { kind: 'endnote', text: 'ii' },
    ])
    expect(footnotes.map((n) => n.id)).toEqual(['1'])
    expect(endnotes.map((n) => n.id)).toEqual(['2', '3'])
    editor.destroy()
  })

  it('refreshNoteMarks re-renders markers after the options change', () => {
    const editor = createEditor(DOC())
    setDocNoteNumbering({
      footnotes: { numFmt: 'lowerLetter' },
      endnotes: { numFmt: 'upperRoman' },
    })
    refreshNoteMarks(editor)
    expect(refAttrs(editor)).toEqual([
      { kind: 'footnote', text: 'a' },
      { kind: 'endnote', text: 'I' },
      { kind: 'footnote', text: 'b' },
    ])
    setDocNoteNumbering(undefined)
    editor.destroy()
  })

  it('navigateNote selects the next/previous reference and wraps', () => {
    const editor = createEditor(DOC())
    // start before everything → first
    navigateNote(reviewCtx(editor, [], []), 1)
    expect(noteRefAtSelection(editor)?.id).toBe('1')
    navigateNote(reviewCtx(editor, [], []), 1)
    expect(noteRefAtSelection(editor)?.id).toBe('2')
    navigateNote(reviewCtx(editor, [], []), -1)
    expect(noteRefAtSelection(editor)?.id).toBe('1')
    // wrap backwards from the first → last
    navigateNote(reviewCtx(editor, [], []), -1)
    expect(noteRefAtSelection(editor)?.id).toBe('3')
    editor.destroy()
  })

  it('noteRefAtSelection returns null on plain text', () => {
    const editor = createEditor([
      { type: 'docParagraph', content: [{ type: 'text', text: 'plain' }] },
    ])
    expect(noteRefAtSelection(editor)).toBeNull()
    editor.destroy()
  })
})
