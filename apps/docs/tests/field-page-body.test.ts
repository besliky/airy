import type { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import {
  collectInlineFieldJobs,
  fieldKeyword,
  topLevelBlockPos,
  type InlineFieldResolvers,
} from '../src/renderer/editor/field-caches'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyFieldCaches } from '../src/renderer/editor/revisions'
import { createTrackedEditor, drainTrackedEditors } from './helpers/tracked-editor'

afterEach(() => drainTrackedEditors())

/** positions of the top-level blocks, in document order */
function topLevelPositions(editor: Editor): number[] {
  const out: number[] = []
  editor.state.doc.forEach((_node, offset) => out.push(offset))
  return out
}

/** position of the first text node carrying an instrField mark */
function instrFieldPos(editor: Editor): number {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found < 0 && node.isText && node.marks.some((m) => m.type.name === 'instrField')) {
      found = pos
    }
  })
  expect(found).toBeGreaterThan(-1)
  return found
}

/** three paragraphs, each ending in a body PAGE field whose author cache is "1"
 * (the shape a document opened at the top saves) */
function pageFieldEditor() {
  const editor = createTrackedEditor({
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [1, 2, 3].map(() => ({
        type: 'docParagraph',
        content: [
          { type: 'text', text: 'Page ' },
          { type: 'text', text: '1', marks: [{ type: 'instrField', attrs: { instr: ' PAGE ' } }] },
        ],
      })),
    },
  })
  return { editor, paraPositions: topLevelPositions(editor) }
}

const noRefCache: InlineFieldResolvers['refCache'] = () => null

describe('BUG-1709: body PAGE fields resolve from the live page slicing', () => {
  it('fieldKeyword reads the first word, switches included', () => {
    expect(fieldKeyword(' PAGE ')).toBe('PAGE')
    expect(fieldKeyword('page \\* MERGEFORMAT')).toBe('PAGE')
    expect(fieldKeyword('REF _Toc1 \\h')).toBe('REF')
    expect(fieldKeyword('')).toBe('')
  })

  it('topLevelBlockPos lifts a nested position to its top-level block', () => {
    const { editor, paraPositions } = pageFieldEditor()
    expect(topLevelBlockPos(editor.state.doc, instrFieldPos(editor))).toBe(paraPositions[0])
    // depth-0 positions have no enclosing block
    expect(topLevelBlockPos(editor.state.doc, 0)).toBeNull()
  })

  it('F9 writes the page of the field block, not one value for the whole body', () => {
    const { editor, paraPositions } = pageFieldEditor()
    const pageByBlock = new Map([
      [paraPositions[0], 1],
      [paraPositions[1], 3],
      [paraPositions[2], 7],
    ])
    const jobs = collectInlineFieldJobs(editor, {
      fieldValue: () => '',
      pageOf: () => (pos) => pageByBlock.get(pos) ?? null,
      refCache: noRefCache,
    })
    // the page-1 field already shows its page: no job for it
    expect(jobs.map((j) => j.text)).toEqual(['3', '7'])
    applyFieldCaches(editor, jobs)
    const paraTexts: string[] = []
    editor.state.doc.forEach((node) => paraTexts.push(node.textContent))
    expect(paraTexts).toEqual(['Page 1', 'Page 3', 'Page 7'])
  })

  it('keeps the cached result when pagination cannot place the block', () => {
    const { editor } = pageFieldEditor()
    const jobs = collectInlineFieldJobs(editor, {
      fieldValue: () => '',
      pageOf: () => null,
      refCache: noRefCache,
    })
    // the viewport page ("1" for a document opened at the top) used to be
    // stamped into every field here — writing a wrong cache is worse than a
    // stale one, so nothing is written
    expect(jobs).toEqual([])
  })

  it('measures the canvas once per update, and only when a field needs it', () => {
    const { editor } = pageFieldEditor()
    let factoryCalls = 0
    const resolvers: InlineFieldResolvers = {
      fieldValue: () => '',
      pageOf: () => {
        factoryCalls += 1
        return () => 2
      },
      refCache: noRefCache,
    }
    collectInlineFieldJobs(editor, resolvers)
    // three PAGE fields, one measurement
    expect(factoryCalls).toBe(1)
  })

  it('PAGE fields inside a table resolve from the table block', () => {
    const editor = createTrackedEditor({
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [
                  {
                    type: 'docTableCell',
                    content: [
                      {
                        type: 'docParagraph',
                        content: [
                          {
                            type: 'text',
                            text: '1',
                            marks: [{ type: 'instrField', attrs: { instr: 'PAGE' } }],
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    })
    const tablePos = topLevelPositions(editor)[0]
    const jobs = collectInlineFieldJobs(editor, {
      fieldValue: () => '',
      pageOf: () => (pos) => (pos === tablePos ? 5 : null),
      refCache: noRefCache,
    })
    expect(jobs.map((j) => j.text)).toEqual(['5'])
  })

  it('position-independent fields keep the fieldValue resolver; REF \\p alone gets the page lookup', () => {
    const editor = createTrackedEditor({
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: '1',
                marks: [{ type: 'instrField', attrs: { instr: ' NUMPAGES ' } }],
              },
              {
                type: 'text',
                text: 'x',
                marks: [{ type: 'refField', attrs: { name: 'Book', instr: ' REF Book \\p \\h ' } }],
              },
              { type: 'text', text: 'y', marks: [{ type: 'refField', attrs: { name: 'Plain' } }] },
            ],
          },
        ],
      },
    })
    const refCalls: Array<[string, boolean]> = []
    const jobs = collectInlineFieldJobs(editor, {
      fieldValue: (instr) => (fieldKeyword(instr) === 'NUMPAGES' ? '12' : ''),
      pageOf: () => () => 4,
      refCache: (instr, pageOf) => {
        refCalls.push([instr, pageOf !== null])
        return null
      },
    })
    // NUMPAGES refreshed to the document total; the \p REF would get its page
    // (null refCache keeps it here), the plain REF never asks for one
    expect(jobs.map((j) => j.text)).toEqual(['12'])
    expect(refCalls).toEqual([
      [' REF Book \\p \\h ', true],
      [' REF Plain \\h ', false],
    ])
  })
})
