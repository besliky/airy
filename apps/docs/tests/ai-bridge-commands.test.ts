import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'

import { createBridgeCommandHandler } from '../src/renderer/ai/bridge-commands'
import { editorExtensions } from '../src/renderer/editor/extensions'

/**
 * Live-bridge command handlers (src/renderer/ai/bridge-commands.ts): the
 * bridge reuses the agent pipeline, so these tests drive a real editor through
 * get_context / insert_content / apply_ops / undo and assert the bridge-undo
 * turn semantics (exactly one turn, refused when the user edited since).
 */

const liveEditors: Editor[] = []

afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

function makeEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: null },
          content: [{ type: 'text', text: 'Body paragraph' }],
        },
      ],
    },
  })
  liveEditors.push(editor)
  return editor
}

function makeHandler(editor: Editor, withDoc = true) {
  return createBridgeCommandHandler({
    getEditor: () => editor,
    getDocState: () =>
      withDoc ? { blocks: [], isBlank: false, filePath: '/tmp/report.docx' } : null,
    isTrackChangesOn: () => false,
  })
}

describe('bridge command handler', () => {
  it('answers unknown_method', async () => {
    const handler = makeHandler(makeEditor())
    expect(await handler('frobnicate', {})).toEqual({
      ok: false,
      error: { code: 'unknown_method', message: 'unknown bridge method "frobnicate"' },
    })
  })

  it('answers no_active_document when the tab has no document', async () => {
    const handler = makeHandler(makeEditor(), false)
    expect(await handler('get_context', {})).toEqual({
      ok: false,
      error: { code: 'no_active_document', message: 'this tab has no open document' },
    })
  })

  it('returns context with the selection scope and file path', async () => {
    const handler = makeHandler(makeEditor())
    const reply = await handler('get_context', {})
    expect(reply.ok).toBe(true)
    if (!reply.ok) return
    const result = reply.result as { context: string; filePath: string }
    expect(result.filePath).toBe('/tmp/report.docx')
    expect(result.context).toContain('Document block list')
    expect(result.context).toContain('Body paragraph')
  })

  it('inserts at the end of the document by default (not at the cursor)', async () => {
    // two blocks with the selection in the first: the bridge contract
    // (live_apply_ops documents "html inserted at the end") must win over
    // the embedded pipeline's cursor default
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: null },
            content: [{ type: 'text', text: 'First' }],
          },
          {
            type: 'docParagraph',
            attrs: { docxIndex: null },
            content: [{ type: 'text', text: 'Second' }],
          },
        ],
      },
    })
    liveEditors.push(editor)
    const handler = makeHandler(editor)
    const reply = await handler('insert_content', { html: '<p>Inserted by bridge</p>' })
    expect(reply.ok).toBe(true)
    expect(editor.state.doc.childCount).toBe(3)
    expect(editor.state.doc.child(0).textContent).toBe('First')
    expect(editor.state.doc.child(1).textContent).toBe('Second')
    expect(editor.state.doc.child(2).textContent).toBe('Inserted by bridge')
  })

  it('inserts content and undo reverts exactly that turn', async () => {
    const editor = makeEditor()
    const handler = makeHandler(editor)
    expect(editor.state.doc.childCount).toBe(1)

    const insert = await handler('insert_content', {
      html: '<p>Inserted by bridge</p>',
      afterBlockIndex: 0,
    })
    expect(insert.ok).toBe(true)
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(1).textContent).toBe('Inserted by bridge')

    const undo = await handler('undo', {})
    expect(undo).toEqual({ ok: true, result: { undone: true } })
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('Body paragraph')

    expect(await handler('undo', {})).toEqual({
      ok: false,
      error: { code: 'nothing_to_undo', message: 'no bridge turn to undo yet' },
    })
  })

  it('applies ops as one turn and undoes them', async () => {
    const editor = makeEditor()
    const handler = makeHandler(editor)
    const reply = await handler('apply_ops', {
      ops: [{ op: 'setFont', target: { blockIndexes: [0] }, bold: true }],
    })
    expect(reply.ok).toBe(true)
    const paragraph = editor.state.doc.child(0)
    expect(paragraph.firstChild?.marks.some((m) => m.type.name === 'bold')).toBe(true)

    await handler('undo', {})
    expect(editor.state.doc.child(0).firstChild?.marks.some((m) => m.type.name === 'bold')).toBe(
      false,
    )
  })

  it('refuses undo when the user edited after the bridge turn', async () => {
    const editor = makeEditor()
    const handler = makeHandler(editor)
    await handler('insert_content', { html: '<p>Bridge text</p>', afterBlockIndex: 0 })
    // a user keystroke lands between the bridge turn and its undo
    editor.view.dispatch(editor.state.tr.insertText('!', 1))
    const undo = await handler('undo', {})
    expect(undo).toEqual({
      ok: false,
      error: {
        code: 'stale_document',
        message:
          'the document changed since the last bridge turn; undoing would discard those edits — fetch fresh context before editing',
      },
    })
  })

  it('guards index-addressed writes with the stale-document baseline', async () => {
    const editor = makeEditor()
    const handler = makeHandler(editor)
    await handler('get_context', {})
    // user edits after the context read: block indexes may have shifted
    editor.view.dispatch(editor.state.tr.insertText('!', 1))
    const reply = await handler('apply_ops', {
      ops: [{ op: 'setFont', target: { blockIndexes: [0] }, bold: true }],
    })
    expect(reply.ok).toBe(false)
    if (!reply.ok) expect(reply.error.code).toBe('stale_document')
    // a fresh context read re-establishes the baseline and the write goes through
    await handler('get_context', {})
    const retry = await handler('apply_ops', {
      ops: [{ op: 'setFont', target: { blockIndexes: [0] }, bold: true }],
    })
    expect(retry.ok).toBe(true)
  })

  it('maps tool failures to invalid_params', async () => {
    const handler = makeHandler(makeEditor())
    const reply = await handler('insert_content', { html: '<p></p>' })
    expect(reply.ok).toBe(false)
    if (!reply.ok) expect(reply.error.code).toBe('invalid_params')
  })
})
