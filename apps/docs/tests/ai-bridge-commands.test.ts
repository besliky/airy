import { afterEach, describe, expect, it } from 'vitest'
import type { Editor } from '@tiptap/core'

import { createTrackedEditor, drainTrackedEditors } from './helpers/tracked-editor'
import { createBridgeCommandHandler } from '../src/renderer/ai/bridge-commands'
import { editorExtensions } from '../src/renderer/editor/extensions'

/**
 * Live-bridge command handlers (src/renderer/ai/bridge-commands.ts): the
 * bridge reuses the agent pipeline, so these tests drive a real editor through
 * get_context / insert_content / apply_ops / undo and assert the bridge-undo
 * turn semantics (exactly one turn, refused when the user edited since).
 */

afterEach(() => drainTrackedEditors())

function makeEditor(): Editor {
  return createTrackedEditor({
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
    const editor = createTrackedEditor({
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

  it('a combined live_apply_ops takes exactly two undos (BUG-1505)', async () => {
    // the MCP server sends a combined live_apply_ops as TWO bridge calls —
    // insert_content first, then apply_ops. A single undo slot let the ops
    // turn overwrite the insert turn: the first undo reverted only the ops,
    // the second answered nothing_to_undo and the insert was stranded.
    const editor = makeEditor()
    const handler = makeHandler(editor)

    // insert turn (html at the end, as the live_apply_ops contract says)
    const insert = await handler('insert_content', { html: '<p>LIVE-EDITED ITEM</p>' }, 'conn-1')
    expect(insert.ok).toBe(true)
    // ops turn touching both the original and the inserted block
    const ops = await handler(
      'apply_ops',
      {
        ops: [
          { op: 'setFont', target: { nodeType: 'paragraph' }, bold: true },
          { op: 'findReplace', find: 'LIVE-EDITED', replace: 'DONE' },
        ],
      },
      'conn-1',
    )
    expect(ops.ok).toBe(true)
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(1).textContent).toBe('DONE ITEM')

    // first undo reverts the ops turn only (the insert survives)
    const first = await handler('undo', {}, 'conn-1')
    expect(first).toEqual({ ok: true, result: { undone: true } })
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(1).textContent).toBe('LIVE-EDITED ITEM')
    expect(editor.state.doc.child(1).firstChild?.marks.some((m) => m.type.name === 'bold')).toBe(
      false,
    )

    // second undo reverts the insert turn — the document is back to base
    const second = await handler('undo', {}, 'conn-1')
    expect(second).toEqual({ ok: true, result: { undone: true } })
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('Body paragraph')

    // and the stack is empty again
    expect(await handler('undo', {}, 'conn-1')).toEqual({
      ok: false,
      error: { code: 'nothing_to_undo', message: 'no bridge turn to undo yet' },
    })
  })

  it('a failed ops turn after an insert still rolls back with one guarded undo', async () => {
    // the MCP auto-rollback path: insert ok, ops batch fails -> one
    // ownTurnsOnly undo must take the document back to the pre-call state
    const editor = makeEditor()
    const handler = makeHandler(editor)
    await handler('insert_content', { html: '<p>Inserted</p>' }, 'conn-1')
    const failed = await handler(
      'apply_ops',
      { ops: [{ op: 'setFont', target: { blockIndexes: [0] }, bogus: 1 }] },
      'conn-1',
    )
    expect(failed.ok).toBe(false)
    const rollback = await handler('undo', { ownTurnsOnly: true }, 'conn-1')
    expect(rollback).toEqual({ ok: true, result: { undone: true } })
    expect(editor.state.doc.childCount).toBe(1)
    expect(await handler('undo', {}, 'conn-1')).toMatchObject({
      ok: false,
      error: { code: 'nothing_to_undo' },
    })
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

describe('bridge turn ownership (multiple copilot clients)', () => {
  const isBold = (editor: Editor) =>
    editor.state.doc.child(0).firstChild?.marks.some((m) => m.type.name === 'bold') ?? false

  it('auto-rollback (ownTurnsOnly) refuses to revert another client turn', async () => {
    const editor = makeEditor()
    const handler = makeHandler(editor)
    // client A inserts; client B then makes the LATEST turn
    await handler('insert_content', { html: '<p>From A</p>', afterBlockIndex: 0 }, 'conn-1')
    await handler(
      'apply_ops',
      { ops: [{ op: 'setFont', target: { blockIndexes: [0] }, bold: true }] },
      'conn-2',
    )
    // A's automatic rollback must not silently revert B's turn
    const guarded = await handler('undo', { ownTurnsOnly: true }, 'conn-1')
    expect(guarded).toMatchObject({ ok: false, error: { code: 'turn_owned_by_other' } })
    if (!guarded.ok) expect(guarded.error.message).toContain('conn-2')
    expect(isBold(editor)).toBe(true)
    // B's own guarded undo reverts B's turn (A's insert survives)
    const own = await handler('undo', { ownTurnsOnly: true }, 'conn-2')
    expect(own).toEqual({ ok: true, result: { undone: true } })
    expect(isBold(editor)).toBe(false)
    expect(editor.state.doc.child(1).textContent).toBe('From A')
  })

  it('explicit undo may revert another client turn and reports whose it was', async () => {
    const editor = makeEditor()
    const handler = makeHandler(editor)
    await handler('insert_content', { html: '<p>From B</p>', afterBlockIndex: 0 }, 'conn-2')
    // live_undo is the user's explicit choice: allowed across clients, and
    // the result says whose turn was reverted
    const explicit = await handler('undo', {}, 'conn-1')
    expect(explicit).toEqual({
      ok: true,
      result: { undone: true, revertedTurnOf: 'conn-2', anotherClient: true },
    })
    expect(editor.state.doc.childCount).toBe(1)
    // undoing the requester's own turn is not flagged as another client's
    await handler('insert_content', { html: '<p>x</p>', afterBlockIndex: 0 }, 'conn-1')
    const own = await handler('undo', {}, 'conn-1')
    expect(own).toEqual({ ok: true, result: { undone: true } })
  })

  it('callers without a connection id share the legacy single-client identity', async () => {
    const editor = makeEditor()
    const handler = makeHandler(editor)
    await handler('insert_content', { html: '<p>legacy</p>', afterBlockIndex: 0 })
    // anonymous auto-rollback of the anonymous turn still works (legacy path)
    const own = await handler('undo', { ownTurnsOnly: true })
    expect(own).toEqual({ ok: true, result: { undone: true } })
    // but a stamped client cannot silently revert the anonymous turn
    await handler('insert_content', { html: '<p>legacy 2</p>', afterBlockIndex: 0 })
    const guarded = await handler('undo', { ownTurnsOnly: true }, 'conn-1')
    expect(guarded).toMatchObject({ ok: false, error: { code: 'turn_owned_by_other' } })
    expect(editor.state.doc.childCount).toBe(2)
  })
})
