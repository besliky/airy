import { afterEach, describe, expect, it } from 'vitest'
import type { Editor } from '@tiptap/core'

import { createTrackedEditor, drainTrackedEditors } from './helpers/tracked-editor'
import { createBridgeCommandHandler } from '../src/renderer/ai/bridge-commands'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setModuleLang } from '../src/renderer/i18n/locale'

/**
 * The MCP ops guide documents the headless nodeType vocabulary
 * (heading/paragraph/listItem); the live bridge forwards agent ops verbatim
 * into the embedded registry, which historically only accepted the canonical
 * names (docHeading/docParagraph/docListItem). These tests drive the exact
 * bridge validation path (createBridgeCommandHandler -> executeTool ->
 * executeOps) with the ops-guide aliases and assert they are accepted and
 * behave identically to the canonical spelling.
 */

setModuleLang('en')

afterEach(() => drainTrackedEditors())

function makeEditor(): Editor {
  return createTrackedEditor({
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docHeading',
          attrs: { docxIndex: null, level: 1 },
          content: [{ type: 'text', text: 'Chapter title' }],
        },
        {
          type: 'docParagraph',
          attrs: { docxIndex: null },
          content: [{ type: 'text', text: 'Body paragraph' }],
        },
        {
          type: 'docListItem',
          attrs: { docxIndex: null, kind: 'bullet', numId: '1' },
          content: [{ type: 'text', text: 'List item' }],
        },
      ],
    },
  })
}

function makeHandler(editor: Editor) {
  return createBridgeCommandHandler({
    getEditor: () => editor,
    getDocState: () => ({ blocks: [], isBlank: false, filePath: '/tmp/report.docx' }),
    isTrackChangesOn: () => false,
  })
}

/** apply one op batch over the bridge (same path shell-bridge.ts forwards to) */
async function applyOps(
  editor: Editor,
  ops: Array<Record<string, unknown>>,
): Promise<{ ok: boolean; error?: { code: string; message: string } }> {
  return makeHandler(editor)('apply_ops', { ops })
}

const markOf = (editor: Editor, block: number, mark: string) =>
  editor.state.doc.child(block).firstChild?.marks.some((m) => m.type.name === mark) ?? false

describe('ops-guide nodeType aliases over the bridge', () => {
  it('accepts heading/paragraph/listItem like the canonical names', async () => {
    const aliasEditor = makeEditor()
    const canonicalEditor = makeEditor()

    const alias = await applyOps(aliasEditor, [
      { op: 'setFont', target: { nodeType: 'heading' }, italic: true },
    ])
    const canonical = await applyOps(canonicalEditor, [
      { op: 'setFont', target: { nodeType: 'docHeading' }, italic: true },
    ])
    expect(alias.ok).toBe(true)
    expect(canonical.ok).toBe(true)

    // identical behavior: both styled exactly the heading
    for (const editor of [aliasEditor, canonicalEditor]) {
      expect(markOf(editor, 0, 'italic')).toBe(true)
      expect(markOf(editor, 1, 'italic')).toBe(false)
      expect(markOf(editor, 2, 'italic')).toBe(false)
    }
  })

  it('aliases drive structural ops on the right blocks', async () => {
    const editor = makeEditor()
    const reply = await applyOps(editor, [
      { op: 'setHeadingLevel', target: { nodeType: 'paragraph' }, level: 2 },
      { op: 'clearList', target: { nodeType: 'listItem' } },
    ])
    expect(reply.ok).toBe(true)
    // the paragraph became a heading; the list item became a paragraph
    expect(editor.state.doc.child(1).type.name).toBe('docHeading')
    expect(editor.state.doc.child(2).type.name).toBe('docParagraph')
    // the original heading is untouched
    expect(editor.state.doc.child(0).type.name).toBe('docHeading')
    expect(editor.state.doc.child(0).attrs.level).toBe(1)
  })

  it('aliases combine with other target conditions', async () => {
    const editor = makeEditor()
    const reply = await applyOps(editor, [
      { op: 'setFont', target: { nodeType: 'heading', headingLevel: 1 }, bold: true },
    ])
    expect(reply.ok).toBe(true)
    expect(markOf(editor, 0, 'bold')).toBe(true)
  })

  it('still rejects genuinely unknown nodeTypes with a helpful message', async () => {
    const editor = makeEditor()
    const reply = await applyOps(editor, [
      { op: 'setFont', target: { nodeType: 'secshun' }, bold: true },
    ])
    expect(reply.ok).toBe(false)
    expect(reply.error?.message).toContain('unknown nodeType "secshun"')
    expect(reply.error?.message).toContain('heading/paragraph/listItem')
  })
})
