import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { compareWithFile, type ReviewContext } from '../src/renderer/review-actions'
import { t, setModuleLang } from '../src/renderer/i18n/locale'
import type { DocState } from '../src/renderer/doc-state'
import type { PmNode } from '../src/renderer/editor/convert'

// compareWithFile reaches the shell file picker through window.desktop
const openDocx = vi.fn()
Object.assign(window, { desktop: { openDocx } })

setModuleLang('en')

const para = (text: string, marks?: PmNode['marks']): PmNode => ({
  type: 'docParagraph',
  ...(text ? { content: [{ type: 'text', text, ...(marks ? { marks } : {}) }] } : {}),
})

function createEditor(content: PmNode[], editable = true): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
    editable,
  })
}

/** minimal ReviewContext covering the fields compareWithFile touches */
function makeCtx(editor: Editor | null, parsedBlocks: unknown[] = []) {
  const status: string[] = []
  const ctx = {
    editor,
    doc: { parsed: { blocks: parsedBlocks } } as unknown as DocState,
    dirtyRef: { current: false },
    setStatus: (value: string) => status.push(value),
    setCompareResult: vi.fn(),
    setRevisionDisplay: vi.fn(),
  } as unknown as ReviewContext
  return { ctx, status }
}

describe('compareWithFile guards (BUG-915 pending revisions, UX-903 read-only editor)', () => {
  beforeAll(async () => {
    const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>New text</w:t></w:r></w:p>' })
    openDocx.mockResolvedValue({ name: 'other.docx', data: bytes })
  })

  beforeEach(() => {
    openDocx.mockClear()
  })

  it('merges into an editable document without pending revisions (positive control)', async () => {
    const editor = createEditor([para('Original text')])
    const dispatch = vi.spyOn(editor.view, 'dispatch')
    const { ctx, status } = makeCtx(editor)
    await compareWithFile(ctx, 'merge')
    expect(openDocx).toHaveBeenCalledTimes(1)
    expect(dispatch).toHaveBeenCalledTimes(1)
    expect(ctx.dirtyRef.current).toBe(true)
    expect(status.at(-1)).toBe(
      t('reviewCompareMerged', { name: 'other.docx', added: 0, removed: 0, changed: 1 }),
    )
    editor.destroy()
  })

  it('refuses the merge on a read-only editor: zero transactions, no picker (UX-903)', async () => {
    const editor = createEditor([para('Original text')], false)
    expect(editor.isEditable).toBe(false)
    const dispatch = vi.spyOn(editor.view, 'dispatch')
    const { ctx, status } = makeCtx(editor)
    await compareWithFile(ctx, 'merge')
    expect(openDocx).not.toHaveBeenCalled()
    expect(dispatch).not.toHaveBeenCalled()
    expect(ctx.dirtyRef.current).toBe(false)
    expect(status.at(-1)).toBe(t('reviewCompareReadonly'))
    editor.destroy()
  })

  it('refuses the merge over pending revisions instead of mixing stamps (BUG-915)', async () => {
    const editor = createEditor([
      para('pending edit', [
        { type: 'ins', attrs: { author: 'Other', date: '2020-01-01T00:00:00Z', id: null } },
      ]),
    ])
    const dispatch = vi.spyOn(editor.view, 'dispatch')
    const { ctx, status } = makeCtx(editor)
    await compareWithFile(ctx, 'merge')
    expect(openDocx).not.toHaveBeenCalled()
    expect(dispatch).not.toHaveBeenCalled()
    expect(ctx.dirtyRef.current).toBe(false)
    expect(status.at(-1)).toBe(t('reviewComparePendingRevisions'))
    editor.destroy()
  })

  it('keeps the read-only-safe differences pane available in a read-only document (UX-903)', async () => {
    const editor = createEditor([para('Original text')], false)
    const dispatch = vi.spyOn(editor.view, 'dispatch')
    const { ctx } = makeCtx(editor, [{ runs: [{ text: 'Original text' }] }])
    await compareWithFile(ctx, 'panel')
    expect(openDocx).toHaveBeenCalledTimes(1)
    expect(dispatch).not.toHaveBeenCalled()
    expect(ctx.setCompareResult).toHaveBeenCalledTimes(1)
    editor.destroy()
  })

  it('warns when the compared documents exceed the paragraph budget (BUG-913 panel path)', async () => {
    // (2200+1)^2 = 4.84M cells > the 4M paragraph budget
    const paras = Array.from({ length: 2200 }, (_, i) => `<w:p><w:r><w:t>p ${i}</w:t></w:r></w:p>`)
    const bytes = await buildDocx({ bodyXml: paras.join('') })
    openDocx.mockResolvedValueOnce({ name: 'big.docx', data: bytes })
    const blocks = Array.from({ length: 2200 }, (_, i) => ({ runs: [{ text: `q ${i}` }] }))
    const { ctx, status } = makeCtx(null, blocks)
    await compareWithFile(ctx, 'panel')
    expect(ctx.setCompareResult).toHaveBeenCalledTimes(1)
    expect(status.at(-1)).toBe(t('reviewCompareDegraded'))
  })
})
