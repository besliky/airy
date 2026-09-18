/**
 * "Place in This Document" hyperlinks (PAR-106): the link dialog lists
 * headings (by outline level) and bookmarks, a picked heading gains a hidden
 * `_Toc…` bookmark (re-emitted as w:bookmarkStart on save), the link saves as
 * w:hyperlink w:anchor WITHOUT r:id and reads back as href="#anchor", plain
 * bookmark targets round-trip by name, external links keep their r:id
 * hyperlink untouched, editing an existing internal link opens the document
 * pane with its target picked, and ⌘/Ctrl+click jumps to the resolved target.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import type { Node as PmDocNode } from '@tiptap/pm/model'
import { parseDocx, saveDocx } from '@airy-office/docx-engine'
import JSZip from 'jszip'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import {
  editorExtensions,
  isInternalLinkNavClick,
  resolveInternalLinkTarget,
} from '../src/renderer/editor/extensions'
import {
  collectLinkTargets,
  ensureHeadingTocAnchor,
  headingTocAnchor,
  uniqueTocAnchor,
} from '../src/renderer/components/ribbon-insert-tab'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { t } from '../src/renderer/i18n/locale'
import { ribbonProps } from './helpers/ribbon-props'

const editors = new Set<Editor>()

afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

const DOC_CONTENT = {
  type: 'doc',
  content: [
    {
      type: 'docHeading',
      attrs: { level: 1 },
      content: [{ type: 'text', text: 'Chapter One' }],
    },
    {
      type: 'docHeading',
      attrs: { level: 2 },
      content: [{ type: 'text', text: 'Section A' }],
    },
    {
      type: 'docParagraph',
      attrs: { bookmarks: ['Intro'] },
      content: [{ type: 'text', text: 'Intro paragraph.' }],
    },
    { type: 'docParagraph', content: [{ type: 'text', text: 'plain text' }] },
  ],
}

const makeEditor = (content: object = DOC_CONTENT) => {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: content as never,
  })
  editors.add(editor)
  return editor
}

/** all link mark hrefs in the doc */
const linkHrefsIn = (doc: PmDocNode): string[] => {
  const hrefs: string[] = []
  doc.descendants((node: PmDocNode) => {
    for (const mark of node.marks) {
      if (mark.type.name === 'link') hrefs.push(String(mark.attrs.href))
    }
  })
  return hrefs
}

describe('link target list (headings + bookmarks)', () => {
  it('collects headings with levels and anchors, then bookmarks by name', () => {
    const editor = makeEditor()
    const targets = collectLinkTargets(editor)
    expect(targets.map((tg) => [tg.kind, tg.label, tg.level, tg.anchor])).toEqual([
      ['heading', 'Chapter One', 1, null],
      ['heading', 'Section A', 2, null],
      ['bookmark', 'Intro', 1, 'Intro'],
    ])
  })

  it('reuses the heading existing hidden _Toc bookmark as its anchor', () => {
    const editor = makeEditor({
      type: 'doc',
      content: [
        {
          type: 'docHeading',
          attrs: { level: 1, hiddenBookmarks: ['_Ref5', '_Toc123456789'] },
          content: [{ type: 'text', text: 'Only Heading' }],
        },
      ],
    })
    expect(headingTocAnchor(editor.state.doc.firstChild)).toBe('_Toc123456789')
    expect(collectLinkTargets(editor)[0].anchor).toBe('_Toc123456789')
  })

  it('generates unique Word-style _Toc anchors', () => {
    const taken = new Set(['_Toc111111111', '_Toc222222222'])
    for (let i = 0; i < 20; i++) {
      const name = uniqueTocAnchor(taken)
      expect(name).toMatch(/^_Toc\d{9}$/)
      expect(taken.has(name)).toBe(false)
      taken.add(name)
    }
  })

  it('stamps a heading without an anchor once and returns the same name after', () => {
    const editor = makeEditor()
    const target = collectLinkTargets(editor).find((tg) => tg.label === 'Chapter One')!
    const first = ensureHeadingTocAnchor(editor, target.pos)
    expect(first).toMatch(/^_Toc\d{9}$/)
    const second = ensureHeadingTocAnchor(editor, target.pos)
    expect(second).toBe(first)
    const heading = editor.state.doc.nodeAt(target.pos)!
    expect((heading.attrs.hiddenBookmarks as string[]).includes(first!)).toBe(true)
  })

  it('never stamps an anchor on a read-only document (BUG-743)', () => {
    const editor = makeEditor()
    const target = collectLinkTargets(editor).find((tg) => tg.label === 'Chapter One')!
    editor.setEditable(false)
    const docBefore = editor.state.doc
    // the app menu / context menu can open the link dialog while editing is
    // locked — stamping would mark the document dirty and save a stray
    // w:bookmarkStart into the file
    expect(ensureHeadingTocAnchor(editor, target.pos)).toBeNull()
    expect(editor.state.doc.nodeAt(target.pos)!.attrs.hiddenBookmarks).toBeNull()
    expect(editor.state.doc).toBe(docBefore) // no transaction landed at all
    // once editable again, stamping works
    editor.setEditable(true)
    const name = ensureHeadingTocAnchor(editor, target.pos)
    expect(name).toMatch(/^_Toc\d{9}$/)
    // an already-stamped anchor stays readable in read-only mode
    editor.setEditable(false)
    expect(ensureHeadingTocAnchor(editor, target.pos)).toBe(name)
  })
})

describe('internal hyperlink save round-trip (w:anchor)', () => {
  const BODY =
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Chapter 1</w:t></w:r></w:p>' +
    '<w:p><w:bookmarkStart w:id="1" w:name="Conclusion"/><w:bookmarkEnd w:id="1"/>' +
    '<w:r><w:t>Conclusion paragraph.</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>' +
    '<w:hyperlink r:id="rId7"><w:r><w:t>the site</w:t></w:r></w:hyperlink>' +
    '<w:r><w:t>.</w:t></w:r></w:p>'
  const RELS =
    '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>'

  async function openDoc() {
    const source = await buildDocx({ bodyXml: BODY, extraRels: RELS })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.add(editor)
    return { source, parsed, editor }
  }

  async function docXmlOf(bytes: Uint8Array): Promise<string> {
    const zip = await JSZip.loadAsync(bytes)
    return zip.file('word/document.xml')!.async('string')
  }

  /** insert a "Place in This Document" link exactly like the dialog does */
  const insertInternalLink = (editor: Editor, label: string) => {
    const target = collectLinkTargets(editor).find((tg) => tg.label === label)!
    const anchor =
      target.kind === 'bookmark' ? target.anchor! : ensureHeadingTocAnchor(editor, target.pos)!
    editor
      .chain()
      .focus()
      .setTextSelection(editor.state.doc.content.size - 1)
      .insertContent({
        type: 'text',
        text: label,
        marks: [{ type: 'link', attrs: { href: `#${anchor}`, rId: null } }],
      })
      .run()
    return anchor
  }

  it('saves a heading link as w:hyperlink w:anchor plus a hidden bookmarkStart, and reads it back', async () => {
    const { parsed, editor } = await openDoc()
    const anchor = insertInternalLink(editor, 'Chapter 1')
    expect(linkHrefsIn(editor.state.doc)).toContain(`#${anchor}`)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xml = await docXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(xml).toContain(`<w:hyperlink w:anchor="${anchor}">`)
    // no relationship id on an internal link
    expect(xml).not.toContain(`<w:hyperlink w:anchor="${anchor}" r:id`)
    // the heading re-emits the hidden bookmark the anchor points at
    expect(xml).toContain(`w:name="${anchor}"`)
    // round-trip: re-opened document keeps href="#anchor" and the heading bookmark
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const reopened = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reparsed.blocks) as never,
    })
    editors.add(reopened)
    expect(linkHrefsIn(reopened.state.doc)).toContain(`#${anchor}`)
    expect(
      (reopened.state.doc.firstChild?.attrs.hiddenBookmarks as string[]).includes(anchor),
    ).toBe(true)
  })

  it('saves a bookmark link by its own name without a new bookmark', async () => {
    const { parsed, editor } = await openDoc()
    insertInternalLink(editor, 'Conclusion')
    expect(linkHrefsIn(editor.state.doc)).toContain('#Conclusion')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xml = await docXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(xml).toContain('<w:hyperlink w:anchor="Conclusion">')
    expect(xml).not.toContain('<w:hyperlink w:anchor="Conclusion" r:id')
    // no duplicate bookmarkStart for the pre-existing "Conclusion" bookmark
    expect(xml.match(/w:name="Conclusion"/g)?.length).toBe(1)
  })

  it('leaves external hyperlinks and their relationships untouched', async () => {
    const { source, parsed, editor } = await openDoc()
    insertInternalLink(editor, 'Chapter 1')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const bytes = await saveDocx(parsed, plan.saveBlocks)
    const zip = await JSZip.loadAsync(bytes)
    const xml = await zip.file('word/document.xml')!.async('string')
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(xml).toContain('<w:hyperlink r:id="rId7">')
    expect(rels).toContain('Id="rId7"')
    // untouched bytes when nothing references the source doc otherwise
    expect(
      await saveDocx(
        parsed,
        pmDocToSavePlan(blocksToPmDoc(parsed.blocks) as never, parsed.blocks).saveBlocks,
      ),
    ).toEqual(source)
  })

  it('reads w:anchor hyperlinks from imported documents as #href links', async () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
      '<w:bookmarkStart w:id="2" w:name="_Toc555"/><w:bookmarkEnd w:id="2"/>' +
      '<w:r><w:t>Target</w:t></w:r></w:p>' +
      '<w:p><w:hyperlink w:anchor="_Toc555"><w:r><w:t>jump</w:t></w:r></w:hyperlink></w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.add(editor)
    expect(linkHrefsIn(editor.state.doc)).toContain('#_Toc555')
    // editing the docx without touching this paragraph keeps the anchor byte-identical
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    const xml = await docXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(xml).toContain('<w:hyperlink w:anchor="_Toc555">')
  })
})

describe('internal link jump (mod+click)', () => {
  it('the jump modifier excludes Ctrl on macOS: that click is the context menu (BUG-711)', () => {
    // mac: Cmd+click jumps; Ctrl+click (right-click) and Cmd+Ctrl must not
    expect(isInternalLinkNavClick({ metaKey: true, ctrlKey: false }, true)).toBe(true)
    expect(isInternalLinkNavClick({ metaKey: false, ctrlKey: true }, true)).toBe(false)
    expect(isInternalLinkNavClick({ metaKey: true, ctrlKey: true }, true)).toBe(false)
    expect(isInternalLinkNavClick({ metaKey: false, ctrlKey: false }, true)).toBe(false)
    // elsewhere both Ctrl+click and Cmd+click jump (unchanged behavior)
    expect(isInternalLinkNavClick({ metaKey: false, ctrlKey: true }, false)).toBe(true)
    expect(isInternalLinkNavClick({ metaKey: true, ctrlKey: false }, false)).toBe(true)
    expect(isInternalLinkNavClick({ metaKey: false, ctrlKey: false }, false)).toBe(false)
  })

  it('resolves bookmark and hidden-bookmark targets to their node positions', () => {
    const editor = makeEditor()
    const doc = editor.state.doc
    // node positions: heading1=0, heading2=heading1 size, bookmark para, plain para
    let heading2Pos = -1
    let bookmarkPos = -1
    doc.forEach((node, offset) => {
      if (node.type.name === 'docHeading' && node.textContent === 'Section A') heading2Pos = offset
      if (Array.isArray(node.attrs.bookmarks) && node.attrs.bookmarks.includes('Intro'))
        bookmarkPos = offset
    })
    // hidden bookmark lookup
    const tr = doc.nodeAt(heading2Pos)!
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(heading2Pos, undefined, {
        ...tr.attrs,
        hiddenBookmarks: ['_Toc777'],
      }),
    )
    expect(resolveInternalLinkTarget(editor.state.doc, 'Intro')).toBe(bookmarkPos)
    expect(resolveInternalLinkTarget(editor.state.doc, '_Toc777')).toBe(heading2Pos)
    expect(resolveInternalLinkTarget(editor.state.doc, 'NoSuchAnchor')).toBeNull()
  })

  it('mod+click on an internal link moves the selection to the target', () => {
    const editor = makeEditor({
      type: 'doc',
      content: [
        {
          type: 'docHeading',
          attrs: { level: 1, hiddenBookmarks: ['_Toc42'] },
          content: [{ type: 'text', text: 'Top' }],
        },
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'go ' },
            {
              type: 'text',
              text: 'there',
              marks: [{ type: 'link', attrs: { href: '#_Toc42', rId: null } }],
            },
          ],
        },
      ],
    })
    document.body.appendChild(editor.view.dom)
    const anchor = editor.view.dom.querySelector('a[href="#_Toc42"]') as HTMLAnchorElement
    expect(anchor).toBeTruthy()
    act(() => {
      anchor.dispatchEvent(new MouseEvent('click', { bubbles: true, metaKey: true }))
    })
    // caret moved into the heading (start of doc)
    expect(editor.state.selection.from).toBeLessThanOrEqual(2)
    // plain click (no mod) keeps the caret where it was
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    act(() => {
      anchor.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(editor.state.selection.from).toBe(editor.state.doc.content.size - 1)
    editor.view.dom.remove()
  })
})

describe('link dialog UI (Place in This Document pane)', () => {
  let editor: Editor
  let root: Root
  let container: HTMLElement

  beforeEach(() => {
    editor = makeEditor()
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
  })

  const openLinkDialog = () => {
    act(() =>
      root.render(
        createElement(Ribbon, {
          ...ribbonProps(editor, computeFormatState(editor)),
        }),
      ),
    )
    const insertTab = [...container.querySelectorAll<HTMLButtonElement>('.ribbon-tab')].find(
      (b) => b.textContent === t('ribbonTabInsert'),
    )!
    act(() => insertTab.click())
    const linkButton = [...container.querySelectorAll<HTMLButtonElement>('.rb-small')].find(
      (b) => b.textContent?.trim() === t('ribbonLink'),
    )!
    act(() => linkButton.click())
    return container.querySelector<HTMLElement>('.modal-backdrop')!
  }

  const tabButton = (
    backdrop: HTMLElement,
    key: 'ribbonLinkTabAddress' | 'ribbonLinkTabDocument',
  ) =>
    [...backdrop.querySelectorAll<HTMLButtonElement>('button')].find(
      (b) => b.textContent === t(key),
    )!

  it('lists headings and bookmarks and inserts an anchored heading link', () => {
    const backdrop = openLinkDialog()
    act(() => tabButton(backdrop, 'ribbonLinkTabDocument').click())
    const rows = [...backdrop.querySelectorAll('.bookmark-list .bookmark-name')] as HTMLElement[]
    const labels = rows.map((r) => r.textContent)
    expect(labels).toContain('Chapter One')
    expect(labels).toContain('Section A')
    expect(labels).toContain('Intro')
    // heading rows indent by level
    expect(rows[0].style.marginLeft).toBe('0px')
    expect(rows[1].style.marginLeft).toBe('14px')
    // pick the heading, insert with its text as display text
    act(() => rows[0].click())
    const ok = [...backdrop.querySelectorAll<HTMLButtonElement>('button')].find(
      (b) => b.textContent === t('ribbonInsert'),
    )!
    expect(ok.disabled).toBe(false)
    act(() => ok.click())
    const href = linkHrefsIn(editor.state.doc).find((h) => h.startsWith('#'))!
    expect(href).toMatch(/^#_Toc\d{9}$/)
    const inserted = editor.state.doc.textBetween(0, editor.state.doc.content.size)
    expect(inserted).toContain('Chapter OneChapter One') // heading + link text
    const heading = editor.state.doc.firstChild!
    expect((heading.attrs.hiddenBookmarks as string[]).includes(href!.slice(1))).toBe(true)
    // dialog closed
    expect(document.querySelector('.modal-backdrop')).toBeNull()
  })

  it('inserts a bookmark link under the bookmark name', () => {
    const backdrop = openLinkDialog()
    act(() => tabButton(backdrop, 'ribbonLinkTabDocument').click())
    const bookmarkRow = [...backdrop.querySelectorAll<HTMLButtonElement>('.bookmark-name')].find(
      (b) => b.textContent === 'Intro',
    )!
    act(() => bookmarkRow.click())
    const ok = [...backdrop.querySelectorAll<HTMLButtonElement>('button')].find(
      (b) => b.textContent === t('ribbonInsert'),
    )!
    act(() => ok.click())
    expect(linkHrefsIn(editor.state.doc)).toContain('#Intro')
    // no hidden bookmark got stamped for a bookmark target
    expect(editor.state.doc.firstChild?.attrs.hiddenBookmarks).toBeNull()
  })

  it('opens on the document pane with the target picked when editing an internal link', () => {
    // insert first, then reopen over the caret inside it
    const backdrop1 = openLinkDialog()
    act(() => tabButton(backdrop1, 'ribbonLinkTabDocument').click())
    const headingRow = [...backdrop1.querySelectorAll<HTMLButtonElement>('.bookmark-name')].find(
      (b) => b.textContent === 'Chapter One',
    )!
    act(() => headingRow.click())
    const ok1 = [...backdrop1.querySelectorAll<HTMLButtonElement>('button')].find(
      (b) => b.textContent === t('ribbonInsert'),
    )!
    act(() => ok1.click())
    const href = linkHrefsIn(editor.state.doc).find((h) => h.startsWith('#'))!

    // caret inside the link, reopen
    let linkPos = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.marks.some((m) => m.type.name === 'link')) linkPos = pos
    })
    editor.commands.setTextSelection(linkPos + 1)
    const backdrop2 = openLinkDialog()
    expect(tabButton(backdrop2, 'ribbonLinkTabDocument').className).toContain('btn-primary')
    expect(tabButton(backdrop2, 'ribbonLinkTabAddress').className).not.toContain('btn-primary')
    // the heading target is pre-selected (highlighted row)
    const selected = backdrop2.querySelector('.bookmark-row[style*="--hover"]')
    expect(selected?.textContent).toContain('Chapter One')
    // Apply with no change keeps the same anchor
    const ok2 = [...backdrop2.querySelectorAll<HTMLButtonElement>('button')].find(
      (b) => b.textContent === t('ribbonApply'),
    )!
    act(() => ok2.click())
    expect(linkHrefsIn(editor.state.doc)).toContain(href)
  })

  it('keeps the address pane for external links and requires a target on the document pane', () => {
    const backdrop = openLinkDialog()
    expect(tabButton(backdrop, 'ribbonLinkTabAddress').className).toContain('btn-primary')
    const ok = [...backdrop.querySelectorAll<HTMLButtonElement>('button')].find(
      (b) => b.textContent === t('ribbonInsert'),
    )!
    expect(ok.disabled).toBe(true) // empty address
    act(() => tabButton(backdrop, 'ribbonLinkTabDocument').click())
    const okDoc = [...backdrop.querySelectorAll<HTMLButtonElement>('button')].find(
      (b) => b.textContent === t('ribbonInsert'),
    )!
    expect(okDoc.disabled).toBe(true) // no target picked yet
    // switch back and insert a URL link the old way (React-controlled input:
    // drive the native value setter so the change event updates the state)
    act(() => tabButton(backdrop, 'ribbonLinkTabAddress').click())
    const input = [...backdrop.querySelectorAll<HTMLInputElement>('input')].find(
      (i) => i.placeholder === 'https://…',
    )!
    const setValue = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
    act(() => {
      setValue.call(input, 'https://example.com/')
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    expect(okDoc.disabled).toBe(false)
  })
})
