import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { Slice } from '@tiptap/pm/model'
import { buildExtensions } from '../src/renderer/editor/extensions'

// BUG-1740: every pasted HTML table degraded to an uneditable rawHtml chip
// holding bare `<tbody>` (auto-inserted by DOMParser) and Ctrl+S wrote that
// `<tbody>` without its `<table>` wrapper — invalid HTML for any other markdown
// renderer. Root cause: the table node's content model (`tableRow+`) cannot
// hold the thead/tbody/tfoot wrapper elements, so the RawHtml catch-all claimed
// them. The schema now skips those structural wrappers (extensions.ts), and a
// pasted or file-sourced table assembles into a native editable table that
// serializes as valid GFM.

// Undestroyed views leave DOMObserver flush timers that fire after jsdom
// teardown ("document is not defined" unhandled error) — destroy everything.
const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

function createEditor(content = ''): Editor {
  return new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content,
  })
}

/** synthetic clipboard mirroring the DataTransfer surface the handler reads */
function clipboard(flavors: Record<string, string>): DataTransfer {
  return { files: [], getData: (type: string) => flavors[type] ?? '' } as unknown as DataTransfer
}

/** run the editor's handlePaste props exactly like ProseMirror: first truthy wins */
function paste(editor: Editor, data: DataTransfer): boolean {
  const event = { clipboardData: data } as unknown as ClipboardEvent
  let handled = false
  editor.view.someProp('handlePaste', (fn) => {
    if (handled) return
    handled = fn(editor.view, event, Slice.empty) === true
  })
  return handled
}

interface TableJSON {
  type: string
  content?: Array<{ type: string; content?: unknown[] }>
}

/** the first table node of the doc, or undefined */
function firstTable(editor: Editor): TableJSON | undefined {
  return editor.state.doc.toJSON().content?.find((n: TableJSON) => n.type === 'table') as
    TableJSON | undefined
}

function rowsOf(table: TableJSON): Array<{ type: string; content?: unknown[] }> {
  return (table.content ?? []) as Array<{ type: string; content?: unknown[] }>
}

describe('rich paste of HTML tables lands as a native table (BUG-1740)', () => {
  it('a tbody-wrapped table parses into table rows, not a rawHtml chip', () => {
    const editor = createEditor()
    const html =
      '<table><tbody><tr><td>A1</td><td>B1</td></tr><tr><td>A2</td><td>B2</td></tr></tbody></table>'
    expect(paste(editor, clipboard({ 'text/html': html, 'text/plain': 'A1 B1' }))).toBe(true)

    const table = firstTable(editor)
    expect(table).toBeDefined()
    expect(rowsOf(table!)).toHaveLength(2)
    expect(JSON.stringify(editor.state.doc.toJSON())).not.toContain('"rawHtml"')

    const md = editor.getMarkdown()
    // valid GFM: the saved file must not carry a bare <tbody> anymore
    expect(md).toContain('| A1')
    expect(md).toMatch(/\|[- :|]+\|/)
    expect(md).not.toContain('<tbody')
    expect(md).not.toContain('<table')
  })

  it('a Word-flavored table (class/style/width junk) assembles the same way', () => {
    const editor = createEditor()
    const html =
      '<table class="MsoTableGrid" style="width:100%"><tbody>' +
      '<tr style="mso-yfti-firstrow"><td width="151" style="padding:0cm"><p class=MsoNormal><b>A1</b></p></td></tr>' +
      '<tr><td>B1</td></tr></tbody></table>'
    paste(editor, clipboard({ 'text/html': html, 'text/plain': 'A1' }))

    const table = firstTable(editor)
    expect(table).toBeDefined()
    expect(rowsOf(table!)).toHaveLength(2)
    const md = editor.getMarkdown()
    expect(md).toContain('**A1**')
    expect(md).not.toContain('mso')
    expect(md).not.toContain('<tbody')
  })

  it('thead/th keeps a real header row with GFM alignment row after it', () => {
    const editor = createEditor()
    const html =
      '<table><thead><tr><th>H1</th><th>H2</th></tr></thead>' +
      '<tbody><tr><td>a</td><td>b</td></tr></tbody></table>'
    paste(editor, clipboard({ 'text/html': html, 'text/plain': 'H1' }))

    const table = firstTable(editor)
    expect(table).toBeDefined()
    const rows = rowsOf(table!)
    expect(rows).toHaveLength(2)
    expect(JSON.stringify(rows[0])).toContain('"tableHeader"')

    const md = editor.getMarkdown()
    expect(md).toContain('| H1')
    expect(md).toMatch(/\| --- /)
  })

  it('colgroup width metadata is dropped without breaking the table apart', () => {
    const editor = createEditor()
    const html =
      '<table><colgroup><col span="2" /></colgroup><tbody><tr><td>A</td><td>B</td></tr></tbody></table>'
    paste(editor, clipboard({ 'text/html': html, 'text/plain': 'A B' }))

    const table = firstTable(editor)
    expect(table).toBeDefined()
    expect(rowsOf(table!)).toHaveLength(1)
    expect(JSON.stringify(editor.state.doc.toJSON())).not.toContain('"rawHtml"')
  })

  it('a caption survives as a chip beside an intact table (no GFM caption exists)', () => {
    const editor = createEditor()
    const html = '<table><caption>CAP</caption><tbody><tr><td>A1</td></tr></tbody></table>'
    paste(editor, clipboard({ 'text/html': html, 'text/plain': 'CAP A1' }))

    const table = firstTable(editor)
    expect(table).toBeDefined()
    const md = editor.getMarkdown()
    expect(md).toContain('<caption>CAP</caption>')
    expect(md).toContain('| A1')
    expect(md).not.toContain('<tbody')
  })

  it('save/reopen is a fixed point: the table survives a second parse/serialize', () => {
    const editor = createEditor()
    const html = '<table><tbody><tr><td>A1</td></tr></tbody></table>'
    paste(editor, clipboard({ 'text/html': html, 'text/plain': 'A1' }))

    const manager = editor.markdown!
    const first = editor.getMarkdown()
    expect(first).not.toContain('<tbody')
    const reopen = manager.serialize(manager.parse(first))
    expect(reopen).toBe(first)
  })
})

describe('markdown files with HTML tables open as native tables (BUG-1740)', () => {
  it('a file holding a full <table><tbody> block parses to a table and saves as GFM', () => {
    const editor = createEditor()
    const md = 'before\n\n<table><tbody><tr><td>A1</td><td>B1</td></tr></tbody></table>\n\nafter\n'
    const json = editor.markdown!.parse(md)
    const table = json.content?.find((n) => n.type === 'table')
    expect(table).toBeDefined()
    const out = editor.markdown!.serialize(json)
    expect(out).toContain('before')
    expect(out).toContain('| A1')
    expect(out).not.toContain('<tbody')
  })

  it('a legacy file with a bare <tbody> block (pre-fix save) parses without crash or malformed chip', () => {
    const editor = createEditor()
    const md = '<tbody><tr><td>A1</td></tr></tbody>\n'
    const json = editor.markdown!.parse(md)
    expect(JSON.stringify(json)).not.toContain('"rawHtml"')
    // the HTML fragment parser foster-parents the orphan tbody tags away; the
    // cell text stays readable plain text instead of an uneditable chip
    expect(JSON.stringify(json)).toContain('A1')
  })
})
