/**
 * UX-1704: the word-wrap toggle. The requirement is that toggling flips the
 * wrap extension in place — the EditorView is never recreated, so undo
 * history, scroll position and the find panel all survive the switch. The
 * extension inside the compartment is read back through the same probe
 * CodeMirror itself uses (`contentAttributes` carrying the cm-lineWrapping
 * class), which is deterministic under jsdom where computed styles are not.
 */
import { describe, expect, it } from 'vitest'
import { EditorState } from '@codemirror/state'
import { EditorView } from '@codemirror/view'
import { undo } from '@codemirror/commands'

import {
  buildExtensions,
  setLineWrap,
  wrapActive,
  wrapCompartment,
} from '../src/renderer/source/cm-setup'

function setup(text: string, wrap = true) {
  const host = document.createElement('div')
  document.body.appendChild(host)
  const view = new EditorView({
    parent: host,
    state: EditorState.create({
      doc: text,
      extensions: buildExtensions(() => {}, { wrap }),
    }),
  })
  return { host, view }
}

describe('word wrap default', () => {
  it('stays on when nothing is passed (the pre-toggle behavior)', () => {
    const { view } = setup('<p>hello</p>')
    expect(wrapActive(view.state)).toBe(true)
  })

  it('can be created unwrapped from a persisted off preference', () => {
    const { view } = setup('<p>hello</p>', false)
    expect(wrapActive(view.state)).toBe(false)
  })
})

describe('setLineWrap (compartment reconfigure)', () => {
  it('flips the wrap extension without recreating the editor', () => {
    const { host, view } = setup('<p>hello</p>')
    const domBefore = view.dom
    const docBefore = view.state.doc.toString()

    setLineWrap(view, false)
    expect(wrapActive(view.state)).toBe(false)
    // same view instance, same DOM subtree: nothing was rebuilt
    expect(view.dom).toBe(domBefore)
    expect(host.contains(view.dom)).toBe(true)
    expect(view.state.doc.toString()).toBe(docBefore)

    setLineWrap(view, true)
    expect(wrapActive(view.state)).toBe(true)
    expect(view.dom).toBe(domBefore)
  })

  it('keeps the undo history and the document across the toggle', () => {
    const { view } = setup('<p>a</p>')
    view.dispatch({ changes: { from: 3, to: 4, insert: 'b' } })
    expect(view.state.doc.toString()).toBe('<p>b</p>')

    setLineWrap(view, false)
    expect(view.state.doc.toString()).toBe('<p>b</p>')
    // history survived the reconfigure: the pre-toggle edit is still undoable
    expect(undo(view)).toBe(true)
    expect(view.state.doc.toString()).toBe('<p>a</p>')
  })

  it('the live editor was built with the exported compartment', () => {
    // if buildExtensions ever used a different compartment, an effect from
    // this one would not address the built state — pin the identity by
    // dispatching through the exported compartment and observing the flip
    const { view } = setup('<p>x</p>')
    view.dispatch({ effects: wrapCompartment.reconfigure(EditorView.lineWrapping) })
    expect(wrapActive(view.state)).toBe(true)
    setLineWrap(view, false)
    expect(wrapActive(view.state)).toBe(false)
  })
})
