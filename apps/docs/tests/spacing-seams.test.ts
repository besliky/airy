import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  SEAM_CLASS,
  SpacingSeamsExtension,
  planSpacingSeams,
  syncSpacingSeams,
} from '../src/renderer/editor/spacing-seams'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createTrackedEditor, drainTrackedEditors } from './helpers/tracked-editor'
import { measureBlocks } from '../src/renderer/pagination-measure'

afterEach(() => {
  drainTrackedEditors()
  vi.restoreAllMocks()
})

const el = (tag: string, style: string): HTMLElement => {
  const node = document.createElement(tag)
  node.setAttribute('style', style)
  return node
}

describe('seam planning (Word sums adjacent spacing; CSS collapses to the max)', () => {
  const plan = (kids: HTMLElement[]) => {
    const pm = document.createElement('div')
    for (const k of kids) pm.appendChild(k)
    return planSpacingSeams(pm, (node) => kids.indexOf(node as HTMLElement) * 10 + 5)
  }

  it('both sides declare spacing: seam between them (p 8pt after + H2 10pt before)', () => {
    const p = el('p', 'margin-bottom:10.67px')
    const h2 = el('h2', 'margin-top:13.33px')
    expect(plan([p, h2])).toEqual([15])
  })

  it('one side zero: no seam — CSS max() already equals the Word sum (Normal p→p)', () => {
    expect(plan([el('p', 'margin-bottom:10.67px'), el('p', 'margin-top:0px')])).toEqual([])
    expect(plan([el('p', 'margin-bottom:0px'), el('h2', 'margin-top:13.33px')])).toEqual([])
  })

  it('a suppressed page leader (.page-break-lead margin-top:0) keeps no seam', () => {
    const lead = el('h2', 'margin-top:0px')
    lead.className = 'page-break-lead'
    expect(plan([el('p', 'margin-bottom:10.67px'), lead])).toEqual([])
  })

  it('tables are exempt: their 2px canvas margin is TABLE_SEAM_PX bookkeeping', () => {
    const tbl = el('div', 'margin-top:2px')
    tbl.appendChild(document.createElement('table'))
    expect(plan([el('p', 'margin-bottom:10.67px'), tbl])).toEqual([])
    const after = el('p', 'margin-top:13.33px')
    expect(plan([tbl, after])).toEqual([])
  })

  it('pagination widgets and existing seams are not flow content: pairs skip over them', () => {
    const gap = el('div', '')
    gap.className = 'page-gap'
    const host = el('div', '')
    host.className = 'page-float-host'
    const oldSeam = el('div', '')
    oldSeam.className = SEAM_CLASS
    const p = el('p', 'margin-bottom:10.67px')
    const h1 = el('h1', 'margin-top:16px')
    expect(plan([p, gap, host, oldSeam, h1])).toEqual([45])
  })

  it('sub-pixel margins are noise: no seam', () => {
    expect(plan([el('p', 'margin-bottom:0.3px'), el('h2', 'margin-top:13.33px')])).toEqual([])
  })
})

describe('syncSpacingSeams (live editor)', () => {
  // jsdom resolves an element's computed style once and never re-resolves it
  // when the inline style changes; serve margins straight from the attribute
  const honorInlineMargins = () => {
    const real = window.getComputedStyle.bind(window)
    return vi
      .spyOn(window, 'getComputedStyle')
      .mockImplementation((elt: Element, pseudo?: string | null) => {
        const cs = real(elt, pseudo)
        const inline = (elt as HTMLElement).getAttribute('style') ?? ''
        const mt = /margin-top:\s*([\d.]+)px/.exec(inline)?.[1]
        const mb = /margin-bottom:\s*([\d.]+)px/.exec(inline)?.[1]
        return new Proxy(cs, {
          get(target, prop) {
            if (prop === 'marginTop' && mt !== undefined) return `${mt}px`
            if (prop === 'marginBottom' && mb !== undefined) return `${mb}px`
            const value = target[prop as keyof typeof target]
            return typeof value === 'function' ? value.bind(target) : value
          },
        }) as CSSStyleDeclaration
      })
  }

  const mount = () =>
    createTrackedEditor({
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'body one' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'body two' }] },
          {
            type: 'docHeading',
            attrs: { level: 2 },
            content: [{ type: 'text', text: 'section' }],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'body three' }] },
        ],
      } as never,
    })

  const setSpacing = (view: { dom: HTMLElement }, mb: string, mt: string) => {
    const blocks = Array.from(view.dom.children).filter(
      (c) => !(c as HTMLElement).classList.contains(SEAM_CLASS),
    )
    const body = blocks[1] as HTMLElement
    const heading = blocks[2] as HTMLElement
    // jsdom's getComputedStyle only re-resolves after a whole-attribute swap
    body.setAttribute('style', `margin-bottom:${mb}`)
    heading.setAttribute('style', `margin-top:${mt}`)
  }

  it('inserts zero-visual flow-root widgets between the declaring pair only', () => {
    honorInlineMargins()
    const editor = mount()
    setSpacing(editor.view, '10.67px', '13.33px')
    syncSpacingSeams(editor.view)
    const seams = editor.view.dom.querySelectorAll(':scope > .' + SEAM_CLASS)
    expect(seams.length).toBe(1)
    // between body two and the heading, ahead of any page-gap widget at the
    // same boundary (extension order, see the registration guard below)
    const heading = Array.from(editor.view.dom.children).find((c) => c.tagName === 'H2')
    expect(heading?.previousElementSibling?.classList.contains(SEAM_CLASS)).toBe(true)
    expect((seams[0] as HTMLElement).contentEditable).toBe('false')
  })

  it('is idempotent: an unchanged plan dispatches nothing', () => {
    honorInlineMargins()
    const editor = mount()
    setSpacing(editor.view, '10.67px', '13.33px')
    const transactions = vi.fn()
    editor.on('transaction', transactions)
    syncSpacingSeams(editor.view)
    expect(transactions).toHaveBeenCalledTimes(1)
    syncSpacingSeams(editor.view)
    expect(transactions).toHaveBeenCalledTimes(1)
    expect(editor.view.dom.querySelectorAll(':scope > .' + SEAM_CLASS).length).toBe(1)
  })

  it('drops the seam when the pair stops declaring both sides', () => {
    honorInlineMargins()
    const editor = mount()
    setSpacing(editor.view, '10.67px', '13.33px')
    syncSpacingSeams(editor.view)
    expect(editor.view.dom.querySelectorAll(':scope > .' + SEAM_CLASS).length).toBe(1)
    setSpacing(editor.view, '10.67px', '0px')
    syncSpacingSeams(editor.view)
    expect(editor.view.dom.querySelectorAll(':scope > .' + SEAM_CLASS).length).toBe(0)
  })
})

describe('registration order (widget order at a shared boundary)', () => {
  it('SpacingSeamsExtension precedes PaginationGapsExtension so seams render ahead of page gaps', () => {
    const names = editorExtensions.map((e) => e.name)
    expect(names.indexOf('spacingSeams')).toBeGreaterThanOrEqual(0)
    expect(names.indexOf('spacingSeams')).toBeLessThan(names.indexOf('paginationGaps'))
    expect(SpacingSeamsExtension.name).toBe('spacingSeams')
  })
})

describe('measurement folds the summed gap into the engine (BUG-1401 geometry)', () => {
  const rect = (top: number, height: number) =>
    ({ top, bottom: top + height, height, width: 400, left: 0, right: 400 }) as DOMRect

  it("with the seam, the DOM gap is the Word sum and lands in the previous block's spaceAfter", () => {
    // p (20px tall, 8pt after = 10.67px) → seam → H2 (10pt before = 13.33px):
    // the flow-root seam keeps both margins, so the H2 border box sits 24px
    // below the paragraph's border box (13.33px without it — the collapsed
    // layout BUG-1400 measured in-product)
    const pm = document.createElement('div')
    const p = el('p', '')
    const seam = el('div', '')
    seam.className = SEAM_CLASS
    const h2 = el('h2', '')
    for (const node of [p, seam, h2]) {
      pm.appendChild(node)
      node.getBoundingClientRect = () => rect(0, 0)
    }
    p.getBoundingClientRect = () => rect(0, 20)
    seam.getBoundingClientRect = () => rect(44, 0)
    h2.getBoundingClientRect = () => rect(44, 25)
    const { blocks } = measureBlocks(pm, 0, 1)
    // the seam itself never becomes a block (zero-height, no breaks)
    expect(blocks.length).toBe(2)
    expect(blocks[0].spaceAfterPx).toBeCloseTo(24, 1)
    expect(blocks[0].height).toBeCloseTo(44, 1)
    expect(blocks[1].top).toBeCloseTo(44, 1)
  })

  it('collapsed fixture: the pre-fix max() gap is what the engine used to see', () => {
    const pm = document.createElement('div')
    const p = el('p', '')
    const h2 = el('h2', '')
    for (const node of [p, h2]) {
      pm.appendChild(node)
      node.getBoundingClientRect = () => rect(0, 0)
    }
    p.getBoundingClientRect = () => rect(0, 20)
    h2.getBoundingClientRect = () => rect(33.33, 25)
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks[0].spaceAfterPx).toBeCloseTo(13.33, 1)
  })
})
