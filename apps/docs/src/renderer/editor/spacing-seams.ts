import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'

const key = new PluginKey<DecorationSet>('spacingSeams')

/** class applied to the seam widgets (styles.css keeps it zero-height) */
export const SEAM_CLASS = 'doc-spacing-seam'

/**
 * Word sums adjacent paragraph spacing: the gap between two blocks is the
 * previous one's space-after PLUS the next one's space-before (LO corpus
 * probe 2026-09-18, doc 13: p→H2 ink gap 20.42pt = 8pt after + 10pt before +
 * leading; H1→H2 19.25pt = 6+10+leading — max() would sit 6–8pt lower). CSS
 * collapses sibling margins to the max, packing intra-page content 8–10.7px
 * tighter than Word (BUG-1401). A zero-height flow-root box between the two
 * blocks cannot self-collapse, so both margins apply in full — exactly the
 * page-gap widget's effect at page boundaries, applied to every DOM-adjacent
 * pair that actually declares spacing on both sides.
 */
export const SpacingSeamsExtension = Extension.create({
  name: 'spacingSeams',
  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            const next = tr.getMeta(key) as DecorationSet | undefined
            if (next) return next
            return set.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)
          },
        },
      }),
    ]
  },
})

const isTableBlock = (el: HTMLElement): boolean =>
  el.tagName === 'TABLE' || !!el.querySelector('table')

/** pagination plumbing between blocks, not flow content (the DOM scans in App/preview skip the same set) */
const isNonContent = (el: HTMLElement): boolean =>
  el.classList.contains('page-gap') ||
  el.classList.contains('page-float-host') ||
  el.classList.contains(SEAM_CLASS)

/** margins below this are noise (sub-px struts): a seam would change nothing visible */
const MIN_SEAM_PX = 0.5

const marginTopPx = (el: HTMLElement): number => parseFloat(getComputedStyle(el).marginTop) || 0
const marginBottomPx = (el: HTMLElement): number =>
  parseFloat(getComputedStyle(el).marginBottom) || 0

/**
 * Boundaries (doc positions, inserted before the NEXT block) that need an
 * anti-collapse seam: both neighbors declare vertical spacing and neither is a
 * table (a table's 2px canvas margin is TABLE_SEAM_PX bookkeeping, not Word
 * paragraph spacing). Pairs where one side is zero keep CSS max() == sum() and
 * stay undecorated, which also keeps the `A + B` sibling selectors of the
 * list/contextual-spacing rules (styles.css, doc-style-css.ts) matching.
 */
export function planSpacingSeams(
  pm: HTMLElement,
  posOf: (el: HTMLElement) => number | null,
): number[] {
  const kids: HTMLElement[] = []
  for (const el of Array.from(pm.children) as HTMLElement[]) {
    if (isNonContent(el)) continue
    kids.push(el)
  }
  const out: number[] = []
  for (let i = 0; i + 1 < kids.length; i++) {
    const prev = kids[i] as HTMLElement
    const next = kids[i + 1] as HTMLElement
    if (isTableBlock(prev) || isTableBlock(next)) continue
    if (marginBottomPx(prev) <= MIN_SEAM_PX) continue
    if (marginTopPx(next) <= MIN_SEAM_PX) continue
    const pos = posOf(next)
    if (pos != null && pos > 0) out.push(pos)
  }
  return out
}

function sameSet(a: DecorationSet, b: DecorationSet): boolean {
  const af = a.find()
  const bf = b.find()
  return (
    af.length === bf.length &&
    af.every(
      (d, i) =>
        d.from === bf[i].from &&
        (d.spec as { key?: string }).key === (bf[i].spec as { key?: string }).key,
    )
  )
}

/**
 * Rebuild the seam decorations from the live DOM (an empty list clears them).
 * Reads computed margins, so the suppression rules that already ran
 * (.page-break-lead, contextual/auto list spacing) are respected: a zeroed
 * side means no seam, and the sibling-combinator selectors keep matching.
 */
export function syncSpacingSeams(view: EditorView): void {
  const pm = view.dom as HTMLElement
  const positions = planSpacingSeams(pm, (el) => {
    try {
      return view.state.doc.resolve(view.posAtDOM(el, 0)).before(1)
    } catch {
      return null
    }
  })
  const decos = positions.map((pos, i) =>
    Decoration.widget(
      pos,
      () => {
        const seam = document.createElement('div')
        seam.className = SEAM_CLASS
        seam.contentEditable = 'false'
        return seam
      },
      // ordinal key (not pos): edits above shift positions without changing
      // the pair, so PM reuses the widget DOM; side -1 with the extension
      // registered before PaginationGaps keeps seams ahead of page-gap widgets
      { side: -1, key: `spacing-seam-${i}` },
    ),
  )
  const next = DecorationSet.create(view.state.doc, decos)
  const prev = key.getState(view.state)
  if (prev && sameSet(prev, next)) return
  view.dispatch(view.state.tr.setMeta(key, next).setMeta('addToHistory', false))
}
