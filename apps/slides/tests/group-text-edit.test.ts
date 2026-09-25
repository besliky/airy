/**
 * BUG-1725 / SL-EDIT-1 regression: double-clicking a group child's text must
 * open the text-edit overlay for THAT child — not re-select the whole group,
 * and never bind the overlay to a sibling shape behind it (one audit run
 * committed the typed text into the group's background auto shape).
 *
 * react-konva cannot mount under jsdom, so this follows the repo's split test
 * style: the pure resolution helper (group-hit.ts) is unit-tested, the
 * SlideCanvas/App wiring is asserted source-contract style, and the typed-text
 * commit is exercised against the real engine (same group-addressed setText op
 * the renderer's commitEdit sends through slides:edit-text).
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { beforeEach, describe, expect, it } from 'vitest'
import {
  addElement,
  createBlankPptx,
  openPptx,
  patchSlideXml,
  savePptx,
  type OpenedPptx,
  type TextElement,
} from '@airy-office/pptx-engine'
import type { RenderNode, ShapeRenderNode } from '@airy-office/pptx-render'
import { runTxn } from '../src/main/ops'
import { groupAllowsChildTextEdit, resolveGroupDblClick } from '../src/renderer/group-hit'

const here = dirname(fileURLToPath(import.meta.url))
const canvasSource = readFileSync(join(here, '../src/renderer/SlideCanvas.tsx'), 'utf8')
const appSource = readFileSync(join(here, '../src/renderer/App.tsx'), 'utf8')

// ── Fixtures: a group like the audit's grpedit.pptx — background rect + text children A/B ──

const bg = {
  type: 'shape',
  sourceId: 'grp-bg',
  box: { x: 0, y: 0, w: 200, h: 120 },
} as unknown as RenderNode
const textA = {
  type: 'text',
  sourceId: 'grp-child-a',
  box: { x: 10, y: 8, w: 120, h: 24 },
} as unknown as RenderNode
const textB = {
  type: 'text',
  sourceId: 'grp-child-b',
  box: { x: 10, y: 40, w: 120, h: 24 },
} as unknown as RenderNode
const group = { children: [bg, textA, textB] }

describe('groupAllowsChildTextEdit (overlay alignment gate)', () => {
  it('allows text editing on a plain (unrotated, unflipped) group', () => {
    expect(groupAllowsChildTextEdit({})).toBe(true)
    expect(groupAllowsChildTextEdit({ rotationDeg: 0 })).toBe(true)
  })

  it('blocks text editing when the overlay could not follow the group transform', () => {
    expect(groupAllowsChildTextEdit({ rotationDeg: 45 })).toBe(false)
    expect(groupAllowsChildTextEdit({ flipH: true })).toBe(false)
    expect(groupAllowsChildTextEdit({ flipV: true })).toBe(false)
  })
})

describe('resolveGroupDblClick picks the topmost child under the cursor', () => {
  it('a point over a text child resolves to that child, never to the background rect behind it', () => {
    // Point (100, 50) is inside textB's box AND the background rect: topmost wins,
    // so the overlay binds to the child (the audit's wrong-binding finding).
    const hit = resolveGroupDblClick(group, true, { x: 100, y: 50 })
    expect(hit.childId).toBe('grp-child-b')
    expect(hit.editText).toBe(true)
  })

  it('resolves each overlapping child by render order (later children win)', () => {
    expect(resolveGroupDblClick(group, true, { x: 20, y: 20 }).childId).toBe('grp-child-a')
    expect(resolveGroupDblClick(group, true, { x: 20, y: 60 }).childId).toBe('grp-child-b')
  })

  it('a point on the background rect alone targets the background rect', () => {
    const hit = resolveGroupDblClick(group, true, { x: 180, y: 100 })
    expect(hit.childId).toBe('grp-bg')
    expect(hit.editText).toBe(true) // it is still an editable auto shape
  })

  it('a point in the group padding outside every child selects nothing', () => {
    const hit = resolveGroupDblClick(group, true, { x: -5, y: -5 })
    expect(hit.childId).toBeNull()
    expect(hit.editText).toBe(false)
  })

  it('a rotated/flipped group enters+selects but never opens text editing (overlay cannot align)', () => {
    const hit = resolveGroupDblClick(group, false, { x: 100, y: 50 })
    expect(hit.childId).toBe('grp-child-b')
    expect(hit.editText).toBe(false)
  })

  it('a missing pointer position (dbltap corner cases) degrades to plain group entry', () => {
    expect(resolveGroupDblClick(group, true, null)).toEqual({ childId: null, editText: false })
  })

  it('a non-text child (connector line) is selected but does not open text editing', () => {
    const line = {
      type: 'shape',
      sourceId: 'grp-line',
      line: { x1: 0, y1: 0, x2: 40, y2: 0 },
      box: { x: 0, y: 100, w: 40, h: 2 },
    } as unknown as RenderNode
    const hit = resolveGroupDblClick({ children: [line] }, true, { x: 10, y: 101 })
    expect(hit.childId).toBe('grp-line')
    expect(hit.editText).toBe(false)
  })
})

describe('dblclick → text-edit wiring (source contracts)', () => {
  it('the group dblclick handler resolves the hit through group-hit and forwards editText', () => {
    expect(canvasSource).toContain(
      'const hit = resolveGroupDblClick(g, groupAllowsChildTextEdit(box), local)',
    )
    expect(canvasSource).toMatch(
      /onEnterGroup\(\s*node\.sourceId,\s*hit\.childId,\s*hit\.editText\s*\?\s*\{ editText: true/,
    )
  })

  it('the entered-mode child-edit gate shares the same alignment helper', () => {
    expect(canvasSource).toContain('const plain = groupAllowsChildTextEdit(box)')
  })

  it('the second click of a double-click pair cannot re-select the group (race suppression)', () => {
    // The suppression must sit between the marquee guard and the onSelect call.
    expect(canvasSource).toMatch(
      /if \(suppressClickRef\?\.current\) \{[\s\S]*?return\s*\}\s*\/\/ The second click of a double-click pair[\s\S]*?if \(e\.evt\.detail >= 2 && !\(e\.evt\.shiftKey \|\| e\.evt\.metaKey\)\) return\s*onSelect\(/,
    )
  })

  it('App enters the group AND binds the overlay to the hit child via the explicit groupId', () => {
    expect(appSource).toMatch(
      /if \(childId && opts\?\.editText\)\s*setEditing\(\{\s*sourceId: childId,\s*groupId,/,
    )
  })

  it('the edit commit keeps addressing the child through the group (typed text lands in the child)', () => {
    // commitEdit must include the groupId in the slides:edit-text payload so the
    // main process resolves the child bytes inside the group blob.
    expect(appSource).toContain('...(editing.groupId ? { groupId: editing.groupId } : {})')
  })
})

// ── Model + XML: the group-addressed text commit writes into the child shape ──

describe('typed text commits into the group child (model + XML)', () => {
  let opened: OpenedPptx
  let childId: string
  let groupId: string

  /** The group's children as the render tree addresses them (post-grouping child ids). */
  const grpChildren = () =>
    (
      opened.deck.slides[0]!.elements.find((e) => e.type === 'group') as unknown as {
        children: TextElement[]
      }
    ).children

  /** Joined run text of a child (empty string for an empty text body). */
  const textOf = (c: TextElement) =>
    c.text?.paragraphs.flatMap((p) => p.runs.map((r) => r.text)).join('')

  beforeEach(async () => {
    opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    // Same composition as the audit's grpedit.pptx: rect background + a text child on top.
    const bgId = addElement(slide, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 1828800, cy: 1097280 },
      fillColor: '#E7E6E6',
    }).id
    const boxId = addElement(slide, {
      kind: 'textbox',
      offset: { x: 91440, y: 91440, cx: 1371600, cy: 457200 },
      paragraphs: [{ runs: [{ text: 'GRP-CHILD-A' }] }],
    }).id
    const grouped = runTxn(opened, {
      ops: [{ op: 'groupElements', target: { slide: 0 }, els: [bgId, boxId] }],
    })
    expect(grouped.applied).toBe(true)
    const grp = opened.deck.slides[0]!.elements.find((e) => e.type === 'group')
    expect(grp).toBeDefined()
    groupId = grp!.id
    // Grouping mints fresh child ids — resolve the textbox child the way the
    // renderer does (from the group's children, by its text).
    childId = grpChildren().find((c) => textOf(c) === 'GRP-CHILD-A')!.id
    const bgChild = grpChildren().find((c) => c.id !== childId)
    expect(textOf(bgChild!)).toBe('') // the rect background is an empty text body
  })

  /** The renderer's commitEdit payload: paragraphs + group addressing. */
  const commitEdit = (text: string) =>
    runTxn(opened, {
      ops: [
        {
          op: 'setText',
          target: { slide: 0, el: childId },
          paragraphs: [{ runs: [{ text: `GRP-CHILD-A${text}` }] }],
          group: groupId,
        },
      ],
    })

  it('updates the child model without touching the background rect', () => {
    expect(commitEdit(' EDITED').applied).toBe(true)
    const child = grpChildren().find((c) => c.id === childId) as TextElement | undefined
    const bgChild = grpChildren().find((c) => c.id !== childId) as TextElement | undefined
    expect(textOf(child!)).toBe('GRP-CHILD-A EDITED')
    expect(textOf(bgChild!)).toBe('')
  })

  it('the typed text reaches the child <p:sp> inside the group blob (not the background sp)', () => {
    expect(commitEdit(' EDITED').applied).toBe(true)
    const xml = patchSlideXml(opened.deck.slides[0]!)
    const sps = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? []
    const edited = sps.find((sp) => sp.includes('GRP-CHILD-A EDITED'))
    expect(edited).toBeDefined()
    expect(edited).toContain('txBox="1"') // the textbox child, not the background auto shape
    // Exactly one sp carries the typed text; the background sp keeps its empty body.
    expect(xml.split('GRP-CHILD-A EDITED').length - 1).toBe(1)
    for (const sp of sps) {
      if (sp === edited) continue
      expect(sp).not.toContain('EDITED')
    }
  })

  it('the typed text survives save → reopen as the child text', async () => {
    expect(commitEdit(' EDITED').applied).toBe(true)
    const reopened = await openPptx(await savePptx(opened))
    const grp = reopened.deck.slides[0]!.elements.find((e) => e.type === 'group') as unknown as {
      children: TextElement[]
    }
    const texts = grp.children.map((c) =>
      c.text?.paragraphs.flatMap((p) => p.runs.map((r) => r.text)).join(''),
    )
    expect(texts).toContain('GRP-CHILD-A EDITED')
    // The background rect the audit saw absorb the text is still empty.
    expect(texts.some((t) => t === '')).toBe(true)
  })

  it('a textless auto-shape render node is still an editable-text dblclick target', () => {
    // The background rect child as the render tree emits it (shape, empty body):
    // it must resolve as editable-text-capable, matching top-level dblclick behavior.
    const node = { type: 'shape', sourceId: 'grp-bg', box: { x: 0, y: 0, w: 10, h: 10 } }
    const hit = resolveGroupDblClick({ children: [node as unknown as ShapeRenderNode] }, true, {
      x: 5,
      y: 5,
    })
    expect(hit.editText).toBe(true)
  })
})
