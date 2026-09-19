// Unit tests for the headless slides session (PAR-001): open/read formats,
// insert_content (text box + existing-shape text replace), save round-trips,
// byte preservation, mtime fencing, save-target ownership and confinement.
import { mkdtemp, readFile, rm, writeFile, rename, stat, utimes } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  addElement,
  duplicateSlide,
  openPptx,
  type OpenedPptx,
  type TextElement,
} from '@airy-office/pptx-engine'

import { FencingError } from '../src/docx/session.js'
import { SlidesSession } from '../src/slides/session.js'
import { buildFixturePptx } from './helpers/pptx-fixture.js'

let root: string
let deckPath: string
let fixture: Uint8Array

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'airy-slides-'))
  fixture = await buildFixturePptx()
  deckPath = join(root, 'deck.pptx')
  await writeFile(deckPath, fixture)
})

afterEach(async () => {
  await rm(root, { recursive: true, force: true })
})

async function openSession(): Promise<SlidesSession> {
  return SlidesSession.open(deckPath, root)
}

/** reopen saved bytes with the engine and return the parsed deck */
async function reopen(path: string): Promise<OpenedPptx> {
  return openPptx(new Uint8Array(await readFile(path)))
}

function slideTextOf(opened: OpenedPptx, slide = 0): string {
  const els = opened.deck.slides[slide]!.elements
  return els
    .map((el) =>
      el.type === 'text' || el.type === 'shape'
        ? ((el as TextElement).text?.paragraphs ?? [])
            .map((p) => p.runs.map((r) => r.text).join(''))
            .join('\n')
        : '',
    )
    .filter(Boolean)
    .join('\n')
}

describe('slides session open/read', () => {
  it('opens a fixture and reports deck/stats meta', async () => {
    const session = await openSession()
    const meta = session.meta()
    expect(meta.kind).toBe('slides')
    expect(meta.path).toBe(deckPath)
    expect(meta.dirty).toBe(false)
    expect(meta.slideCount).toBe(1)
    expect(meta.elementCount).toBe(2)
    expect(meta.wordCount).toBeGreaterThan(3)
    // 12192000 x 6858000 EMU = 13.3 x 7.5 in
    expect(meta.size).toMatchObject({ widthIn: '13.3', heightIn: '7.5' })
  })

  it('renders the index|elements|preview deck overview', async () => {
    const session = await openSession()
    const text = session.readDeck()
    expect(text).toContain('The deck has 1 slide(s)')
    expect(text).toMatch(/^0\|2\|Quarterly Review$/m)
    expect(text).toContain('Deck stats: 2 elements')
  })

  it('reads one slide in detail: element list plus full text', async () => {
    const session = await openSession()
    const text = session.readSlide(0)
    expect(text).toContain('Slide 0 of 1 — 2 element(s)')
    expect(text).toMatch(/^0\|text\|TextBox 2\|Quarterly Review$/m)
    expect(text).toMatch(/^1\|shape\|Shape 3\|Revenue grew by 12 percent$/m)
    expect(text).toContain('Full slide text (EOLs normalized to LF):')
    expect(text).toContain('Quarterly Review\nRevenue grew by 12 percent')
  })

  it('rejects out-of-range slide reads', async () => {
    const session = await openSession()
    expect(() => session.readSlide(1)).toThrow(/slide index 1 is out of range/)
    expect(() => session.readSlide(-1)).toThrow(/out of range/)
  })

  it('tightens previews of a huge deck overview to stay inside the budget', async () => {
    // 500 duplicated slides with a long first line put the wide-preview
    // overview over the 30k budget; the tightened render (20-char previews)
    // brings it back under (the docx session's middle-elide fallback is the
    // same code shape and is covered there — building a >1100-slide fixture
    // costs 30+ s of duplicateSlide time for the same branch)
    const opened = await openPptx(fixture)
    const first = opened.deck.slides[0]!.elements[0] as TextElement
    first.text!.paragraphs = [{ runs: [{ text: 'L'.repeat(80) }] }]
    first.dirty = true
    for (let i = 0; i < 500; i++) duplicateSlide(opened, 0)
    const { savePptx } = await import('@airy-office/pptx-engine')
    const big = join(root, 'big.pptx')
    await writeFile(big, await savePptx(opened))
    const session = await SlidesSession.open(big, root)
    const text = session.readDeck()
    expect(text.length).toBeLessThanOrEqual(31_000)
    expect(text).toContain('The deck has 501 slide(s)')
    // tightened previews end with the ellipsis marker
    expect(text).toMatch(/\|2\|L{19}…$/m)
  })

  it('clips a slide detail read at the 30k budget', async () => {
    const session = await openSession()
    session.insertContent('x'.repeat(40_000), { slide: 0 })
    const text = session.readSlide(0)
    expect(text).toContain('output truncated at 30000 characters')
  })
})

describe('slides session insert_content', () => {
  it('adds a text box with default geometry and round-trips it through save', async () => {
    const session = await openSession()
    const result = session.insertContent('Added by agent\nSecond paragraph', { slide: 0 })
    expect(result.added).toBe(true)
    expect(result.element).toBe(2)
    expect(result.paragraphs).toBe(2)
    expect(session.meta().dirty).toBe(true)

    await session.save()
    const reopened = await reopen(deckPath)
    expect(reopened.deck.slides[0]!.elements).toHaveLength(3)
    const added = reopened.deck.slides[0]!.elements[2] as TextElement
    expect(added.type).toBe('text')
    // default box: 6 x 1 in at (1", 1") in EMU
    expect(added.transform.offset).toMatchObject({
      x: 914_400,
      y: 914_400,
      cx: 5_486_400,
      cy: 914_400,
    })
    expect(slideTextOf(reopened)).toContain('Added by agent\nSecond paragraph')
  })

  it('places a text box at explicit inch coordinates', async () => {
    const session = await openSession()
    session.insertContent('Positioned', { slide: 0, x: 2, y: 3, width: 4, height: 0.5 })
    await session.save()
    const added = (await reopen(deckPath)).deck.slides[0]!.elements[2] as TextElement
    expect(added.transform.offset).toMatchObject({
      x: 1_828_800,
      y: 2_743_200,
      cx: 3_657_600,
      cy: 457_200,
    })
  })

  it('replaces the text of an existing element and round-trips it', async () => {
    const session = await openSession()
    const result = session.insertContent('Rewritten\nshape text', { slide: 0, element: 1 })
    expect(result.added).toBe(false)
    await session.save()
    const reopened = await reopen(deckPath)
    expect(reopened.deck.slides[0]!.elements).toHaveLength(2)
    expect(slideTextOf(reopened)).toBe('Quarterly Review\nRewritten\nshape text')
  })

  it('creates a centered body for a bare autoshape, like the app first-text defaults', async () => {
    // strip the txBody from the roundRect so the session hits the fresh-body path
    const opened = await openPptx(fixture)
    const shape = opened.deck.slides[0]!.elements[1] as TextElement
    shape.anchor.originalXml = shape.anchor.originalXml.replace(/<p:txBody>[\s\S]*<\/p:txBody>/, '')
    delete shape.text
    // flag the element dirty: a clean slide saves the archive's original bytes,
    // and only the dirty path splices the (stripped) anchor XML back in
    shape.dirty = true
    const { savePptx } = await import('@airy-office/pptx-engine')
    const bare = join(root, 'bare.pptx')
    await writeFile(bare, await savePptx(opened))

    const session = await SlidesSession.open(bare, root)
    session.insertContent('Fresh shape text', { slide: 0, element: 1 })
    await session.save()
    const reopened = await reopen(bare)
    const el = reopened.deck.slides[0]!.elements[1] as TextElement
    expect(el.text?.paragraphs[0]?.runs[0]?.text).toBe('Fresh shape text')
    expect(el.text?.paragraphs[0]?.align).toBe('center')
    expect(el.text?.anchor).toBe('middle')
  })

  it('refuses a connector, which cannot hold text', async () => {
    const opened = await openPptx(fixture)
    addElement(opened.deck.slides[0]!, {
      kind: 'line',
      offset: { x: 100, y: 100, cx: 2_000, cy: 0 },
    })
    const { savePptx } = await import('@airy-office/pptx-engine')
    const withLine = join(root, 'line.pptx')
    await writeFile(withLine, await savePptx(opened))

    const session = await SlidesSession.open(withLine, root)
    expect(() => session.insertContent('no', { slide: 0, element: 2 })).toThrow(/connector/)
  })

  it('rejects bad slide/element indexes and payload caps', async () => {
    const session = await openSession()
    expect(() => session.insertContent('x', { slide: 9 })).toThrow(/slide index 9 is out of range/)
    expect(() => session.insertContent('x', { slide: 0, element: 5 })).toThrow(
      /element index 5 is out of range/,
    )
    expect(() => session.insertContent('', { slide: 0 })).toThrow(/text is required/)
    expect(() => session.insertContent('x'.repeat(200_001), { slide: 0 })).toThrow(
      /the cap is 200000/,
    )
  })
})

describe('slides session save fences', () => {
  it('round-trips the original bytes verbatim on a zero-edit save', async () => {
    const session = await openSession()
    const result = await session.save()
    expect(result.unchanged).toBe(true)
    expect(result.format).toBe('pptx')
    expect(await readFile(deckPath)).toEqual(Buffer.from(fixture))
  })

  it('saves edits in place and keeps the deck reopenable', async () => {
    const session = await openSession()
    session.insertContent('Persisted', { slide: 0 })
    const result = await session.save()
    expect(result.unchanged).toBe(false)
    expect(result.path).toBe(deckPath)
    expect(slideTextOf(await reopen(deckPath))).toContain('Persisted')
  })

  it('refuses to save over a file that changed on disk since open', async () => {
    const session = await openSession()
    session.insertContent('Edit', { slide: 0 })
    // external writer: same size, bumped mtime (mtimeMs has ms resolution)
    const before = await stat(deckPath)
    await utimes(deckPath, new Date(), new Date(before.mtimeMs + 2000))
    await expect(session.save()).rejects.toThrow(FencingError)
  })

  it('refuses an existing save-as target unless overwrite is passed', async () => {
    const session = await openSession()
    session.insertContent('Edit', { slide: 0 })
    const target = join(root, 'copy.pptx')
    await writeFile(target, 'unrelated')
    await expect(session.save(target)).rejects.toThrow(/already exists/)
    expect(await readFile(target, 'utf8')).toBe('unrelated')
    await expect(session.save(target, { overwrite: true })).resolves.toMatchObject({
      path: target,
    })
    expect(slideTextOf(await reopen(target))).toContain('Edit')
  })

  it('promotes a fresh save-as target without overwrite', async () => {
    const session = await openSession()
    session.insertContent('Fresh target', { slide: 0 })
    const target = join(root, 'fresh.pptx')
    const result = await session.save(target)
    expect(result.path).toBe(target)
    // a second save to the session's own output needs no overwrite
    await expect(session.save(target)).resolves.toMatchObject({ path: target })
  })

  it('refuses to save when the pinned workspace root disappeared', async () => {
    const nested = join(root, 'nested')
    const { mkdir } = await import('node:fs/promises')
    await mkdir(nested)
    const nestedDeck = join(nested, 'deck.pptx')
    await writeFile(nestedDeck, fixture)
    const session = await SlidesSession.open(nestedDeck, nested)
    await rename(nested, join(root, 'renamed'))
    await expect(session.save()).rejects.toThrow(/no longer exists/)
  })

  it('keeps saves confined to the workspace root', async () => {
    const session = await openSession()
    await expect(session.save('../../escape.pptx')).rejects.toThrow(/outside the workspace root/)
  })

  it('close is a no-op with nothing to clean', async () => {
    const session = await openSession()
    expect(await session.close()).toEqual([])
  })
})
