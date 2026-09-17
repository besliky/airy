/**
 * Threaded comments: replies (parentId) and resolve state survive save → reopen
 * through the commentsExtended part + the threadingInfo ext; comment edits leave
 * every unrelated part byte-identical.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  addSlideComment,
  createBlankPptx,
  deleteSlideComment,
  getSlideComments,
  openPptx,
  savePptx,
  setSlideCommentResolved,
} from '../src/index'
import { PackageArchive, relsPathFor } from '../src/zip'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('threaded comments round-trip', () => {
  it('threads and resolve state survive save → reopen', async () => {
    const opened = await openPptx(await createBlankPptx())

    const head = addSlideComment(opened, 0, { author: 'Carol', text: 'Rework the title' })!
    const reply = addSlideComment(opened, 0, {
      author: 'Dave',
      text: 'Agreed &lt; keeping the serif',
      parent: { authorId: head.authorId, idx: head.idx },
    })!
    expect(reply.parentId).toEqual({ authorId: head.authorId, idx: head.idx })

    expect(
      setSlideCommentResolved(opened, 0, { authorId: head.authorId, idx: head.idx }, true),
    ).toBe(true)
    expect(
      setSlideCommentResolved(opened, 0, { authorId: reply.authorId, idx: reply.idx }, true),
    ).toBe(true)

    const reopened = await openPptx(await savePptx(opened))
    const slidePath2 = reopened.deck.slides[0]!.path
    const list = getSlideComments(reopened.archive, slidePath2)
    expect(list).toHaveLength(2)
    expect(list[0]!.text).toBe('Rework the title')
    expect(list[0]!.resolved).toBe(true)
    expect(list[1]!.text).toBe('Agreed &lt; keeping the serif')
    expect(list[1]!.parentId).toEqual({ authorId: head.authorId, idx: head.idx })
    expect(list[1]!.resolved).toBe(true)

    // the extension part is wired the way PowerPoint expects: content type
    // override + relationship from the comments part
    const commentsRel = [...reopened.archive.readRels(slidePath2).values()].find((rel) =>
      rel.type.endsWith('/comments'),
    )!
    const commentsPath = commentsRel.target.replace(/^\.\.\//, 'ppt/')
    const extRel = [...reopened.archive.readRels(commentsPath).values()].find((rel) =>
      rel.type.includes('/commentsExt'),
    )!
    expect(extRel.target).toBe('../commentsExtended/comment1.xml')
    expect(reopened.archive.has('ppt/commentsExtended/comment1.xml')).toBe(true)
    expect(reopened.archive.readText('[Content_Types].xml')!).toContain(
      'presentationml.commentsExtended+xml',
    )
    const extXml = reopened.archive.readText('ppt/commentsExtended/comment1.xml')!
    expect(extXml).toContain('http://schemas.microsoft.com/office/powerpoint/2012/main')
    expect(extXml).toMatch(/<p15:cmEx\b[^>]*p15:parentId="\d+"/)
    expect(extXml).toMatch(/<p15:cmEx\b[^>]*p15:done="1"/)
    // the documented thread link rides on the reply's p:cm as well
    const commentsXml = reopened.archive.readText(commentsPath)!
    expect(commentsXml).toContain('{C676402C-5697-4E1C-873F-D02D1690AC5C}')
    expect(commentsXml).toContain('<p15:parentCm ')
  })

  it('reopening clears the done flag without losing the thread', async () => {
    const opened = await openPptx(await createBlankPptx())
    const head = addSlideComment(opened, 0, { author: 'Carol', text: 'head' })!
    const reply = addSlideComment(opened, 0, {
      author: 'Dave',
      text: 'reply',
      parent: { authorId: head.authorId, idx: head.idx },
    })!
    setSlideCommentResolved(opened, 0, { authorId: head.authorId, idx: head.idx }, true)
    setSlideCommentResolved(opened, 0, { authorId: reply.authorId, idx: reply.idx }, true)
    setSlideCommentResolved(opened, 0, { authorId: head.authorId, idx: head.idx }, false)
    setSlideCommentResolved(opened, 0, { authorId: reply.authorId, idx: reply.idx }, false)

    const reopened = await openPptx(await savePptx(opened))
    const list = getSlideComments(reopened.archive, reopened.deck.slides[0]!.path)
    expect(list.every((c) => !c.resolved)).toBe(true)
    expect(list[1]!.parentId).toEqual({ authorId: head.authorId, idx: head.idx })
  })

  it('reads commentsExtended-only threads (no threadingInfo ext)', async () => {
    // simulate a PowerPoint-authored file: flat comments part + extended part,
    // no ext on the p:cm elements
    const opened = await openPptx(await createBlankPptx())
    const a = addSlideComment(opened, 0, { author: 'Carol', text: 'head' })!
    const b = addSlideComment(opened, 0, {
      author: 'Carol',
      text: 'reply',
      parent: { authorId: a.authorId, idx: a.idx },
    })!
    const slide = opened.deck.slides[0]!
    const rel = [...opened.archive.readRels(slide.path).values()].find((r) =>
      r.type.endsWith('/comments'),
    )!
    const commentsPath = rel.target.replace(/^\.\.\//, 'ppt/')
    // strip the threadingInfo ext the writer adds, then rewrite the extended
    // part as PowerPoint would have written it (flat entries + done)
    const stripped = opened.archive
      .readText(commentsPath)!
      .replace(/<p:extLst>[\s\S]*?<\/p:extLst>/, '')
    opened.archive.entries.set(commentsPath, Buffer.from(stripped, 'utf8'))
    const extPath = 'ppt/commentsExtended/comment1.xml'
    const extXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<p15:cmLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main">' +
      `<p15:cmEx p15:authorId="${a.authorId}" p15:idx="${a.idx}" p15:createdAt="${a.dt}" p15:updatedAt="${a.dt}" p15:done="1"/>` +
      `<p15:cmEx p15:authorId="${b.authorId}" p15:idx="${b.idx}" p15:parentId="${a.idx}" p15:createdAt="${b.dt}" p15:updatedAt="${b.dt}"/>` +
      '</p15:cmLst>'
    opened.archive.entries.set(extPath, Buffer.from(extXml, 'utf8'))

    const list = getSlideComments(opened.archive, slide.path)
    expect(list[1]!.parentId).toEqual({ authorId: a.authorId, idx: a.idx })
    expect(list[0]!.resolved).toBe(true)
    expect(list[1]!.resolved).toBeUndefined()
  })

  it('deleting a thread head keeps working (caller cascades replies)', async () => {
    const opened = await openPptx(await createBlankPptx())
    const head = addSlideComment(opened, 0, { author: 'Carol', text: 'head' })!
    const reply = addSlideComment(opened, 0, {
      author: 'Dave',
      text: 'reply',
      parent: { authorId: head.authorId, idx: head.idx },
    })!
    deleteSlideComment(opened, 0, reply)
    deleteSlideComment(opened, 0, head)

    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideComments(reopened.archive, reopened.deck.slides[0]!.path)).toEqual([])
    expect(reopened.archive.has('ppt/comments/comment1.xml')).toBe(false)
    expect(reopened.archive.has('ppt/commentsExtended/comment1.xml')).toBe(false)
  })
})

describe('byte preservation', () => {
  it('a commentless file round-trips byte-for-byte per entry', async () => {
    const bytes = fx('01_standard_business.pptx')
    const reopened = await openPptx(await savePptx(await openPptx(bytes)))
    const before = await PackageArchive.open(bytes)
    expect([...reopened.archive.entries.keys()].sort()).toEqual([...before.entries.keys()].sort())
    for (const [name, buf] of before.entries) {
      expect(
        Buffer.compare(Buffer.from(buf), Buffer.from(reopened.archive.entries.get(name)!)),
      ).toBe(0)
    }
  })

  it('comment edits leave unrelated parts untouched', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const before = new Map(opened.archive.entries)

    const head = addSlideComment(opened, 0, { author: 'Carol', text: 'head' })!
    const reply = addSlideComment(opened, 0, {
      author: 'Dave',
      text: 'reply',
      parent: { authorId: head.authorId, idx: head.idx },
    })!
    setSlideCommentResolved(opened, 0, { authorId: head.authorId, idx: head.idx }, true)
    setSlideCommentResolved(opened, 0, { authorId: reply.authorId, idx: reply.idx }, true)

    const slide = opened.deck.slides[0]!
    const commentsRel = [...opened.archive.readRels(slide.path).values()].find((r) =>
      r.type.endsWith('/comments'),
    )!
    const commentsPath = commentsRel.target.replace(/^\.\.\//, 'ppt/')
    const touched = new Set([
      '[Content_Types].xml',
      'ppt/commentAuthors.xml',
      relsPathFor('ppt/presentation.xml'),
      commentsPath,
      relsPathFor(commentsPath),
      'ppt/commentsExtended/comment1.xml',
      relsPathFor(slide.path),
    ])
    for (const [name, buf] of before) {
      if (touched.has(name)) continue
      expect(Buffer.compare(Buffer.from(buf), Buffer.from(opened.archive.entries.get(name)!))).toBe(
        0,
      )
    }
  })
})
