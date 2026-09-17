/**
 * Slide comments (classic comments part) read/write — archive surgery.
 *
 * Structure: ppt/commentAuthors.xml (author table, referenced by
 * presentation.xml.rels) + one ppt/comments/commentN.xml per slide (referenced by
 * slide rels). A comment is uniquely identified by (authorId, idx). All changes
 * land in archive.entries: savePptx persists automatically, and the main process's
 * snapshot-style undo covers them automatically.
 *
 * Threading + resolve state: PowerPoint stores replies and the "done" flag out
 * of the flat comments part. We write/read two complementary representations:
 * - the ECMA-documented [MS-PPTX] §2.2.10 threadingInfo ext on each reply's
 *   <p:cm> (uri {C676402C-…}, p15:parentCm with the full parent authorId+idx);
 * - the de-facto ppt/commentsExtended/commentN.xml part (p15:cmEx entries keyed
 *   by (authorId, idx) carrying p15:parentId, p15:done and timestamps), related
 *   from the comments part. The resolve flag has no home outside it, so the part
 *   is created lazily — only once a reply or resolve actually exists, keeping
 *   commentless and flat-comment decks byte-identical.
 */
import type { OpenedPptx } from './index'
import { resolveTarget, type PackageArchive } from './zip'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'
import { appendRelationship, unescapeXml } from './notes'
import { removeRelationshipAndCollectOwnedTarget } from './resource-cleanup'

const XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const NS_P15 = 'http://schemas.microsoft.com/office/powerpoint/2012/main'

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const COMMENTS_REL = `${REL_BASE}/comments`
const AUTHORS_REL = `${REL_BASE}/commentAuthors`

const COMMENTS_CT = 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml'
const AUTHORS_CT = 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml'
const COMMENTS_EXT_CT =
  'application/vnd.openxmlformats-officedocument.presentationml.commentsExtended+xml'

const AUTHORS_PATH = 'ppt/commentAuthors.xml'

/** the comments→commentsExtended relationship: both spellings occur in the wild */
const COMMENTS_EXT_RELS = [
  'http://schemas.microsoft.com/office/2011/relationships/commentsExt',
  'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
]

/** [MS-PPTX] §2.2.10: ext uri that hosts p15:threadingInfo inside a reply's <p:cm> */
const THREADING_EXT_URI = '{C676402C-5697-4E1C-873F-D02D1690AC5C}'

export interface SlideComment {
  authorId: number
  author: string
  initials: string
  /** ISO timestamp (the dt attribute in the part) */
  dt: string
  idx: number
  text: string
  /** set on replies: identifies the parent comment (threaded comments) */
  parentId?: { authorId: number; idx: number }
  /** thread resolve state (commentsExtended p15:done) */
  resolved?: boolean
}

/** identity of one comment (the pair that PowerPoint uses everywhere) */
export interface CommentRef {
  authorId: number
  idx: number
}

function setEntry(archive: PackageArchive, path: string, xml: string): void {
  archive.entries.set(path, Buffer.from(xml, 'utf8'))
}

function addContentTypeOverride(
  archive: PackageArchive,
  partPath: string,
  contentType: string,
): void {
  const ctPath = '[Content_Types].xml'
  const ct = archive.readText(ctPath)
  if (!ct || ct.includes(`PartName="/${partPath}"`)) return
  setEntry(
    archive,
    ctPath,
    ct.replace(
      '</Types>',
      `<Override PartName="/${partPath}" ContentType="${contentType}"/></Types>`,
    ),
  )
}

function removeContentTypeOverride(archive: PackageArchive, partPath: string): void {
  const ctPath = '[Content_Types].xml'
  const ct = archive.readText(ctPath)
  if (!ct) return
  setEntry(archive, ctPath, ct.replace(new RegExp(`<Override PartName="/${partPath}"[^>]*/>`), ''))
}

/** Author table: id → {name, initials}. */
function readAuthors(archive: PackageArchive): Map<number, { name: string; initials: string }> {
  const map = new Map<number, { name: string; initials: string }>()
  const xml = archive.readText(AUTHORS_PATH)
  if (!xml) return map
  for (const m of xml.matchAll(/<p:cmAuthor\b([^>]*)\/>/g)) {
    const attrs = m[1]!
    const id = Number(/\bid="(\d+)"/.exec(attrs)?.[1] ?? -1)
    if (id < 0) continue
    map.set(id, {
      name: unescapeXml(/\bname="([^"]*)"/.exec(attrs)?.[1] ?? ''),
      initials: unescapeXml(/\binitials="([^"]*)"/.exec(attrs)?.[1] ?? ''),
    })
  }
  return map
}

/** Relationship and path of a slide's comments part (null when absent). */
function commentsRelationshipForSlide(
  archive: PackageArchive,
  slidePath: string,
): { id: string; path: string } | null {
  for (const rel of archive.readRels(slidePath).values()) {
    if (rel.type === COMMENTS_REL) {
      return { id: rel.id, path: resolveTarget(slidePath, rel.target) }
    }
  }
  return null
}

/** Path of a slide's comments part (null if the slide has no comments). */
function commentsPathForSlide(archive: PackageArchive, slidePath: string): string | null {
  return commentsRelationshipForSlide(archive, slidePath)?.path ?? null
}

/** Read all comments on a slide (in order of appearance), with thread/resolve state. */
export function getSlideComments(archive: PackageArchive, slidePath: string): SlideComment[] {
  const partPath = commentsPathForSlide(archive, slidePath)
  if (!partPath) return []
  const xml = archive.readText(partPath)
  if (!xml) return []
  const authors = readAuthors(archive)
  const out: SlideComment[] = []
  for (const m of xml.matchAll(/<p:cm\b([^>]*)>([\s\S]*?)<\/p:cm>/g)) {
    const attrs = m[1]!
    const authorId = Number(/\bauthorId="(\d+)"/.exec(attrs)?.[1] ?? 0)
    const a = authors.get(authorId)
    const body = m[2]!
    const comment: SlideComment = {
      authorId,
      author: a?.name ?? 'Unknown author',
      initials: a?.initials ?? '',
      dt: /\bdt="([^"]*)"/.exec(attrs)?.[1] ?? '',
      idx: Number(/\bidx="(\d+)"/.exec(attrs)?.[1] ?? 0),
      text: unescapeXml(/<p:text>([\s\S]*?)<\/p:text>/.exec(body)?.[1] ?? ''),
    }
    // documented thread link: <p:ext uri="{C676402C-…}"><p15:threadingInfo><p15:parentCm …/>
    const parentCm =
      /<p15:parentCm\b([^>]*)\/?>/.exec(
        /<p15:threadingInfo\b[\s\S]*?<\/p15:threadingInfo>/.exec(body)?.[0] ?? '',
      ) ?? null
    if (parentCm) {
      const pAuthor = Number(/(?:p15:)?authorId="(\d+)"/.exec(parentCm[1]!)?.[1] ?? -1)
      const pIdx = Number(/(?:p15:)?idx="(\d+)"/.exec(parentCm[1]!)?.[1] ?? -1)
      if (pAuthor >= 0 && pIdx >= 0) comment.parentId = { authorId: pAuthor, idx: pIdx }
    }
    out.push(comment)
  }
  mergeCommentsExtended(archive, partPath, out)
  return out
}

/** one parsed p15:cmEx entry */
interface CommentExEntry {
  parentIdIdx?: number
  done?: boolean
  createdAt?: string
}

/** attributes of a p15:cmEx start tag, tolerant of the p15: prefix and both rel spellings */
function parseCmExAttrs(attrs: string): { authorId: number; idx: number } & CommentExEntry {
  const num = (name: string) =>
    Number(new RegExp(`(?:p15:)?${name}="(\\d+)"`).exec(attrs)?.[1] ?? -1)
  const iso = (name: string) =>
    new RegExp(`(?:p15:)?${name}="([^"]*)"`).exec(attrs)?.[1] ?? undefined
  return {
    authorId: num('authorId'),
    idx: num('idx'),
    parentIdIdx: attrs.includes('parentId="') ? num('parentId') : undefined,
    done: /(?:p15:)?done="(?:1|true)"/.test(attrs) ? true : undefined,
    createdAt: iso('createdAt'),
  }
}

/**
 * Overlay the commentsExtended part onto the parsed comments: resolve state
 * always; parent linkage only for replies that lack the (preferred,
 * unambiguous) threadingInfo ext. A numeric p15:parentId names the parent's idx;
 * when several authors could own that idx, the same-author non-reply comment wins.
 */
function mergeCommentsExtended(
  archive: PackageArchive,
  commentsPartPath: string,
  comments: SlideComment[],
): void {
  const rel = commentsExtendedRelationship(archive, commentsPartPath)
  const xml = rel ? archive.readText(rel.path) : null
  if (!xml) return
  const ex = new Map<string, CommentExEntry>()
  // both shapes occur: self-closing entries and entries wrapping a p15:extLst
  for (const m of xml.matchAll(/<p15:cmEx\b([^>]*?)(?:\/>|>([\s\S]*?)<\/p15:cmEx>)/g)) {
    const e = parseCmExAttrs(m[1]!)
    if (e.authorId >= 0 && e.idx >= 0) ex.set(`${e.authorId}:${e.idx}`, e)
  }
  for (const c of comments) {
    const e = ex.get(`${c.authorId}:${c.idx}`)
    if (!e) continue
    if (e.done) c.resolved = true
    if (!c.parentId && e.parentIdIdx !== undefined && e.parentIdIdx >= 0) {
      const parent =
        comments.find((p) => p.authorId === c.authorId && p.idx === e.parentIdIdx && !p.parentId) ??
        comments.find((p) => p.idx === e.parentIdIdx && !p.parentId && p !== c)
      if (parent) c.parentId = { authorId: parent.authorId, idx: parent.idx }
    }
  }
}

/** Relationship and path of the commentsExtended part of a comments part (null when absent). */
function commentsExtendedRelationship(
  archive: PackageArchive,
  commentsPartPath: string,
): { id: string; path: string } | null {
  for (const rel of archive.readRels(commentsPartPath).values()) {
    if (COMMENTS_EXT_RELS.includes(rel.type)) {
      return { id: rel.id, path: resolveTarget(commentsPartPath, rel.target) }
    }
  }
  return null
}

/** p15:cmEx entry XML for one comment (self-closing, attribute-only). */
function cmExXml(c: SlideComment, times?: { createdAt?: string }): string {
  const createdAt = times?.createdAt ?? c.dt ?? '1970-01-01T00:00:00.000Z'
  const attrs = [
    `p15:authorId="${c.authorId}"`,
    `p15:idx="${c.idx}"`,
    c.parentId ? `p15:parentId="${c.parentId.idx}"` : '',
    `p15:createdAt="${createdAt || '1970-01-01T00:00:00.000Z'}"`,
    `p15:updatedAt="${c.dt || createdAt || '1970-01-01T00:00:00.000Z'}"`,
    c.resolved ? 'p15:done="1"' : '',
  ].filter(Boolean)
  return `<p15:cmEx ${attrs.join(' ')}/>`
}

/** commentsExtended part path matching the comments part's number (comment3.xml → comment3.xml). */
function commentsExtendedPathFor(commentsPartPath: string): string {
  const num = /comment(\d+)\.xml$/.exec(commentsPartPath)?.[1] ?? '1'
  return `ppt/commentsExtended/comment${num}.xml`
}

/**
 * Ensure the commentsExtended part exists (content type + relationship from the
 * comments part) and return its path.
 */
function ensureCommentsExtended(archive: PackageArchive, commentsPartPath: string): string {
  const existing = commentsExtendedRelationship(archive, commentsPartPath)
  if (existing) return existing.path
  const path = commentsExtendedPathFor(commentsPartPath)
  setEntry(archive, path, XMLDECL + `<p15:cmLst xmlns:p15="${NS_P15}"></p15:cmLst>`)
  addContentTypeOverride(archive, path, COMMENTS_EXT_CT)
  appendRelationship(
    archive,
    commentsPartPath,
    COMMENTS_EXT_RELS[0]!,
    `../commentsExtended/${path.split('/').pop()!}`,
  )
  return path
}

/** Drop the commentsExtended part, its relationship and override (nothing left in it). */
function removeCommentsExtended(archive: PackageArchive, commentsPartPath: string): void {
  const rel = commentsExtendedRelationship(archive, commentsPartPath)
  if (!rel) return
  removeRelationshipAndCollectOwnedTarget(archive, commentsPartPath, rel.id)
  archive.entries.delete(rel.path)
  removeContentTypeOverride(archive, rel.path)
}

/**
 * Append a p15:cmEx entry for one comment. The part is created lazily: only when
 * the comment carries thread/resolve state, or the part already exists (PowerPoint
 * writes an entry per comment once the part exists; timestamps then stay present).
 */
function appendCommentEx(archive: PackageArchive, commentsPartPath: string, c: SlideComment): void {
  const existingRel = commentsExtendedRelationship(archive, commentsPartPath)
  if (!existingRel && !c.parentId && !c.resolved) return
  const path = ensureCommentsExtended(archive, commentsPartPath)
  const xml = archive.readText(path)
  if (!xml) return
  setEntry(archive, path, xml.replace(/<\/p15:cmLst>/, `${cmExXml(c)}</p15:cmLst>`))
}

/** Remove one comment's p15:cmEx entry; drops the emptied part (and its wiring). */
function removeCommentEx(archive: PackageArchive, commentsPartPath: string, ref: CommentRef): void {
  const rel = commentsExtendedRelationship(archive, commentsPartPath)
  if (!rel) return
  const xml = archive.readText(rel.path)
  if (!xml) return
  // covers self-closing entries and entries wrapping a p15:extLst
  for (const m of xml.matchAll(/<p15:cmEx\b([^>]*?)(?:\/>|>([\s\S]*?)<\/p15:cmEx>)/g)) {
    const e = parseCmExAttrs(m[1]!)
    if (e.authorId === ref.authorId && e.idx === ref.idx) {
      const next = xml.slice(0, m.index!) + xml.slice(m.index! + m[0].length)
      setEntry(archive, rel.path, next)
      if (!/<p15:cmEx\b/.test(next)) removeCommentsExtended(archive, commentsPartPath)
      return
    }
  }
}

/** thread/resolve state of one existing comment (by ref), or null when absent */
function findComment(comments: SlideComment[], ref: CommentRef): SlideComment | undefined {
  return comments.find((c) => c.authorId === ref.authorId && c.idx === ref.idx)
}

/**
 * Set the resolve flag on one comment (p15:done in commentsExtended). Creates the
 * part — and the comment's entry — when missing. Existing entries keep their
 * createdAt; updatedAt follows the comment's dt.
 */
export function setSlideCommentResolved(
  opened: OpenedPptx,
  slideIndex: number,
  ref: CommentRef,
  done: boolean,
): boolean {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return false
  const { archive } = opened
  const commentsPartPath = commentsPathForSlide(archive, slide.path)
  if (!commentsPartPath) return false
  const target = findComment(getSlideComments(archive, slide.path), ref)
  if (!target) return false

  const rel = commentsExtendedRelationship(archive, commentsPartPath)
  const path = rel?.path ?? ensureCommentsExtended(archive, commentsPartPath)
  const xml = archive.readText(path)
  if (!xml) return false

  // rewrite this comment's entry (or append one) with the flag applied
  let rewritten = false
  const out = xml.replace(
    /<p15:cmEx\b([^>]*?)(?:\/>|>([\s\S]*?)<\/p15:cmEx>)/g,
    (full: string, attrs: string) => {
      const e = parseCmExAttrs(attrs)
      if (e.authorId !== ref.authorId || e.idx !== ref.idx) return full
      rewritten = true
      return cmExXml({ ...target, resolved: done }, { createdAt: e.createdAt ?? target.dt })
    },
  )
  if (!rewritten)
    setEntry(
      archive,
      path,
      out.replace(/<\/p15:cmLst>/, `${cmExXml({ ...target, resolved: done })}</p15:cmLst>`),
    )
  else setEntry(archive, path, out)
  return true
}

/**
 * Ensure the author exists and return {authorId, nextIdx} (bumping lastIdx).
 * Creates the author table / content type / presentation rel if missing.
 */
function ensureAuthor(
  archive: PackageArchive,
  name: string,
  initials: string,
): { authorId: number; nextIdx: number } {
  let xml = archive.readText(AUTHORS_PATH)
  if (!xml) {
    xml = XMLDECL + `<p:cmAuthorLst xmlns:p="${NS_P}"></p:cmAuthorLst>`
    addContentTypeOverride(archive, AUTHORS_PATH, AUTHORS_CT)
    appendRelationship(archive, 'ppt/presentation.xml', AUTHORS_REL, 'commentAuthors.xml')
  }
  // Existing author with the same name: reuse its id, lastIdx + 1
  for (const m of xml.matchAll(/<p:cmAuthor\b([^>]*)\/>/g)) {
    const attrs = m[1]!
    if (unescapeXml(/\bname="([^"]*)"/.exec(attrs)?.[1] ?? '') !== name) continue
    const authorId = Number(/\bid="(\d+)"/.exec(attrs)?.[1] ?? 0)
    const lastIdx = Number(/\blastIdx="(\d+)"/.exec(attrs)?.[1] ?? 0)
    const nextIdx = lastIdx + 1
    const bumped = m[0].includes('lastIdx="')
      ? m[0].replace(/\blastIdx="\d+"/, `lastIdx="${nextIdx}"`)
      : m[0].replace('/>', ` lastIdx="${nextIdx}"/>`)
    setEntry(archive, AUTHORS_PATH, xml.replace(m[0], bumped))
    return { authorId, nextIdx }
  }
  // New author
  let maxId = -1
  for (const m of xml.matchAll(/<p:cmAuthor\b[^>]*\bid="(\d+)"/g))
    maxId = Math.max(maxId, Number(m[1]))
  const authorId = maxId + 1
  const tag =
    `<p:cmAuthor id="${authorId}" name="${escapeXmlAttr(name)}"` +
    ` initials="${escapeXmlAttr(initials)}" lastIdx="1" clrIdx="${authorId}"/>`
  setEntry(archive, AUTHORS_PATH, xml.replace('</p:cmAuthorLst>', `${tag}</p:cmAuthorLst>`))
  return { authorId, nextIdx: 1 }
}

/** Add a comment (optionally a reply to an existing one); returns the new comment (with assigned authorId/idx). */
export function addSlideComment(
  opened: OpenedPptx,
  slideIndex: number,
  opts: { author: string; initials?: string; text: string; parent?: CommentRef },
): SlideComment | null {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return null
  const { archive } = opened
  const initials =
    opts.initials ??
    opts.author
      .split(/\s+/)
      .map((w) => w[0] ?? '')
      .join('')
      .toUpperCase() ??
    ''
  const { authorId, nextIdx } = ensureAuthor(archive, opts.author, initials)

  let partPath = commentsPathForSlide(archive, slide.path)
  if (!partPath) {
    let maxNum = 0
    for (const path of archive.entries.keys()) {
      const m = /^ppt\/comments\/comment(\d+)\.xml$/.exec(path)
      if (m) maxNum = Math.max(maxNum, Number(m[1]))
    }
    partPath = `ppt/comments/comment${maxNum + 1}.xml`
    setEntry(archive, partPath, XMLDECL + `<p:cmLst xmlns:p="${NS_P}"></p:cmLst>`)
    addContentTypeOverride(archive, partPath, COMMENTS_CT)
    appendRelationship(archive, slide.path, COMMENTS_REL, `../${partPath.slice(4)}`)
  }
  const xml = archive.readText(partPath)
  if (!xml) return null

  const dt = new Date().toISOString()
  // Positions staggered along the top-left diagonal so comment badges don't overlap
  const count = [...xml.matchAll(/<p:cm\b/g)].length
  const pos = 10 + (count % 8) * 6
  // replies carry the documented threadingInfo ext pointing at the parent's
  // (authorId, idx) — the unambiguous thread link PowerPoint 2013+ honors
  const threading = opts.parent
    ? `<p:extLst><p:ext uri="${THREADING_EXT_URI}">` +
      `<p15:threadingInfo xmlns:p15="${NS_P15}">` +
      `<p15:parentCm authorId="${opts.parent.authorId}" idx="${opts.parent.idx}"/>` +
      `</p15:threadingInfo></p:ext></p:extLst>`
    : ''
  const cm =
    `<p:cm authorId="${authorId}" dt="${dt}" idx="${nextIdx}">` +
    `<p:pos x="${pos}" y="${pos}"/>` +
    `<p:text>${escapeXmlText(opts.text)}</p:text>${threading}</p:cm>`
  setEntry(archive, partPath, xml.replace('</p:cmLst>', `${cm}</p:cmLst>`))

  const created: SlideComment = {
    authorId,
    author: opts.author,
    initials,
    dt,
    idx: nextIdx,
    text: opts.text,
    ...(opts.parent ? { parentId: { ...opts.parent } } : {}),
  }
  // mirror the reply link (and any resolve state) in commentsExtended
  appendCommentEx(archive, partPath, created)
  return created
}

/** Delete a comment (by authorId + idx); replies are left to the caller (cascade policy). */
export function deleteSlideComment(
  opened: OpenedPptx,
  slideIndex: number,
  ref: CommentRef,
): boolean {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return false
  const { archive } = opened
  const relationship = commentsRelationshipForSlide(archive, slide.path)
  if (!relationship) return false
  const partPath = relationship.path
  const xml = archive.readText(partPath)
  if (!xml) return false
  for (const m of xml.matchAll(/<p:cm\b([^>]*)>[\s\S]*?<\/p:cm>/g)) {
    const attrs = m[1]!
    if (
      Number(/\bauthorId="(\d+)"/.exec(attrs)?.[1] ?? -1) === ref.authorId &&
      Number(/\bidx="(\d+)"/.exec(attrs)?.[1] ?? -1) === ref.idx
    ) {
      removeCommentEx(archive, partPath, ref)
      const next = xml.slice(0, m.index!) + xml.slice(m.index! + m[0].length)
      setEntry(archive, partPath, next)
      if (!/<p:cm\b/.test(next)) {
        removeCommentsExtended(archive, partPath)
        removeRelationshipAndCollectOwnedTarget(archive, slide.path, relationship.id)
      }
      return true
    }
  }
  return false
}
