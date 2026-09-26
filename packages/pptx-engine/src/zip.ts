/**
 * pptx package management — open the zip, archive the original by SHA-256, and read
 * parts and .rels.
 *
 * Byte fidelity: PackageArchive holds the original bytes of every entry; on save,
 * unmodified entries are written back byte-for-byte (handled by the patch layer).
 * This module only handles reading and metadata.
 */
import JSZip from 'jszip'
import { createHash } from 'node:crypto'
import { XMLParser } from 'fast-xml-parser'
import type { SlideSize } from './types'
import { asXmlNode, xmlArray } from './xml-utils'

const relsParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  isArray: (name) => name === 'Relationship' || name === 'sldId' || name === 'Override',
})

// Open-time size fences (SEC-1304). The engine is the single open gate shared
// by the desktop app (slides-main reads the file from disk and hands over the
// bytes), the merge/insert paths, and the headless MCP session — which used to
// be the only fenced one (mcp-server size-fence), leaving the desktop open to
// the zip-bomb/raw-size class the docx engine had already closed inside
// parseDocx. The budgets mirror the docx-engine fence (zip-load.ts) and the
// mcp-server fence (size-fence.ts) — same numbers on purpose, do not drift.
const MAX_OPEN_BYTES = 512 * 1024 * 1024
const MAX_ZIP_PARTS = 10_000
const MAX_PART_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
const MAX_TOTAL_UNCOMPRESSED_BYTES = 1.5 * 1024 * 1024 * 1024

/** refuse an oversized raw package before the central directory is parsed */
function assertWithinOpenCap(bytes: Uint8Array): void {
  if (bytes.byteLength <= MAX_OPEN_BYTES) return
  throw new Error(
    `pptx rejected: ${String(bytes.byteLength)} raw bytes exceeds the ` +
      `${String(MAX_OPEN_BYTES)} open cap (512 MiB)`,
  )
}

/**
 * Reject a package whose central directory declares more uncompressed bytes
 * than the budget allows, before any entry is materialized (the open loop
 * inflates every part). Same numbers as the docx-engine and mcp-server
 * fences, so every open path refuses the same bombs.
 */
function assertZipWithinLimits(zip: JSZip): void {
  const files = Object.values(zip.files).filter((f) => !f.dir)
  if (files.length > MAX_ZIP_PARTS) {
    throw new Error(
      `pptx rejected: ${String(files.length)} parts exceeds the ${String(MAX_ZIP_PARTS)} limit`,
    )
  }
  let total = 0
  for (const file of files) {
    // JSZip keeps the declared size on the lazy compressed object without
    // inflating the entry — same seam the docx-engine fence reads. The CD
    // field is unsigned 32-bit but surfaces here as a signed int32, so a
    // declared size ≥ 2 GiB wraps negative: normalize before comparing or
    // the biggest bombs slip through both budgets.
    const declared =
      (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0
    const size = declared < 0 ? declared + 0x1_0000_0000 : declared
    if (size > MAX_PART_UNCOMPRESSED_BYTES) {
      throw new Error(
        `pptx rejected: part ${file.name} declares ${String(size)} uncompressed bytes ` +
          `(limit ${String(MAX_PART_UNCOMPRESSED_BYTES)})`,
      )
    }
    if (size > 0) total += size
  }
  if (total > MAX_TOTAL_UNCOMPRESSED_BYTES) {
    throw new Error(
      `pptx rejected: total uncompressed size ${String(total)} exceeds the ` +
        `${String(MAX_TOTAL_UNCOMPRESSED_BYTES)} limit`,
    )
  }
}

export interface Relationship {
  id: string
  type: string
  target: string
  targetMode?: string
}

export class PackageArchive {
  private constructor(
    private readonly zip: JSZip,
    /** Original bytes of every entry, keyed by path inside the zip */
    readonly entries: Map<string, Uint8Array>,
    /**
     * Original last-modification timestamp of every entry, keyed by path. buildZip
     * writes these back on save: JSZip would otherwise stamp every header with the
     * wall clock, so two saves of the same deck could disagree byte-for-byte
     * (TEST-1747). Entries added after open have no date here.
     */
    readonly dates: Map<string, Date>,
    readonly originalHash: string,
  ) {}

  static async open(bytes: Uint8Array): Promise<PackageArchive> {
    // fences before anything is parsed or inflated (SEC-1304): raw cap first,
    // then the central-directory budget while sizes are still declared-only
    assertWithinOpenCap(bytes)
    const originalHash = createHash('sha256').update(bytes).digest('hex')
    const zip = await JSZip.loadAsync(bytes)
    assertZipWithinLimits(zip)
    const entries = new Map<string, Uint8Array>()
    const dates = new Map<string, Date>()
    const names = Object.keys(zip.files)
    for (const name of names) {
      const file = zip.files[name]
      if (file.dir) continue
      entries.set(name, await file.async('uint8array'))
      dates.set(name, file.date)
    }
    return new PackageArchive(zip, entries, dates, originalHash)
  }

  has(path: string): boolean {
    return this.entries.has(path)
  }

  /** Read a part as a UTF-8 string (for XML parts). */
  readText(path: string): string | null {
    const bytes = this.entries.get(path)
    if (!bytes) return null
    return Buffer.from(bytes).toString('utf8')
  }

  readBytes(path: string): Uint8Array | null {
    return this.entries.get(path) ?? null
  }

  /**
   * Read a part's relationships file. partPath e.g. 'ppt/slides/slide1.xml' →
   * 'ppt/slides/_rels/slide1.xml.rels'.
   */
  readRels(partPath: string): Map<string, Relationship> {
    const relsPath = relsPathFor(partPath)
    const rels = new Map<string, Relationship>()
    const xml = this.readText(relsPath)
    if (!xml) return rels
    const doc = asXmlNode(relsParser.parse(xml))
    const list = asXmlNode(doc.Relationships).Relationship
    for (const r of xmlArray(list)) {
      const id = String(r['@_Id'] ?? '')
      rels.set(id, {
        id,
        type: String(r['@_Type'] ?? ''),
        target: String(r['@_Target'] ?? ''),
        ...(r['@_TargetMode'] != null ? { targetMode: String(r['@_TargetMode']) } : {}),
      })
    }
    return rels
  }

  /**
   * Read the presentation's slide size and the slide part paths in order.
   */
  readPresentation(): { size: SlideSize; slidePaths: string[] } {
    const presXml = this.readText('ppt/presentation.xml')
    if (!presXml) throw new Error('pptx: missing ppt/presentation.xml')

    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
      isArray: (name) => name === 'p:sldId',
    })
    const pres = asXmlNode(parser.parse(presXml))
    const rootRaw = pres['p:presentation'] ?? pres.presentation
    if (!rootRaw) throw new Error('pptx: malformed presentation.xml')
    const root = asXmlNode(rootRaw)

    // Slide size
    const szRaw = root['p:sldSz'] ?? root.sldSz
    const sz = szRaw ? asXmlNode(szRaw) : null
    const size: SlideSize = {
      cx: sz ? parseInt(String(sz['@_cx']), 10) : 9144000,
      cy: sz ? parseInt(String(sz['@_cy']), 10) : 6858000,
    }

    // Slide order: presentation.xml.rels maps r:id to slide parts
    const rels = this.readRels('ppt/presentation.xml')
    const sldIdLst = asXmlNode(root['p:sldIdLst'] ?? root.sldIdLst)
    const slidePaths: string[] = []
    for (const id of xmlArray(sldIdLst['p:sldId'])) {
      const rId = id['@_r:id'] ?? id['@_id']
      if (!rId) continue
      const rel = rels.get(String(rId))
      if (!rel) continue
      slidePaths.push(resolveTarget('ppt/presentation.xml', rel.target))
    }
    return { size, slidePaths }
  }

  /** Resolve a slide's layout / master part paths (via the rels chain). */
  resolveSlideChain(slidePath: string): {
    layoutPath?: string
    masterPath?: string
    themePath?: string
  } {
    const slideRels = this.readRels(slidePath)
    let layoutPath: string | undefined
    for (const rel of slideRels.values()) {
      if (rel.type.endsWith('/slideLayout')) {
        layoutPath = resolveTarget(slidePath, rel.target)
        break
      }
    }
    // Damaged decks ship slides without a rels part (or without the mandatory slideLayout
    // relationship). PowerPoint refuses such files; LibreOffice renders them on the first
    // layout of the first master, which keeps the deck's background/decorations — do the same.
    if (!layoutPath) layoutPath = this.fallbackLayoutPath(slidePath)
    let masterPath: string | undefined
    let themePath: string | undefined
    if (layoutPath) {
      const layoutRels = this.readRels(layoutPath)
      for (const rel of layoutRels.values()) {
        if (rel.type.endsWith('/slideMaster')) {
          masterPath = resolveTarget(layoutPath, rel.target)
          break
        }
      }
    }
    if (masterPath) {
      const masterRels = this.readRels(masterPath)
      for (const rel of masterRels.values()) {
        if (rel.type.endsWith('/theme')) {
          themePath = resolveTarget(masterPath, rel.target)
          break
        }
      }
    }
    return { layoutPath, masterPath, themePath }
  }

  /**
   * Layout for a slide that lost its slideLayout relationship: the first master's layouts in
   * sldLayoutIdLst order, preferring the title layout for a slide carrying a ctrTitle placeholder
   * (what the deck's own title page would have referenced), else the first layout.
   */
  private fallbackLayoutPath(slidePath: string): string | undefined {
    const presXml = this.readText('ppt/presentation.xml')
    if (!presXml) return undefined
    const presRels = this.readRels('ppt/presentation.xml')
    const masterRid = /<p:sldMasterId\b[^>]*\br:id="([^"]+)"/.exec(presXml)?.[1]
    const masterRel = masterRid
      ? presRels.get(masterRid)
      : [...presRels.values()].find((r) => r.type.endsWith('/slideMaster'))
    if (!masterRel) return undefined
    const masterPath = resolveTarget('ppt/presentation.xml', masterRel.target)
    const masterXml = this.readText(masterPath)
    if (!masterXml) return undefined
    const masterRels = this.readRels(masterPath)
    const layouts: string[] = []
    for (const m of masterXml.matchAll(/<p:sldLayoutId\b[^>]*\br:id="([^"]+)"/g)) {
      const rel = masterRels.get(m[1]!)
      if (rel) layouts.push(resolveTarget(masterPath, rel.target))
    }
    if (!layouts.length)
      for (const rel of masterRels.values())
        if (rel.type.endsWith('/slideLayout')) layouts.push(resolveTarget(masterPath, rel.target))
    if (!layouts.length) return undefined
    const wantsTitle = /<p:ph\b[^>]*\btype="ctrTitle"/.test(this.readText(slidePath) ?? '')
    if (wantsTitle) {
      const title = layouts.find((p) =>
        /<p:sldLayout\b[^>]*\btype="title"/.test(this.readText(p) ?? ''),
      )
      if (title) return title
    }
    return layouts[0]
  }
}

/** 'ppt/slides/slide1.xml' → 'ppt/slides/_rels/slide1.xml.rels' */
export function relsPathFor(partPath: string): string {
  const idx = partPath.lastIndexOf('/')
  const dir = idx >= 0 ? partPath.slice(0, idx) : ''
  const file = idx >= 0 ? partPath.slice(idx + 1) : partPath
  return `${dir ? dir + '/' : ''}_rels/${file}.rels`
}

/**
 * Resolve a relative target into an absolute path inside the zip.
 * basePart is the referencing part's path (its directory is the base); target may be
 * something like '../slideLayouts/slideLayout1.xml'.
 */
export function resolveTarget(basePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const baseDir = basePart.slice(0, basePart.lastIndexOf('/'))
  const parts = baseDir.split('/').filter(Boolean)
  for (const seg of target.split('/')) {
    if (seg === '.' || seg === '') continue
    if (seg === '..') parts.pop()
    else parts.push(seg)
  }
  return parts.join('/')
}
