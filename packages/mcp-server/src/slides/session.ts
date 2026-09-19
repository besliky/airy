// Headless slides editing session (PAR-001): one open .pptx -> the engine's
// OpenedPptx model (Slide element trees with byte anchors) -> a save that
// keeps untouched zip entries byte-identical. This is the headless twin of the
// slides app's main-process pipeline, reduced to the engine calls that need
// no renderer:
//
//   open:    openPptx(bytes) -> { deck, archive }
//   edit:    insertContent — addElement (text box) or a plain-text body
//            replace on an existing text/shape element (te.dirty = true);
//            untouched elements keep their original XML slices
//   save:    savePptxToFile (engine-atomic stream) behind the docx session's
//            fences (drift refusal, save-target ownership, stale-root
//            refusal); a zero-edit save writes the original bytes verbatim
//
// Scope decisions (fixed):
// - .pptx only. Legacy .ppt is not converted here: soffice pptx round-trips
//   are too lossy (animations, SmartArt, fonts) to promise the "original
//   untouched parts stay byte-identical" contract, so open_document refuses
//   .ppt/.odp with a conversion hint instead of silently degrading a deck.
// - No apply_ops: the slides op vocabulary (setText/setFont/setParagraphFormat/
//   ...) is an app-side registry over rich EditParagraph records; a headless
//   port is its own PAR. insert_content covers the practical text paths.
import { randomUUID } from 'node:crypto'
import { mkdir, readFile, rename, stat, writeFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'

import {
  addElement,
  commitSaved,
  DEFAULT_BODY_INSETS,
  isConnectorXml,
  openPptx,
  savePptxToFile,
  type OpenedPptx,
  type Paragraph,
  type Slide,
  type SlideElement,
  type TableElement,
  type TextElement,
} from '@airy-office/pptx-engine'

import {
  assertSaveTargetFree,
  countWords,
  FencingError,
  promoteNewFileExclusively,
} from '../docx/session.js'
import { assertWorkspaceRootExists, resolveConfined, workspaceRoot } from '../docx/paths.js'

// ---- limits (mirror the docx session, scaled to the MCP 30k answer budget) ----

const CONTEXT_MAX_CHARS = 30_000
// Test seam: tripping the absolute 30k overview budget needs a ~500-slide
// fixture whose duplicateSlide/save round-trip costs seconds of honest CPU
// (PERF-1101). The budget ladder is linear in slide count, so tests scale
// the budget down against a proportionally smaller deck instead; production
// always runs the 30k default (null restores it).
let contextMaxChars = CONTEXT_MAX_CHARS

/** Test-only: scale the deck-overview budget (null restores the 30k default). */
export function _setDeckOverviewBudgetForTests(maxChars: number | null): void {
  contextMaxChars = maxChars ?? CONTEXT_MAX_CHARS
}
const PREVIEW_MAX_CHARS = 60
const PREVIEW_TIGHT_CHARS = 20
const READ_MAX_CHARS = 30_000
/** insert_content text payload cap (same as the markdown/html line sessions) */
export const INSERT_MAX_CHARS = 200_000
/** EMU per inch (OOXML canonical 914400); agent-facing geometry is in inches */
const EMU_PER_INCH = 914_400
/** default text box: 6 x 1 in at (1", 1") — inside both 16:9 and 4:3 canvases */
const DEFAULT_BOX_INCHES = { x: 1, y: 1, width: 6, height: 1 }

function clip(text: string, max: number): string {
  return text.length <= max ? text : `${text.slice(0, Math.max(1, max - 1))}…`
}

function inchesFromEmu(emu: number): string {
  return (Math.round((emu / EMU_PER_INCH) * 10) / 10).toFixed(1)
}

// ---- text extraction ----

function runsText(paragraphs: Paragraph[] | undefined): string {
  if (!paragraphs) return ''
  return paragraphs
    .map((p) =>
      p.runs
        .filter((r) => !r.paraMark)
        .map((r) => r.text)
        .join(''),
    )
    .join('\n')
}

/** all text on an element, recursing into groups and table cells */
function elementText(el: SlideElement): string {
  if (el.type === 'text' || el.type === 'shape') return runsText(el.text?.paragraphs)
  if (el.type === 'table')
    return (el as TableElement).rows
      .map((row) =>
        row
          .filter((cell) => !cell.merged)
          .map((cell) => runsText(cell.text?.paragraphs))
          .join(' | '),
      )
      .join('\n')
  if (el.type === 'group') return el.children.map(elementText).filter(Boolean).join('\n')
  return ''
}

function slideText(slide: Slide): string {
  return slide.elements.map(elementText).filter(Boolean).join('\n')
}

function slidePreview(slide: Slide, max: number): string {
  const firstLine =
    slideText(slide)
      .split('\n')
      .find((line) => line.trim() !== '') ?? ''
  // newlines would break the one-line-per-slide overview format
  return clip(firstLine.replace(/\s+/g, ' ').trim(), max)
}

function elementTypeName(el: SlideElement): string {
  switch (el.type) {
    case 'text':
    case 'shape': {
      const te = el as TextElement
      if (te.placeholder) return `${te.type}:${te.placeholder}`
      return te.type
    }
    case 'group':
      return `group(${String(el.children.length)})`
    case 'table': {
      const tbl = el as TableElement
      return `table(${String(tbl.rows.length)}x${String(tbl.colWidths.length)})`
    }
    default:
      return el.type
  }
}

// ---- session ----

export interface SlidesSessionMeta {
  handle: string
  kind: 'slides'
  path: string
  fileName: string
  format: 'pptx'
  converted: boolean
  editable: boolean
  warnings: string[]
  slideCount: number
  elementCount: number
  wordCount: number
  charCount: number
  /** slide canvas size, EMU plus inches for agent reasoning */
  size: { cxEmu: number; cyEmu: number; widthIn: string; heightIn: string }
  dirty: boolean
}

export interface SlideSaveResult {
  path: string
  bytes: number
  unchanged: boolean
  format: 'pptx'
}

export interface SlidesInsertOptions {
  /** 0-based slide index (required) */
  slide: number
  /** element index on the slide; given -> replace its text, absent -> add a text box */
  element?: number
  /** text box geometry in inches (ignored when element is given) */
  x?: number
  y?: number
  width?: number
  height?: number
}

interface FileStamp {
  mtimeMs: number
  size: number
}

export class SlidesSession {
  readonly handle: string
  readonly path: string

  private readonly root: string
  private readonly opened: OpenedPptx
  private readonly originalBytes: Uint8Array
  private baseline: FileStamp | null
  /** every target this session has written (repeat save-as needs no overwrite) */
  private readonly savedTargets = new Set<string>()
  private edited = false

  private constructor(
    handle: string,
    path: string,
    root: string,
    opened: OpenedPptx,
    originalBytes: Uint8Array,
    stamp: FileStamp | null,
  ) {
    this.handle = handle
    this.path = path
    this.root = root
    this.opened = opened
    this.originalBytes = originalBytes
    this.baseline = stamp
  }

  /** Open a .pptx inside the workspace root and parse it into the engine model. */
  static async open(rawPath: string, root?: string): Promise<SlidesSession> {
    const path = resolveConfined(rawPath, root)
    let bytes: Uint8Array
    let stamp: FileStamp
    try {
      bytes = new Uint8Array(await readFile(path))
      const info = await stat(path)
      stamp = { mtimeMs: info.mtimeMs, size: info.size }
    } catch (e) {
      throw new Error(`Cannot read "${path}": ${e instanceof Error ? e.message : String(e)}`, {
        cause: e,
      })
    }
    let opened: OpenedPptx
    try {
      opened = await openPptx(bytes)
    } catch (e) {
      throw new Error(
        `Cannot parse "${path}" as a .pptx presentation: ${e instanceof Error ? e.message : String(e)}`,
        { cause: e },
      )
    }
    return new SlidesSession(randomUUID(), path, root ?? workspaceRoot(), opened, bytes, stamp)
  }

  // ---- reading ----

  meta(): SlidesSessionMeta {
    const { deck } = this.opened
    const fullText = deck.slides.map((s) => slideText(s)).join('\n')
    return {
      handle: this.handle,
      kind: 'slides',
      path: this.path,
      fileName: basename(this.path) || this.path,
      format: 'pptx',
      converted: false,
      editable: true,
      warnings: [],
      slideCount: deck.slides.length,
      elementCount: deck.slides.reduce((sum, s) => sum + s.elements.length, 0),
      wordCount: countWords(fullText),
      charCount: fullText.length,
      size: {
        cxEmu: deck.size.cx,
        cyEmu: deck.size.cy,
        widthIn: inchesFromEmu(deck.size.cx),
        heightIn: inchesFromEmu(deck.size.cy),
      },
      dirty: this.edited,
    }
  }

  private requireSlide(index: number): Slide {
    const slides = this.opened.deck.slides
    if (!Number.isInteger(index) || index < 0 || index >= slides.length) {
      throw new Error(
        `slide index ${String(index)} is out of range (deck has ${String(slides.length)} slides, ` +
          'indexes are 0-based) — re-read the deck after edits',
      )
    }
    return slides[index]!
  }

  /**
   * Deck overview for the agent: `index|elements|preview` per slide plus
   * stats, under the shared ~30k budget (previews tighten and the middle
   * elides first, numbering stays verifiable).
   */
  readDeck(): string {
    const slides = this.opened.deck.slides
    const meta = this.meta()
    const header =
      `The deck has ${String(slides.length)} slide(s) on a ${meta.size.widthIn} x ` +
      `${meta.size.heightIn} in canvas (index|elements|content preview):`
    const render = (bodyMax: number) =>
      slides.map(
        (slide, index) =>
          `${String(index)}|${String(slide.elements.length)}|${slidePreview(slide, bodyMax)}`,
      )
    const stats =
      `Deck stats: ${String(meta.elementCount)} elements, ${String(meta.wordCount)} words, ` +
      `${String(meta.charCount)} characters of slide text.`
    let body = render(PREVIEW_MAX_CHARS)
    let out = [header, ...body, stats].join('\n')
    if (out.length > contextMaxChars) {
      body = render(PREVIEW_TIGHT_CHARS)
      out = [header, ...body, stats].join('\n')
      if (out.length > contextMaxChars) {
        const dropStart = Math.floor(body.length / 3)
        const dropEnd = body.length - Math.floor(body.length / 3)
        out = [
          header,
          ...body.slice(0, dropStart),
          `…(${String(dropEnd - dropStart)} slides elided here; numbering is continuous)…`,
          ...body.slice(dropEnd),
          stats,
        ].join('\n')
      }
    }
    return out
  }

  /**
   * One slide in detail: the element list (`index|type|name|text preview` —
   * element indexes are the addressing scheme for insert_content) and the
   * slide's full text. Output stays under ~30k characters.
   */
  readSlide(slideIndex: number): string {
    const slide = this.requireSlide(slideIndex)
    const lines = slide.elements.map((el, index) => {
      const name = el.name ? clip(el.name, PREVIEW_MAX_CHARS) : ''
      const preview = clip(elementText(el).replace(/\s+/g, ' ').trim(), PREVIEW_MAX_CHARS)
      return [String(index), elementTypeName(el), name, preview].filter(Boolean).join('|')
    })
    const header =
      `Slide ${String(slideIndex)} of ${String(this.opened.deck.slides.length)} — ` +
      `${String(slide.elements.length)} element(s) (index|type|name|text preview; ` +
      'element indexes address insert_content):'
    const fullText = slideText(slide)
    const textBlock = fullText.trim() === '' ? '(no text on this slide)' : fullText
    let out = [header, ...lines, '', 'Full slide text (EOLs normalized to LF):', textBlock].join(
      '\n',
    )
    if (out.length > READ_MAX_CHARS) {
      out =
        out.slice(0, READ_MAX_CHARS) +
        `\n…(output truncated at ${String(READ_MAX_CHARS)} characters)`
    }
    return out
  }

  // ---- editing ----

  /**
   * Insert plain text into the deck (in memory; persist with save). With
   * `element`: replace that element's text body (text boxes and autoshapes;
   * a first text on a bare autoshape gets PowerPoint's centered authoring
   * defaults, connectors refuse — they cannot hold text). Without: add a new
   * text box at the given geometry (inches; default 6 x 1 in at 1", 1").
   */
  insertContent(
    text: string,
    options: SlidesInsertOptions,
  ): { slide: number; element: number; added: boolean; paragraphs: number; detail: string } {
    if (typeof text !== 'string' || text.length === 0) {
      throw new Error('text is required (non-empty plain text; line breaks become paragraphs)')
    }
    if (text.length > INSERT_MAX_CHARS) {
      throw new Error(
        `text is ${String(text.length)} characters; the cap is ${String(INSERT_MAX_CHARS)}`,
      )
    }
    const slide = this.requireSlide(options.slide)
    const paragraphs: Paragraph[] = text.split(/\r\n|\r|\n/).map((line) => ({
      runs: [{ text: line }],
    }))
    if (options.element !== undefined) {
      const el = slide.elements[options.element]
      if (!el || !Number.isInteger(options.element) || options.element < 0) {
        throw new Error(
          `element index ${String(options.element)} is out of range (slide has ` +
            `${String(slide.elements.length)} elements, indexes are 0-based) — re-read the slide`,
        )
      }
      if (el.type !== 'text' && el.type !== 'shape') {
        throw new Error(
          `element ${String(options.element)} is a "${el.type}" element; only text boxes and ` +
            'shapes hold editable text (pictures, tables, charts and groups do not — see read_deck)',
        )
      }
      const te = el as TextElement
      if (isConnectorXml(te.anchor.originalXml)) {
        throw new Error(
          `element ${String(options.element)} is a connector, which cannot hold text.`,
        )
      }
      if (!te.text) {
        // mirror the app's first-text defaults: an autoshape centers, a text
        // box or placeholder stays top-left and inherits anchor from the layout
        const autoShape = te.type === 'shape' && !te.placeholder && !te.txBox
        te.text = {
          paragraphs: [],
          insets: { ...DEFAULT_BODY_INSETS },
          ...(autoShape ? { anchor: 'middle' as const } : {}),
        }
        if (autoShape) for (const p of paragraphs) p.align = 'center'
      }
      te.text.paragraphs = paragraphs
      te.dirty = true
      this.edited = true
      return {
        slide: options.slide,
        element: options.element,
        added: false,
        paragraphs: paragraphs.length,
        detail: `replaced the text of element ${String(options.element)} on slide ${String(options.slide)}`,
      }
    }
    const box = {
      x: options.x ?? DEFAULT_BOX_INCHES.x,
      y: options.y ?? DEFAULT_BOX_INCHES.y,
      width: options.width ?? DEFAULT_BOX_INCHES.width,
      height: options.height ?? DEFAULT_BOX_INCHES.height,
    }
    addElement(slide, {
      kind: 'textbox',
      offset: {
        x: Math.round(box.x * EMU_PER_INCH),
        y: Math.round(box.y * EMU_PER_INCH),
        cx: Math.round(box.width * EMU_PER_INCH),
        cy: Math.round(box.height * EMU_PER_INCH),
      },
      paragraphs,
    })
    this.edited = true
    const index = slide.elements.length - 1
    return {
      slide: options.slide,
      element: index,
      added: true,
      paragraphs: paragraphs.length,
      detail:
        `added a text box (${box.width} x ${box.height} in at ${box.x}", ${box.y}") as element ` +
        `${String(index)} on slide ${String(options.slide)}`,
    }
  }

  // ---- saving ----

  /**
   * Save atomically with the docx session's fences: saving over the opened
   * file refuses when it changed on disk since open; a target that already
   * exists is refused unless the session owns it or overwrite is true. With
   * no edits the original bytes round-trip verbatim (an untouched file never
   * changes on disk); an edited save streams through the engine's atomic
   * savePptxToFile (untouched zip entries byte-identical, dirty slides
   * rebuilt) and then syncs the model with commitSaved.
   */
  async save(rawPath?: string, options: { overwrite?: boolean } = {}): Promise<SlideSaveResult> {
    // a pinned root that vanished (moved/renamed workspace directory) must
    // fail here, before confinement + mkdir silently resurrects it (BUG-1103)
    await assertWorkspaceRootExists(this.root)
    const target = resolveConfined(rawPath ?? this.path, this.root)
    await assertSaveTargetFree(target, [this.path, ...this.savedTargets], options.overwrite)
    if (target === this.path && this.baseline) {
      let current: FileStamp
      try {
        const info = await stat(this.path)
        current = { mtimeMs: info.mtimeMs, size: info.size }
      } catch {
        current = { mtimeMs: -1, size: -1 }
      }
      if (current.mtimeMs !== this.baseline.mtimeMs || current.size !== this.baseline.size) {
        throw new FencingError(this.path)
      }
    }

    await mkdir(dirname(target), { recursive: true })
    // basename, not a '/'-split: on Windows the split leaves the whole path
    // in the temp name and writeFile fails on the colons/backslashes
    const tmp = join(dirname(target), `.${basename(target) || 'deck'}.airy-${randomUUID()}`)
    let unchanged = false
    if (!this.edited) {
      // zero-edit save: the original bytes round-trip verbatim (the engine's
      // regenerated zip container would re-compress identical entries)
      await writeFile(tmp, this.originalBytes)
      unchanged = true
    } else {
      // savePptxToFile lands the deck at tmp atomically (its own temp +
      // rename); the promote below carries the docx session's ownership rules
      await savePptxToFile(this.opened, tmp)
      commitSaved(this.opened)
    }
    // a fresh (guarded) target promotes exclusively: a file created between
    // the guard's stat and this write surfaces the clobber error instead of
    // being silently replaced; targets this session owns replace by intent
    if (options.overwrite === true || target === this.path || this.savedTargets.has(target)) {
      await rename(tmp, target)
    } else {
      await promoteNewFileExclusively(tmp, target)
    }

    // refresh the fence so chained saves keep working
    if (target === this.path) {
      try {
        const info = await stat(this.path)
        this.baseline = { mtimeMs: info.mtimeMs, size: info.size }
      } catch {
        this.baseline = null
      }
    }
    this.savedTargets.add(target)
    const info = await stat(target)
    return { path: target, bytes: info.size, unchanged, format: 'pptx' }
  }

  /** Release session resources: nothing to clean (no temp files, no sidecar). */
  close(): Promise<string[]> {
    return Promise.resolve([])
  }
}
