import { existsSync } from 'node:fs'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join } from 'node:path'
import { PNG } from 'pngjs'
import { afterEach, describe, expect, it } from 'vitest'

import { buildPdfExportHtml, exportSlidesPdf, type PdfExportWindow } from '../src/main/pdf-export'
import { renderSlideSvg } from '../src/renderer/slide-svg'
import type { RenderNode, RenderSlide, ShapeRenderNode } from '@airy-office/pptx-render'

const roots: string[] = []

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

async function outputPath(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), 'airy-slides-pdf-test-'))
  roots.push(root)
  return join(root, 'export.pdf')
}

function singlePixelPngBase64(): string {
  const png = new PNG({ width: 1, height: 1 })
  png.data.set([0x12, 0x34, 0x56, 0xff])
  return PNG.sync.write(png).toString('base64')
}

function deterministicNoisePngBase64(): string {
  const png = new PNG({ width: 800, height: 800 })
  let state = 0x12345678
  const nextByte = (): number => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    return state >>> 24
  }
  for (let i = 0; i < png.data.length; i += 4) {
    png.data[i] = nextByte()
    png.data[i + 1] = nextByte()
    png.data[i + 2] = nextByte()
    png.data[i + 3] = 255
  }
  return PNG.sync.write(png).toString('base64')
}

/** minimal vector slide with real text, like the renderer's renderSlideSvg output */
const vectorSlide = (label: string): string =>
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1600 900" width="100%" height="100%">` +
  `<rect width="1600" height="900" fill="#123456"/>` +
  `<text x="80" y="120" font-family="Carlito" font-size="44.00" fill="#ffffff">${label}</text></svg>`

/** a real gradient slide through the vector painter (defs + url(#id) reference) */
function gradientSlideSvg(angleDeg: number, prefix: string): string {
  const node: RenderNode = {
    id: 'g',
    type: 'shape',
    sourceId: 'g',
    box: { x: 0, y: 0, w: 800, h: 450, rotationDeg: 0, flipH: false, flipV: false },
    presetGeometry: 'rect',
    fill: {
      kind: 'gradient',
      stops: [
        { pos: 0, color: 'FF0000' },
        { pos: 1, color: '0000FF' },
      ],
      angleDeg,
    },
  } as unknown as ShapeRenderNode
  const slide = {
    widthPx: 1600,
    heightPx: 900,
    scale: 1,
    background: { kind: 'solid', color: 'FFFFFF' },
    nodes: [node],
  } as unknown as RenderSlide
  return renderSlideSvg(slide, new Map(), prefix)
}

class TestPdfWindow implements PdfExportWindow {
  loadedPath: string | null = null
  loadedHtml = ''
  destroyed = false
  tempDirectoryExistedAtDestroy: boolean | null = null
  shouldFailLoad = false
  shouldFailPrint = false
  printOptions: Electron.PrintToPDFOptions | null = null

  async loadFile(path: string): Promise<void> {
    this.loadedPath = path
    this.loadedHtml = await readFile(path, 'utf8')
    if (this.shouldFailLoad) throw new Error('load failed')
  }

  webContents = {
    printToPDF: async (options: Electron.PrintToPDFOptions): Promise<Buffer> => {
      this.printOptions = options
      if (this.shouldFailPrint) throw new Error('print failed')
      return Buffer.from('PDF')
    },
  }

  destroy(): void {
    if (this.loadedPath) this.tempDirectoryExistedAtDestroy = existsSync(dirname(this.loadedPath))
    this.destroyed = true
  }
}

describe('slides PDF export', () => {
  it('builds vector pages with selectable text and raster fallback pages side by side', () => {
    const html = buildPdfExportHtml(
      [
        { svg: vectorSlide('Quarterly <Review>') },
        { pngBase64: singlePixelPngBase64() },
        { svg: vectorSlide('Second page') },
      ],
      13.333,
      7.5,
    )
    // vector pages keep real text (becomes selectable PDF text through printToPDF)
    expect(html).toContain('<svg xmlns="http://www.w3.org/2000/svg"')
    expect(html).toContain('<text x="80" y="120"')
    expect(html).toContain('Quarterly <Review>')
    expect(html).toContain('Second page')
    // raster fallback still embeds the bitmap
    expect(html).toContain(`data:image/png;base64,${singlePixelPngBase64()}`)
    expect((html.match(/<div class="page">/g) ?? []).length).toBe(3)
    expect(html).toContain('@page { size: 13.333in 7.5in; margin: 0; }')
    expect(html).toContain('.page svg { display: block; width: 100%; height: 100%; }')
  })

  it('joins vector pages without defs id collisions (each url(#) finds its own defs)', () => {
    // url(#id) is document-global: two inline SVGs that both say id="grad0"
    // would cross-reference (slide 2 painted with slide 1's gradient)
    const html = buildPdfExportHtml(
      [
        { svg: gradientSlideSvg(0, 'p0-') },
        { svg: gradientSlideSvg(90, 'p1-') },
        { svg: vectorSlide('text page') },
      ],
      13.333,
      7.5,
    )
    // every id in the joined document is unique
    const ids = [...html.matchAll(/ id="([^"]+)"/g)].map((m) => m[1])
    expect(ids.length).toBeGreaterThanOrEqual(2)
    expect(new Set(ids).size).toBe(ids.length)
    // each page's reference points at the defs emitted in the same page
    const pages = html.split('<div class="page">').slice(1)
    expect(pages).toHaveLength(3)
    expect(pages[0]).toContain('<linearGradient id="p0-grad0"')
    expect(pages[0]).toContain('fill="url(#p0-grad0)"')
    expect(pages[1]).toContain('<linearGradient id="p1-grad0"')
    expect(pages[1]).toContain('fill="url(#p1-grad0)"')
  })

  it('loads a temporary HTML file containing all pages, writes the PDF, then removes the directory', async () => {
    const win = new TestPdfWindow()
    const filePath = await outputPath()
    const opened: string[] = []
    const chromiumDataUrlLimit = 2 * 1024 * 1024
    const firstPng = singlePixelPngBase64()
    const secondPng = deterministicNoisePngBase64()
    const decodedSecondPng = PNG.sync.read(Buffer.from(secondPng, 'base64'))

    expect(decodedSecondPng.width).toBeGreaterThan(0)
    expect(decodedSecondPng.height).toBeGreaterThan(0)

    const result = await exportSlidesPdf({
      pages: [
        { svg: vectorSlide('Vector page') },
        { pngBase64: firstPng },
        { pngBase64: secondPng },
      ],
      widthPx: 1600,
      heightPx: 900,
      filePath,
      createWindow: () => win,
      openExportedPdf: (path) => opened.push(path),
    })

    expect(result).toEqual({ ok: true, path: filePath })
    expect(basename(win.loadedPath!)).toBe('slides.html')
    expect(basename(dirname(win.loadedPath!))).toMatch(/^airy-slides-pdf-/)
    expect(win.loadedHtml).toContain('<text x="80" y="120"')
    expect(win.loadedHtml).toContain('Vector page')
    expect(win.loadedHtml).toContain(`data:image/png;base64,${firstPng}`)
    expect(win.loadedHtml.length).toBeGreaterThan(chromiumDataUrlLimit)
    expect(win.loadedHtml).toContain('@page { size: 13.333in 7.5in; margin: 0; }')
    // the hidden window is scripting-disabled; readiness rides on loadFile's onload
    expect(win.printOptions).toEqual({
      landscape: false,
      printBackground: true,
      pageSize: { width: 13.333, height: 7.5 },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      preferCSSPageSize: false,
    })
    await expect(readFile(filePath)).resolves.toEqual(Buffer.from('PDF'))
    expect(opened).toEqual([filePath])
    expect(existsSync(dirname(win.loadedPath!))).toBe(false)
    expect(win.destroyed).toBe(true)
    expect(win.tempDirectoryExistedAtDestroy).toBe(true)
  })

  it('removes the temporary directory and destroys the window when loading fails', async () => {
    const win = new TestPdfWindow()
    win.shouldFailLoad = true

    const result = await exportSlidesPdf({
      pages: [{ svg: vectorSlide('x') }],
      widthPx: 4,
      heightPx: 3,
      filePath: await outputPath(),
      createWindow: () => win,
      openExportedPdf: () => {},
    })

    expect(result).toEqual({ ok: false, error: 'Error: load failed' })
    expect(win.loadedPath).not.toBeNull()
    expect(existsSync(dirname(win.loadedPath!))).toBe(false)
    expect(win.destroyed).toBe(true)
  })

  it('removes the temporary directory and destroys the window when PDF printing fails', async () => {
    const win = new TestPdfWindow()
    win.shouldFailPrint = true

    const result = await exportSlidesPdf({
      pages: [{ svg: vectorSlide('x') }],
      widthPx: 4,
      heightPx: 3,
      filePath: await outputPath(),
      createWindow: () => win,
      openExportedPdf: () => {},
    })

    expect(result).toEqual({ ok: false, error: 'Error: print failed' })
    expect(win.loadedPath).not.toBeNull()
    expect(existsSync(dirname(win.loadedPath!))).toBe(false)
    expect(win.destroyed).toBe(true)
  })

  it('handout2/handout3 layouts pack ceil(slides/N) A4 pages with the vector slides inline', async () => {
    const win = new TestPdfWindow()
    const pages = ['a', 'b', 'c', 'd', 'e'].map((label) => ({ svg: vectorSlide(label) }))
    const result = await exportSlidesPdf({
      pages,
      widthPx: 1600,
      heightPx: 900,
      filePath: await outputPath(),
      layout: 'handout3',
      createWindow: () => win,
      openExportedPdf: () => {},
    })
    expect(result.ok).toBe(true)
    // 5 slides at 3 per sheet = 2 pages; the assembly is the print sheet's
    expect(win.loadedHtml).toContain('@page { size: 8.27in 11.69in; margin: 0; }')
    expect(win.loadedHtml.match(/class="page handout h3"/g)).toHaveLength(2)
    // the vector slides stay inline (selectable text) in their cells
    expect(win.loadedHtml).toContain('<div class="cell"><svg xmlns=')
    expect(win.loadedHtml).toContain('>e</text>')
    expect(win.printOptions).toEqual({
      landscape: false,
      printBackground: true,
      pageSize: { width: 8.27, height: 11.69 },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      preferCSSPageSize: false,
    })

    const win2 = new TestPdfWindow()
    await exportSlidesPdf({
      pages: pages.slice(0, 4),
      widthPx: 1600,
      heightPx: 900,
      filePath: await outputPath(),
      layout: 'handout2',
      createWindow: () => win2,
      openExportedPdf: () => {},
    })
    expect(win2.loadedHtml.match(/class="page handout h2"/g)).toHaveLength(2)
  })

  it('notes layout pairs each vector slide with its notes text on A4', async () => {
    const win = new TestPdfWindow()
    const result = await exportSlidesPdf({
      pages: [{ svg: vectorSlide('One') }, { pngBase64: singlePixelPngBase64() }],
      widthPx: 1600,
      heightPx: 900,
      filePath: await outputPath(),
      layout: 'notes',
      notes: ['Remember the <agenda>', ''],
      createWindow: () => win,
      openExportedPdf: () => {},
    })
    expect(result.ok).toBe(true)
    expect(win.loadedHtml.match(/class="page notes"/g)).toHaveLength(2)
    // notes text is present and escaped, slide text stays selectable
    expect(win.loadedHtml).toContain('Remember the &lt;agenda&gt;')
    expect(win.loadedHtml).toContain('>One</text>')
    // the raster fallback page keeps its bitmap slot
    expect(win.loadedHtml).toContain(`data:image/png;base64,${singlePixelPngBase64()}`)
    expect(win.loadedHtml).toContain('@page { size: 8.27in 11.69in; margin: 0; }')
  })
})
