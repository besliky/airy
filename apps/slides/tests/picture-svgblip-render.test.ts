/**
 * BUG-1656: pptx pictures carrying an Office 2016 <asvg:svgBlip> extension must
 * render from the vector part (correct size comes from the frame's xfrm), with
 * the co-embedded raster kept only as a decode fallback — instead of silently
 * stretching the (often 1x1) fallback PNG over the whole frame.
 */
import { describe, it, expect, vi } from 'vitest'
import {
  createBlankPptx,
  openPptx,
  savePptx,
  solidPng,
  type OpenedPptx,
  type PictureElement,
} from '@airy-office/pptx-engine'
import { buildRenderSlide } from '@airy-office/pptx-render'
import { makeMediaResolver } from '../src/main/session-state'
import { pictureDisplayImage } from '../src/renderer/konva-adapter'

vi.mock('electron', () => ({
  BrowserWindow: { getFocusedWindow: () => null },
  webContents: { getAllWebContents: () => [] },
}))
vi.mock('../src/main/fonts', () => ({
  createSystemFontMetrics: () => ({}),
}))

const GOOD_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20">' +
  '<circle cx="10" cy="10" r="9" fill="#3a7"/></svg>'
// Well-formed-looking but undecodable: Chromium's img.onerror fires for this
const BROKEN_SVG = '<svg xmlns="http://www.w3.org/2000/svg"><circ'

const SVG_EXT =
  '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
  '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId3"/>' +
  '</a:ext></a:extLst>'

const slideXml = (blipInner: string) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>' +
  '<p:pic><p:nvPicPr><p:cNvPr id="7" name="Vector"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
  `<p:blipFill><a:blip r:embed="rId2">${blipInner}</a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="2743200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>' +
  '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'

const relsXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>' +
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.svg"/>' +
  '</Relationships>'

async function deckWith(blipInner: string, svgText: string): Promise<OpenedPptx> {
  const opened = await openPptx(await createBlankPptx())
  opened.archive.entries.set('ppt/media/image1.png', new Uint8Array(solidPng(1, 1, [255, 0, 0])))
  opened.archive.entries.set('ppt/media/image2.svg', Buffer.from(svgText, 'utf8'))
  opened.archive.entries.set('ppt/slides/_rels/slide1.xml.rels', Buffer.from(relsXml, 'utf8'))
  opened.archive.entries.set('ppt/slides/slide1.xml', Buffer.from(slideXml(blipInner), 'utf8'))
  return openPptx(await savePptx(opened))
}

const picOf = (opened: OpenedPptx) => opened.deck.slides[0]!.elements[0] as PictureElement
const renderPic = (opened: OpenedPptx) => {
  const slide = opened.deck.slides[0]!
  const rs = buildRenderSlide(slide, opened.deck.size, {
    fitWidthPx: 1280,
    media: makeMediaResolver(opened, slide.path),
  })
  return rs.nodes.find((n) => n.type === 'picture')
}

describe('svgBlip picture renders the vector part with a raster fallback (BUG-1656)', () => {
  it('full chain: model mediaRef points at the SVG part, raster kept as fallback', async () => {
    const opened = await deckWith(SVG_EXT, GOOD_SVG)
    const el = picOf(opened)
    expect(el.mediaRef).toBe('ppt/media/image2.svg')
    expect(el.fallbackMediaRef).toBe('ppt/media/image1.png')
  })

  it('render node carries an image/svg+xml dataUrl plus a real PNG fallbackDataUrl', async () => {
    const opened = await deckWith(SVG_EXT, GOOD_SVG)
    const node = renderPic(opened)
    expect(node).toBeTruthy()
    expect(node!.dataUrl?.startsWith('data:image/svg+xml')).toBe(true)
    expect(node!.fallbackDataUrl?.startsWith('data:image/png;base64,')).toBe(true)
  })

  it('PNG-only picture keeps the legacy behavior: raster dataUrl, no fallback slot', async () => {
    const opened = await deckWith('', GOOD_SVG)
    const el = picOf(opened)
    expect(el.mediaRef).toBe('ppt/media/image1.png')
    expect(el.fallbackMediaRef).toBeUndefined()
    const node = renderPic(opened)
    expect(node!.dataUrl?.startsWith('data:image/png;base64,')).toBe(true)
    expect(node!.fallbackDataUrl).toBeUndefined()
  })

  it('broken SVG bytes: the SVG still drives the render while the raster stays as fallback', async () => {
    const opened = await deckWith(SVG_EXT, BROKEN_SVG)
    const node = renderPic(opened)
    // The resolver serves by sniffed/extension mime; decode failure is only
    // observable in the renderer, so the fallback must be threaded regardless
    expect(node!.dataUrl?.startsWith('data:image/svg+xml')).toBe(true)
    expect(node!.fallbackDataUrl?.startsWith('data:image/png;base64,')).toBe(true)
  })
})

describe('pictureDisplayImage substitutes the raster when the SVG fails to decode', () => {
  const svgImg = { width: 20, height: 20 } as HTMLImageElement
  const pngImg = { width: 1, height: 1 } as HTMLImageElement
  const pic = {
    dataUrl: 'data:image/svg+xml;base64,PHN2Zy8+',
    fallbackDataUrl: 'data:image/png;base64,AAAA',
  }

  it('prefers the decoded SVG', () => {
    const images = new Map([
      [pic.dataUrl, svgImg],
      [pic.fallbackDataUrl, pngImg],
    ])
    expect(pictureDisplayImage(pic, images)).toBe(svgImg)
  })

  it('falls back to the raster when only the fallback decoded (SVG onerror)', () => {
    const images = new Map([[pic.fallbackDataUrl, pngImg]])
    expect(pictureDisplayImage(pic, images)).toBe(pngImg)
  })

  it('returns undefined when neither decoded', () => {
    expect(pictureDisplayImage(pic, new Map())).toBeUndefined()
  })
})
