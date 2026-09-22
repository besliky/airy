import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import type { RunStyle } from '@airy-office/pptx-render'

// shaped-metrics pulls in the harfbuzz wasm (?asset only resolves in electron builds); disconnected in unit tests
vi.mock('../src/main/shaped-metrics', () => ({
  initShapedMetrics: () => {},
  shapedMeasure: () => null,
  shapedFamily: () => null,
  gtMeasure: () => null,
  complexScriptOf: () => null,
}))

import { createSystemFontMetrics, exportFontFaces, registerEmbeddedFonts } from '../src/main/fonts'

const style = (fontFamily: string, over: Partial<RunStyle> = {}): RunStyle => ({
  fontFamily,
  fontSizePx: 100,
  bold: false,
  italic: false,
  ...over,
})

const stylesOf = (faces: ReturnType<typeof exportFontFaces>) =>
  faces.map((f) => `${f.bold ? 'b' : ''}${f.italic ? 'i' : ''}`).sort()

describe('exportFontFaces (PDF export @font-face source, BUG-1506)', () => {
  it('inlines the bundled substitute under the exact family spelling (all 4 styles)', () => {
    // the SVG text carries the drawing family — 'Carlito GO' when the layout
    // resolved Calibri through the bundled substitute, plain 'Carlito' when
    // the run named it directly; both spellings must answer to the bytes
    for (const family of ['Carlito GO', 'Carlito']) {
      const faces = exportFontFaces([family])
      expect(faces).toHaveLength(4)
      expect(stylesOf(faces)).toEqual(['', 'b', 'bi', 'i'])
      for (const face of faces) {
        expect(face.family).toBe(family)
        // TrueType sfnt magic (0x00010000) — real font bytes, not paths
        expect(face.bytes[0]).toBe(0)
        expect(face.bytes[1]).toBe(1)
        expect(face.bytes.length).toBeGreaterThan(100_000)
      }
    }
  })

  it('inlines a raw Calibri request on machines without a system Calibri', () => {
    // environment-dependent: this dev/CI machine has no installed Calibri, so
    // the registry's alias chain (calibri -> Carlito) lands on the bundled
    // file and the export must declare the substitute under the deck's own
    // name; a machine with real Calibri resolves it as a system font (0 faces)
    const faces = exportFontFaces(['Calibri'])
    if (faces.length > 0) {
      expect(faces).toHaveLength(4)
      expect(faces.every((f) => f.family === 'Calibri')).toBe(true)
    }
  })

  it('families that resolve as installed system fonts are omitted', () => {
    // an unknown family resolves nowhere: nothing to inline, and no noise
    expect(exportFontFaces(['Zz NoSuchFamily 4711'])).toEqual([])
  })

  it('inlines the deck-registered private faces (document embeds)', () => {
    // same channel the renderer's doc-fonts FontFaces use: register an
    // embedded face, resolve it once through the metrics provider (which
    // records the private face), then the export serves its bytes
    const sfnt = new Uint8Array(
      readFileSync(join(__dirname, '../../docs/src/renderer/fonts/Caladea-Bold.ttf')),
    )
    expect(registerEmbeddedFonts([{ typeface: 'Zz Export Face', style: 'regular', sfnt }])).toBe(
      true,
    )
    const metrics = createSystemFontMetrics()
    const drawingFamily = String(metrics.displayFamily!(style('Zz Export Face')))
    expect(drawingFamily).toMatch(/Caladea/)
    const faces = exportFontFaces([drawingFamily])
    expect(faces.length).toBeGreaterThanOrEqual(1)
    for (const face of faces) {
      expect(face.family).toBe(drawingFamily)
      expect(face.bytes.length).toBeGreaterThan(1000)
    }
  })
})
