import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
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

import {
  createSystemFontMetrics,
  exportFontFaces,
  registerEmbeddedFonts,
  resetFontRegistry,
  setUserFontDir,
} from '../src/main/fonts'

const style = (fontFamily: string, over: Partial<RunStyle> = {}): RunStyle => ({
  fontFamily,
  fontSizePx: 100,
  bold: false,
  italic: false,
  ...over,
})

const stylesOf = (faces: ReturnType<typeof exportFontFaces>) =>
  faces.map((f) => `${f.bold ? 'b' : ''}${f.italic ? 'i' : ''}`).sort()

// neutral user-font store (empty): tests that swap the store restore this, so
// the shared registry state stays equivalent to "no user fonts"
const EMPTY_USER_STORE = mkdtempSync(join(tmpdir(), 'export-font-faces-store-'))
setUserFontDir(EMPTY_USER_STORE)

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

  it('inlines a raw Calibri request unconditionally (BUG-1621)', () => {
    // the deck names Calibri; the export window cannot resolve it by name on
    // typical Linux, so the bundled substitute must ride along under the
    // deck's own spelling no matter what the system font index says — a
    // system-side resolve hit must never suppress the inline again
    const faces = exportFontFaces(['Calibri'])
    expect(faces).toHaveLength(4)
    expect(faces.every((f) => f.family === 'Calibri')).toBe(true)
    expect(stylesOf(faces)).toEqual(['', 'b', 'bi', 'i'])
  })

  it('inlines the alias chain onto the bundle (Calibri Light)', () => {
    const faces = exportFontFaces(['Calibri Light'])
    expect(faces).toHaveLength(4)
    expect(faces.every((f) => f.family === 'Calibri Light')).toBe(true)
  })

  it('inlines the same bundled bytes under every accepted spelling', () => {
    const byName = (family: string) =>
      exportFontFaces([family])
        .map((f) => Buffer.from(f.bytes).toString('base64'))
        .sort()
    expect(byName('Calibri')).toEqual(byName('Carlito GO'))
  })

  it('families outside the bundle stay omitted', () => {
    // an unknown family resolves nowhere: nothing to inline, and no noise
    expect(exportFontFaces(['Zz NoSuchFamily 4711'])).toEqual([])
    // a system-resolvable alias target (Arial -> Liberation Sans et al.) is
    // still resolved by the export window by name — never inlined from the
    // bundle (the deck's non-bundle fonts must not regress)
    expect(exportFontFaces(['Arial'])).toEqual([])
  })

  it('inlines the bundle even when a system Carlito wins the registry lookup (BUG-1621)', () => {
    // Simulates the audit machine: a system-like Carlito file that beats the
    // bundled copy in the font index (user store files key 'carlito', which the
    // bundled 'Carlito-Regular' filenames never occupy). The pre-fix code
    // consulted the system resolve here, saw a non-bundled path and skipped
    // the inline — the export window then fell back to the default sans.
    const userDir = mkdtempSync(join(tmpdir(), 'bug1621-userfonts-'))
    // regular face of the same bundled files, under a plain 'Carlito' name
    writeFileSync(join(userDir, 'Carlito.ttf'), exportFontFaces(['Carlito'])[0]!.bytes)
    setUserFontDir(userDir)
    resetFontRegistry()
    try {
      const faces = exportFontFaces(['Calibri'])
      expect(faces).toHaveLength(4)
      expect(faces.every((f) => f.family === 'Calibri')).toBe(true)
      // and the direct Carlito spelling keeps answering with the bundle too
      expect(exportFontFaces(['Carlito'])).toHaveLength(4)
    } finally {
      rmSync(userDir, { recursive: true, force: true })
      setUserFontDir(EMPTY_USER_STORE)
      resetFontRegistry()
    }
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
