/**
 * BUG-1671 regression: the zip written by savePptxToFile must be accepted by strict
 * loaders (LibreOffice Impress rejected the whole package when the writer streamed
 * per-entry and emitted data descriptors). Pins the container contract:
 * no general-purpose bit 3, local headers carry real crc/sizes and agree with the
 * central directory, no trailing data-descriptor signatures — and the streamed file
 * stays byte-identical to the in-memory savePptx output.
 */
import { describe, it, expect } from 'vitest'
import { mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import JSZip from 'jszip'
import { openPptx, savePptx, savePptxToFile, addElement, addPicture } from '../src/index'
import { auditZipStructure } from './helpers/zip-structure'
import { noisePng } from './helpers/test-media'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))
const out = () => join(mkdtempSync(join(tmpdir(), 'save-zip-compat-')), 'out.pptx')

function expectStrictZip(bytes: Buffer): void {
  const audit = auditZipStructure(bytes)
  expect(audit.descriptorFlagged, 'entries flagged with data-descriptor bit 3').toEqual([])
  expect(audit.headerProblems, 'local headers disagreeing with the central directory').toEqual([])
  expect(audit.trailingDescriptors, 'entries followed by a data descriptor').toEqual([])
}

describe('savePptxToFile zip container (BUG-1671)', () => {
  it('writes every entry without data descriptors, local headers matching the central directory', async () => {
    for (const fixture of ['01_standard_business.pptx', '05_unicode_cjk_emoji.pptx']) {
      const opened = await openPptx(fx(fixture))
      const target = out()
      await savePptxToFile(opened, target)
      const saved = readFileSync(target)
      expect(saved.length).toBeGreaterThan(0)
      expectStrictZip(saved)
    }
  })

  it('is byte-identical to the in-memory savePptx output', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const target = out()
    await savePptxToFile(opened, target)
    const fromFile = readFileSync(target)
    const fromMemory = Buffer.from(await savePptx(await openPptx(fx('01_standard_business.pptx'))))
    expect(Buffer.compare(fromFile, fromMemory)).toBe(0)
  })

  it('keeps the container strict after a content edit (the Ctrl+S path)', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    addElement(opened.deck.slides[0]!, {
      kind: 'textbox',
      offset: { x: 914400, y: 914400, cx: 6096000, cy: 914400 },
      paragraphs: [{ runs: [{ text: 'edited before save' }] }],
    })
    const target = out()
    await savePptxToFile(opened, target)
    expectStrictZip(readFileSync(target))
    const reopened = await openPptx(readFileSync(target))
    expect(reopened.archive.readText(reopened.deck.slides[0]!.path)!).toContain(
      'edited before save',
    )
  })

  it('keeps the container strict with multi-megabyte stored media parts', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const added: Array<{ path: string; bytes: Buffer }> = []
    for (let i = 0; i < 3; i++) {
      const png = noisePng(1600, 400, 0x9e3779b9 + i)
      const el = addPicture(opened, slide, {
        bytes: png,
        ext: 'png',
        offset: { x: 914400 + i * 914400, y: 3200400, cx: 1828800, cy: 457200 },
        name: `bulk ${i}`,
      })
      expect(el, `picture ${i} inserted`).not.toBeNull()
      added.push({ path: el!.mediaRef, bytes: png })
    }
    const target = out()
    await savePptxToFile(opened, target)
    const saved = readFileSync(target)
    expectStrictZip(saved)

    // media rides along verbatim (STORE) and survives the streamed write
    const zip = await JSZip.loadAsync(saved)
    for (const { path, bytes } of added) {
      const entry = zip.file(path)
      expect(entry, path).not.toBeNull()
      expect(Buffer.compare(await entry!.async('nodebuffer'), bytes)).toBe(0)
    }
  })
})
