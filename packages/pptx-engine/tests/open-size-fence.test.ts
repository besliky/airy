import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { createBlankPptx, openPptx } from '../src/index'

/**
 * Open-time size fences (SEC-1304): the engine is the single open gate for
 * the desktop app, the merge/insert paths, and the headless MCP session.
 * The desktop open path used to run unfenced — a hostile .pptx opened in the
 * app reproduced exactly the bomb class SEC-1102 closed for headless (a few
 * KiB declaring gigabytes, materialized entry by entry). The budgets mirror
 * the docx-engine fence (zip-load.ts) and the mcp-server size fence.
 */

/** rewrite every central-directory record's declared uncompressed size */
function patchCentralSizes(bytes: Uint8Array, size: number): Uint8Array {
  const out = bytes.slice()
  const dv = new DataView(out.buffer, out.byteOffset, out.byteLength)
  for (let i = 0; i + 28 <= out.length; i++) {
    if (out[i] === 0x50 && out[i + 1] === 0x4b && out[i + 2] === 0x01 && out[i + 3] === 0x02) {
      dv.setUint32(i + 24, size, true)
    }
  }
  return out
}

describe('open size fences (SEC-1304)', () => {
  it('opens a legitimate deck unchanged', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(opened.deck.slides.length).toBe(1)
  })

  it('refuses a raw buffer over the 512 MiB open cap before parsing', async () => {
    // allocated, not sparse: the cap is a bytes.byteLength check, so the
    // refusal must land before JSZip even looks for the central directory
    const oversized = new Uint8Array(512 * 1024 * 1024 + 1)
    await expect(openPptx(oversized)).rejects.toThrow(
      /pptx rejected: \d+ raw bytes exceeds the 536870912 open cap/,
    )
  })

  it('refuses a bomb whose central directory declares a huge part', async () => {
    const bomb = patchCentralSizes(await createBlankPptx(), 3 * 1024 * 1024 * 1024)
    await expect(openPptx(bomb)).rejects.toThrow(
      /pptx rejected: part .* declares 3221225472 uncompressed bytes/,
    )
  })

  it('refuses a bomb whose declared parts sum over the total budget', async () => {
    // many legit-sized entries whose declared sizes add up past 1.5 GiB
    const zip = new JSZip()
    for (let i = 0; i < 20; i++) zip.file(`ppt/slides/slide${i}.xml`, '<ppt/>')
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const bomb = patchCentralSizes(bytes, 100 * 1024 * 1024)
    await expect(openPptx(bomb)).rejects.toThrow(
      /pptx rejected: total uncompressed size 2097152000 exceeds the 1610612736 limit/,
    )
  })

  it('refuses a package with more parts than the 10000 limit', async () => {
    const zip = new JSZip()
    for (let i = 0; i < 10_001; i++) zip.file(`ppt/media/f${i}.bin`, 'x')
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    await expect(openPptx(bytes)).rejects.toThrow(
      /pptx rejected: 10001 parts exceeds the 10000 limit/,
    )
  })
})
