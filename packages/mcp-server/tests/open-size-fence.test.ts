// Open-time size fences (SEC-1102): the binary document sessions used to
// read whatever was on disk — a hostile file pulled into an agent's
// workspace could balloon the headless server's memory as a multi-gigabyte
// raw file or as a zip bomb (a few KiB declaring gigabytes of uncompressed
// entries). Raw stat-first caps now refuse oversized files before a byte is
// read (docx/slides/xlsx), and the slides open additionally walks the zip
// central directory's declared uncompressed sizes against the same budget
// the docx-engine fence uses. For workbooks the zip budget runs inside the
// Rust sidecar (SEC-1103); the test at the bottom pins that its refusals
// surface through the session open unchanged, and xlsx-integration.test.ts
// drives the real binary against a forged bomb.
import { mkdtemp, open, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import JSZip from 'jszip'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { DocxSession } from '../src/docx/session.js'
import { SlidesSession } from '../src/slides/session.js'
import { XlsxSession } from '../src/xlsx/session.js'
import type { XlsxIo } from '../src/xlsx/sidecar-client.js'
import { buildFixturePptx } from './helpers/pptx-fixture.js'
import { makeStubIo } from './helpers/stub-sidecar.js'
import { patchCentralSizes } from './helpers/zip-bomb.js'

let root: string

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'airy-size-fence-'))
})

afterEach(async () => {
  await rm(root, { recursive: true, force: true })
})

/** a file whose stat size is huge but that occupies no real space (sparse) */
async function writeSparseHuge(path: string, size: number): Promise<void> {
  const handle = await open(path, 'w')
  try {
    await handle.truncate(size)
  } finally {
    await handle.close()
  }
}

describe('open size fences (SEC-1102)', () => {
  it('slides refuse a raw file over the 512 MiB cap before reading it', async () => {
    const huge = join(root, 'huge.pptx')
    await writeSparseHuge(huge, 512 * 1024 * 1024 + 1)
    await expect(SlidesSession.open(huge, root)).rejects.toThrow(
      /huge\.pptx" is \d+ bytes; slides sessions cap open size at 536870912 bytes/,
    )
  })

  it('docx refuse a raw file over the 512 MiB cap before reading it', async () => {
    const huge = join(root, 'huge.docx')
    await writeSparseHuge(huge, 600 * 1024 * 1024)
    await expect(DocxSession.open(huge, root)).rejects.toThrow(
      /docx sessions cap open size at 536870912 bytes/,
    )
  })

  it('xlsx refuse a raw file over the 512 MiB cap before any sidecar or conversion work', async () => {
    const huge = join(root, 'huge.xlsx')
    await writeSparseHuge(huge, 600 * 1024 * 1024)
    // the io must never be reached: the fence refuses before the sidecar is
    // spawned or LibreOffice converts
    const refuse = async (): Promise<never> => {
      throw new Error('io must not be called for an over-cap file')
    }
    const io = { open: refuse, convertWorkbook: refuse } as unknown as XlsxIo
    await expect(XlsxSession.open(huge, root, io)).rejects.toThrow(
      /xlsx sessions cap open size at 536870912 bytes/,
    )
  })

  it('xlsx surface the sidecar zip-bomb refusal unchanged (the budget lives in the sidecar)', async () => {
    // SEC-1103: workbook bytes never transit Node, so the declared-size
    // budget runs inside the Rust sidecar and reaches this session as the
    // open reply's error. The stub replays the exact refusal message the
    // sidecar emits; the session must pass it through untouched — any
    // generic "Cannot read" wrap here would hide the fence from the agent.
    const book = join(root, 'bomb.xlsx')
    await writeFile(book, 'stub-xlsx-bytes')
    const io = makeStubIo({
      openError: new Error(
        'Workbook declares 2147483648 uncompressed bytes across its ZIP entries, ' +
          'above the 1610612736 byte (1.5 GiB) open budget — the file may be a zip bomb.',
      ),
    })
    await expect(XlsxSession.open(book, root, io)).rejects.toThrow(
      /Workbook declares \d+ uncompressed bytes.*open budget/,
    )
    expect(io.calls.open).toEqual([book])
  })

  it('slides refuse a zip bomb declared in the central directory (per part)', async () => {
    const bomb = patchCentralSizes(await buildFixturePptx(), 600 * 1024 * 1024)
    const path = join(root, 'bomb.pptx')
    await writeFile(path, bomb)
    await expect(SlidesSession.open(path, root)).rejects.toThrow(
      /pptx rejected: part .* declares \d+ uncompressed bytes/,
    )
  })

  it('slides refuse a zip bomb declared in the central directory (total)', async () => {
    // 450 MiB per part across the fixture's >= 4 parts exceeds the 1.5 GiB
    // total while every single part stays under the per-part limit
    const bomb = patchCentralSizes(await buildFixturePptx(), 450 * 1024 * 1024)
    const path = join(root, 'bomb-total.pptx')
    await writeFile(path, bomb)
    await expect(SlidesSession.open(path, root)).rejects.toThrow(/total uncompressed size/)
  })

  it('slides refuse an archive with too many parts', async () => {
    const zip = await JSZip.loadAsync(await buildFixturePptx())
    for (let i = 0; i < 10_010; i++) zip.file(`junk/${String(i)}.bin`, 'x')
    const path = join(root, 'many-parts.pptx')
    await writeFile(path, await zip.generateAsync({ type: 'uint8array' }))
    await expect(SlidesSession.open(path, root)).rejects.toThrow(/parts exceeds the 10000 limit/)
  })

  it('a normal deck still opens under the fences', async () => {
    const path = join(root, 'fine.pptx')
    await writeFile(path, await buildFixturePptx())
    const session = await SlidesSession.open(path, root)
    expect(session.meta().slideCount).toBe(1)
  })
})
