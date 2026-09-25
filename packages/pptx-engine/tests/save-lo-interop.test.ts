/**
 * Integration with a REAL LibreOffice (soffice on PATH, the `libreoffice` snap
 * wrapper, or AIRY_SOFFICE): every package written by savePptxToFile must load in
 * Impress. This is the BUG-1671 end-to-end gate — LibreOffice rejected the streamed
 * writer's data-descriptor zip with "source file could not be loaded" while exiting
 * 0, so the assertion is that the converted PDF is actually produced, not the exit
 * code. Skipped with an explicit reason when LibreOffice is absent.
 *
 * Scratch files live under a non-hidden directory in $HOME: snap-confined
 * LibreOffice cannot read /tmp (nor hidden directories), see LOGS/BUG-1607.
 * python-pptx is probed opportunistically as a second independent loader.
 */
import { describe, it, expect, afterAll } from 'vitest'
import { execFile, execFileSync } from 'node:child_process'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync } from 'node:fs'
import { homedir } from 'node:os'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { promisify } from 'node:util'
import { openPptx, savePptxToFile, addElement, addPicture } from '../src/index'
import { noisePng } from './helpers/test-media'

const execFileAsync = promisify(execFile)
const here = dirname(fileURLToPath(import.meta.url))
const fxPath = (name: string) => join(here, 'fixtures', name)

function findSoffice(): string | null {
  const fromEnv = process.env.AIRY_SOFFICE
  if (fromEnv) return existsSync(fromEnv) ? fromEnv : null
  for (const name of ['soffice', 'libreoffice']) {
    for (const dir of (process.env.PATH ?? '').split(':')) {
      if (!dir) continue
      const candidate = join(dir, name)
      if (existsSync(candidate)) return candidate
    }
  }
  return null
}

function hasPythonPptx(): boolean {
  try {
    execFileSync('python3', ['-c', 'import pptx'], { stdio: 'ignore' })
    return true
  } catch {
    return false
  }
}

const soffice = findSoffice()
const describeWithSoffice = describe.skipIf(!soffice)
const hasPyPptx = hasPythonPptx()

// Snap-confined LO cannot read /tmp; a non-hidden $HOME path works.
const workRoot = join(homedir(), 'airy-lo-interop')
let work: string | null = null

function ensureWork(): string {
  if (!work) {
    mkdirSync(workRoot, { recursive: true })
    work = mkdtempSync(join(workRoot, 'bug1671-'))
  }
  return work
}

afterAll(() => {
  if (work) rmSync(work, { recursive: true, force: true })
})

async function convertToPdf(pptxPaths: string[]): Promise<void> {
  const dir = dirname(pptxPaths[0]!)
  // Isolated profile: never touch a user's running LibreOffice instance.
  const profile = join(dir, '.lo-profile')
  const { stderr } = await execFileAsync(
    soffice!,
    [
      '--headless',
      '--convert-to',
      'pdf',
      '--outdir',
      dir,
      `-env:UserInstallation=file://${profile}`,
      ...pptxPaths,
    ],
    { timeout: 240_000 },
  )
  // PERF-1720: all decks go through ONE soffice process — per-deck process
  // startup (~6 s each on a snap install) dominated this file's serial wall.
  // Every deck is still asserted individually: LO converts each input on its
  // own and exits 0 even when a source fails to load ("source file could not
  // be loaded" on stderr, no output) — the produced PDF is the real assertion.
  for (const pptxPath of pptxPaths) {
    const pdf = pptxPath.replace(/\.pptx$/, '.pdf')
    expect(statSync(pdf).size, `PDF produced for ${pptxPath}; stderr: ${stderr}`).toBeGreaterThan(0)
  }
}

describeWithSoffice(
  'LibreOffice opens pptx saved by savePptxToFile ' +
    '(skipped when absent: no soffice/libreoffice on PATH and AIRY_SOFFICE unset — install LibreOffice to run)',
  () => {
    it('converts saved decks to PDF: as-is, after an edit, and media-heavy', async () => {
      work = ensureWork()

      const edited = await openPptx(
        new Uint8Array(readFileSync(fxPath('01_standard_business.pptx'))),
      )
      addElement(edited.deck.slides[0]!, {
        kind: 'textbox',
        offset: { x: 914400, y: 914400, cx: 6096000, cy: 914400 },
        paragraphs: [{ runs: [{ text: 'LO interop after edit' }] }],
      })

      const mediaDeck = await openPptx(
        new Uint8Array(readFileSync(fxPath('01_standard_business.pptx'))),
      )
      const slide = mediaDeck.deck.slides[0]!
      for (let i = 0; i < 3; i++) {
        addPicture(mediaDeck, slide, {
          bytes: noisePng(1600, 400, 0x9e3779b9 + i),
          ext: 'png',
          offset: { x: 914400 + i * 914400, y: 3200400, cx: 1828800, cy: 457200 },
          name: `bulk ${i}`,
        })
      }

      const decks: Array<[string, typeof edited]> = [
        [
          'standard-as-is.pptx',
          await openPptx(new Uint8Array(readFileSync(fxPath('01_standard_business.pptx')))),
        ],
        [
          'unicode-as-is.pptx',
          await openPptx(new Uint8Array(readFileSync(fxPath('05_unicode_cjk_emoji.pptx')))),
        ],
        ['edited.pptx', edited],
        ['media-heavy.pptx', mediaDeck],
      ]

      const savedPaths: string[] = []
      for (const [name, opened] of decks) {
        const path = join(work, name)
        await savePptxToFile(opened, path)
        savedPaths.push(path)
      }
      await convertToPdf(savedPaths)
    }, 600_000)

    it.skipIf(!hasPyPptx)(
      'reopens a saved deck with python-pptx',
      async () => {
        work = ensureWork()
        const target = join(work, 'pycheck.pptx')
        const opened = await openPptx(
          new Uint8Array(readFileSync(fxPath('01_standard_business.pptx'))),
        )
        await savePptxToFile(opened, target)
        execFileSync(
          'python3',
          ['-c', `from pptx import Presentation; Presentation(${JSON.stringify(target)})`],
          { stdio: 'ignore' },
        )
      },
      60_000,
    )
  },
)
