import { test, expect } from '@playwright/test'
import { execFileSync, execSync } from 'node:child_process'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { Page } from '@playwright/test'
import {
  launchShell,
  closeAndSaveVideo,
  waitForPageWithUrl,
  screenshotPath,
  waitForPaintSettled,
} from './helpers'

/**
 * The fixture deck reuses the font-manager corpus parts (a plain single-slide
 * pptx kept as OOXML text under assets/ so no binary lives in the repo) — its
 * headline is an ordinary text run, which is all the edit round-trip needs.
 * Geometry comes from the slide XML: the shape sits at off(914400, 914400)
 * ext(6096000, 1143000) on a 12192000x6858000 (16:9) canvas.
 */
async function buildFixtureDeck(): Promise<string> {
  const out = join(await mkdtemp(join(tmpdir(), 'airy-slides-e2e-')), 'edit-save.pptx')
  execFileSync('zip', ['-X', '-q', '-r', out, '.'], {
    cwd: resolve(__dirname, 'assets/font-manager-rubik'),
  })
  return out
}

/** center of the headline shape in slide-relative fractions */
const HEADLINE = { x: 0.325, y: 0.2166 }
/** fixture slide size (16:9) in EMU: the slide is aspect-fit inside the canvas */
const SLIDE_ASPECT = 12192000 / 6858000

/**
 * The headline position in viewport coordinates. The konva stage letterboxes
 * the slide inside its canvas (canvas aspect ≠ slide aspect), so the slide
 * rect is recovered by fitting the slide width and centering vertically.
 */
async function headlinePoint(page: Page): Promise<{ x: number; y: number }> {
  const rect = await page.evaluate(() => {
    for (const canvas of document.querySelectorAll('canvas')) {
      const r = canvas.getBoundingClientRect()
      if (r.width > 500 && r.height > 300) return { x: r.x, y: r.y, w: r.width, h: r.height }
    }
    return null
  })
  if (!rect) throw new Error('slide canvas not found')
  const slideHeight = rect.w / SLIDE_ASPECT
  const slideTop = rect.y + (rect.h - slideHeight) / 2
  return { x: rect.x + HEADLINE.x * rect.w, y: slideTop + HEADLINE.y * slideHeight }
}

function slideXml(pptxPath: string): string {
  return execSync(`unzip -p "${pptxPath}" ppt/slides/slide1.xml`).toString()
}

test.describe('slides: edit a text run and save the deck', () => {
  test('edited headline text round-trips into the pptx', async () => {
    test.setTimeout(120_000)
    const marker = `E2E ${Date.now()}`
    const deck = await buildFixtureDeck()

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'slides-edit-save',
      openFile: deck,
    })
    try {
      const editor = await waitForPageWithUrl(launched.app, 'slides/out')
      await editor.waitForSelector('.stage-wrap canvas', { timeout: 30_000 })
      // font-triggered canvas redraws queue behind fonts.ready — wait for the
      // paint to settle so the dblclick overlay positions over real text
      await waitForPaintSettled(editor)

      // Select the headline shape, then double-click it: the run-level
      // contentEditable overlay opens over the canvas text (TextEditOverlay).
      const headline = await headlinePoint(editor)
      await editor.mouse.click(headline.x, headline.y)
      // small settle so the selection click and the dblclick register as two
      // gestures (Konva pairs a too-fast second click into the first one);
      // no DOM signal exists for a canvas selection commit
      await editor.waitForTimeout(300)
      await editor.mouse.dblclick(headline.x, headline.y)
      const overlay = editor.locator('[contenteditable="true"]').first()
      await overlay.waitFor({ timeout: 10_000 })

      // Append to the run (keep the existing text so the assertion can also
      // check it survived), then Esc commits (TextEditOverlay semantics).
      await overlay.click()
      await editor.keyboard.press('End')
      await editor.keyboard.type(` ${marker}`, { delay: 20 })
      await editor.keyboard.press('Escape')
      await expect(overlay).toHaveCount(0, { timeout: 10_000 })
      await editor.screenshot({ path: screenshotPath('slides-edited') })

      // File > Save, routed to the slides view the same way the app menu does it
      await launched.app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('slides/out'))
        wc?.send('slides:menu', 'save')
      })
      await expect(() => {
        expect(slideXml(deck)).toContain(marker)
      }).toPass({ timeout: 20_000 })
      // the untouched prefix of the run survives the save
      expect(slideXml(deck)).toContain('Rubik headline')
    } finally {
      await closeAndSaveVideo(launched, 'slides-edit-save')
    }
  })
})
