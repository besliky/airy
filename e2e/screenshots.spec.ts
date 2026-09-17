import { test } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import { copyFile, mkdir, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { ElectronApplication, Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * README showcase screenshots. Regenerate with:
 *   npm run build:all && npm run fixtures -w @airy-office/sheets
 *   npm run test:e2e -- screenshots
 * Captured at a fixed 1440x900 window and 2x device scale for crispness,
 * light theme (the default), and committed under docs/assets/screenshots/.
 */
const OUT_DIR = resolve(__dirname, '../docs/assets/screenshots')
const WINDOW = { width: 1440, height: 900 }

/** deterministic window size for every capture */
async function fixViewport(app: ElectronApplication): Promise<void> {
  await app.evaluate(({ BrowserWindow }, size) => {
    const win = BrowserWindow.getAllWindows()[0]
    win?.setContentSize(size.width, size.height)
  }, WINDOW)
}

/**
 * 2x capture: emulate a retina display and pull the bitmap straight from
 * `Page.captureScreenshot`. Playwright's own `page.screenshot` bypasses the
 * CDP device-metrics override and always emits a 1x image in Electron, so the
 * raw CDP call is what actually honors `deviceScaleFactor: 2`.
 */
async function capture(page: Page, name: string): Promise<void> {
  const session = await page.context().newCDPSession(page)
  await session.send('Emulation.setDeviceMetricsOverride', {
    width: WINDOW.width,
    height: WINDOW.height,
    deviceScaleFactor: 2,
    mobile: false,
  })
  const shot = await session.send('Page.captureScreenshot', { format: 'png' })
  await writeFile(join(OUT_DIR, `${name}.png`), Buffer.from(shot.data, 'base64'))
  await session.send('Emulation.clearDeviceMetricsOverride')
  await session.detach()
}

test.describe('README screenshots', () => {
  test('home, docs, sheets, and slides showcase captures', async () => {
    test.setTimeout(180_000)
    await mkdir(OUT_DIR, { recursive: true })
    const scratch = await mkdtemp(join(tmpdir(), 'airy-shots-'))

    // ── Home ──
    const home = await launchShell({ onboardingSeen: true, videoDir: 'shots-home' })
    try {
      await fixViewport(home.app)
      await home.page.waitForSelector('.home', { timeout: 30_000 })
      await home.page.waitForTimeout(1_000)
      await capture(home.page, 'home')
    } finally {
      await closeAndSaveVideo(home, 'shots-home')
    }

    // ── Docs (rich fixture: headings, tables, images) ──
    const docx = join(scratch, 'shot.docx')
    await copyFile(resolve(__dirname, '../fixtures/generated/kitchen-sink.docx'), docx)
    const docs = await launchShell({
      onboardingSeen: true,
      videoDir: 'shots-docs',
      openFile: docx,
    })
    try {
      await fixViewport(docs.app)
      const editor = await waitForPageWithUrl(docs.app, 'docs/out')
      await editor.locator('.doc-page').first().waitFor({ timeout: 30_000 })
      await editor.evaluate(() => document.fonts.ready.then(() => undefined))
      await editor.waitForTimeout(800)
      await capture(editor, 'docs')
    } finally {
      await closeAndSaveVideo(docs, 'shots-docs')
    }

    // ── Sheets (data-rich multi-sheet workbook) ──
    const xlsx = join(scratch, 'shot.xlsx')
    await copyFile(
      resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-kitchen-sink.xlsx'),
      xlsx,
    )
    const sheets = await launchShell({
      onboardingSeen: true,
      videoDir: 'shots-sheets',
      openFile: xlsx,
    })
    try {
      await fixViewport(sheets.app)
      const editor = await waitForPageWithUrl(sheets.app, 'sheets/out')
      // the fixture's first sheet is "Data": wait until its tab strip names it
      await editor.waitForFunction(
        () =>
          document.querySelectorAll('canvas').length > 0 &&
          (document.body.textContent ?? '').includes('Data'),
        null,
        { timeout: 30_000 },
      )
      await editor.waitForTimeout(1_500)
      await capture(editor, 'sheets')
    } finally {
      await closeAndSaveVideo(sheets, 'shots-sheets')
    }

    // ── Slides (title/subtitle/body deck, kept as OOXML parts) ──
    const pptx = join(scratch, 'shot.pptx')
    execFileSync('zip', ['-X', '-q', '-r', pptx, '.'], {
      cwd: resolve(__dirname, 'assets/readme-deck'),
    })
    const slides = await launchShell({
      onboardingSeen: true,
      videoDir: 'shots-slides',
      openFile: pptx,
    })
    try {
      await fixViewport(slides.app)
      const editor = await waitForPageWithUrl(slides.app, 'slides/out')
      await editor.waitForSelector('.stage-wrap canvas', { timeout: 30_000 })
      await editor.waitForTimeout(1_500)
      await capture(editor, 'slides')
    } finally {
      await closeAndSaveVideo(slides, 'shots-slides')
    }
  })
})
