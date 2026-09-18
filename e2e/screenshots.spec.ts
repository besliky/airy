import { test } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import { copyFile, mkdir, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { ElectronApplication, Page } from '@playwright/test'
import {
  launchShell,
  closeAndSaveVideo,
  waitForPageWithUrl,
  waitForPaintSettled,
  waitForSheetsGrid,
} from './helpers'

/**
 * README showcase screenshots. Regenerate the committed PNGs with:
 *   npm run build:all && npm run fixtures -w @airy-office/sheets
 *   UPDATE_SCREENSHOTS=1 npm run test:e2e -- screenshots
 * Captured at a fixed 1440x900 window and 2x device scale for crispness,
 * light theme (the default), and committed under docs/assets/screenshots/.
 *
 * Without UPDATE_SCREENSHOTS=1 the captures go to a temp directory: the spec
 * still proves capture works end to end, but a plain e2e run no longer
 * dirties the tree by overwriting the committed PNGs.
 */
const REPO_OUT_DIR = resolve(__dirname, '../docs/assets/screenshots')
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
async function capture(page: Page, name: string, outDir: string): Promise<void> {
  const session = await page.context().newCDPSession(page)
  await session.send('Emulation.setDeviceMetricsOverride', {
    width: WINDOW.width,
    height: WINDOW.height,
    deviceScaleFactor: 2,
    mobile: false,
  })
  const shot = await session.send('Page.captureScreenshot', { format: 'png' })
  await writeFile(join(outDir, `${name}.png`), Buffer.from(shot.data, 'base64'))
  await session.send('Emulation.clearDeviceMetricsOverride')
  await session.detach()
}

test.describe('README screenshots', () => {
  test('home, docs, sheets, and slides showcase captures', async () => {
    test.setTimeout(180_000)
    const outDir =
      process.env.UPDATE_SCREENSHOTS === '1'
        ? REPO_OUT_DIR
        : await mkdtemp(join(tmpdir(), 'airy-shots-out-'))
    await mkdir(outDir, { recursive: true })
    const scratch = await mkdtemp(join(tmpdir(), 'airy-shots-'))

    // ── Home ──
    const home = await launchShell({ onboardingSeen: true, videoDir: 'shots-home' })
    try {
      await fixViewport(home.app)
      await home.page.waitForSelector('.home', { timeout: 30_000 })
      await waitForPaintSettled(home.page)
      await capture(home.page, 'home', outDir)
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
      // capture only after font-driven reflow/repaint has settled
      await waitForPaintSettled(editor)
      await capture(editor, 'docs', outDir)
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
      // the fixture's first sheet is "Data": wait until the grid is live and
      // the first paint (with loaded fonts) has settled before capturing
      await waitForSheetsGrid(editor, 'Data')
      await waitForPaintSettled(editor)
      await capture(editor, 'sheets', outDir)
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
      await waitForPaintSettled(editor)
      await capture(editor, 'slides', outDir)
    } finally {
      await closeAndSaveVideo(slides, 'shots-slides')
    }
  })
})
