import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { mkdtemp, readdir } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  launchShell,
  closeAndSaveVideo,
  waitForPageWithUrl,
  screenshotPath,
  waitForSheetsGrid,
} from './helpers'

/**
 * Regression for "new spreadsheet cannot be saved" (feedback 2368785): the
 * quick-create card must stage a real backing .xlsx so the save pipeline
 * works from the first edit — but since untitled staging (4eb93d5) the blank
 * lives under userData/untitled-staging, never in the save dir; the first
 * Save materializes it in the default save folder through the Save dialog.
 */
test.describe('sheets: new blank workbook', () => {
  test('quick-create stages the workbook and the first save lands in the save dir', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'airy-sheets-blank-'))
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-new-blank' })
    try {
      const { app, page } = launched
      // keep the saved workbook out of the real ~/Documents/Airy
      await app.evaluate(({ app: electronApp }, dir) => {
        electronApp.setPath('documents', dir)
      }, scratch)

      await expect(page.locator('.quick-card').nth(1)).toContainText('AI Sheets')
      await page.locator('.quick-card').nth(1).click()

      const sheets = await waitForPageWithUrl(app, 'sheets/out')
      await waitForSheetsGrid(sheets)

      // the backing file exists before any edit — staged under userData, and
      // nothing is written into the default save dir yet
      const saveDir = join(scratch, 'Airy')
      const stagingDir = join(launched.userDataDir, 'untitled-staging')
      const staged = (await readdir(stagingDir)).filter((f) => f.endsWith('.xlsx'))
      expect(staged).toEqual(['Untitled Spreadsheet.xlsx'])
      const premature = await readdir(saveDir).catch(() => [])
      expect(premature.filter((f) => f.endsWith('.xlsx'))).toHaveLength(0)
      const workbook = join(saveDir, 'Untitled Spreadsheet.xlsx')

      const grid = await sheets.evaluate(() => {
        for (const canvas of document.querySelectorAll('canvas')) {
          const rect = canvas.getBoundingClientRect()
          if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
        }
        return null
      })
      if (!grid) throw new Error('worksheet canvas not found')
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)
      await expect(sheets.locator('.name-box')).toHaveValue('A1')
      await sheets.keyboard.type('42', { delay: 50 })
      await sheets.keyboard.press('Enter')
      await sheets.screenshot({ path: screenshotPath('sheets-new-blank-edited') })

      // The staged tab's first plain Save opens the Save dialog anchored in
      // the default save dir with the untitled name — stub the native dialog
      // to capture that anchor and confirm into the scratch save dir. (The
      // target path is built on the spec side: closures over Node imports
      // don't survive evaluate serialization.)
      await app.evaluate(({ dialog }, targetPath) => {
        const marker = globalThis as { __airySaveDialogDefaultPath?: string }
        dialog.showSaveDialog = (async (...args: unknown[]) => {
          const options = (args.length > 1 ? args[1] : args[0]) as { defaultPath?: string }
          marker.__airySaveDialogDefaultPath = options?.defaultPath
          return { canceled: false, filePath: targetPath }
        }) as typeof dialog.showSaveDialog
      }, workbook)

      await app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'save')
      })
      await expect(() => {
        const xml = execSync(`unzip -p "${workbook}" xl/worksheets/sheet1.xml`).toString()
        expect(xml).toContain('<v>42</v>')
      }).toPass({ timeout: 15_000 })

      // the save dialog was anchored in the default save dir under the
      // untitled name (suggestSaveAs for staged workbooks)
      const defaultPath = await app.evaluate(() => {
        const marker = globalThis as { __airySaveDialogDefaultPath?: string }
        return marker.__airySaveDialogDefaultPath
      })
      expect(defaultPath).toBe(workbook)

      // the first save materialized exactly one .xlsx in the save dir and
      // removed the staged file (the tab rebound to the picked path)
      expect((await readdir(saveDir)).filter((f) => f.endsWith('.xlsx'))).toEqual([
        'Untitled Spreadsheet.xlsx',
      ])
      expect(existsSync(join(stagingDir, staged[0]))).toBe(false)
    } finally {
      await closeAndSaveVideo(launched, 'sheets-new-blank')
    }
  })
})
