import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, waitForSheetsGrid } from './helpers'

// the preload exposes window.__airyDebug only under this env var
process.env.AIRY_DEBUG_HOOKS = '1'

/**
 * Regression for "tile-pasted cells vanish after save" (alpha
 * feedback): pasting a copied formula row into a taller target writes follower
 * cells as {si, v} — shared-formula id without formula text. The journal
 * must materialize those into real formulas; before the fix the follow-up
 * recalc mutation wiped them and the saved file lost the whole block.
 */
test.describe('sheets: tiled paste of formulas survives save', () => {
  test('followers keep their (shifted) formula in the saved xlsx', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'airy-paste-save-'))
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-paste-shared-formula-save',
    })
    try {
      const { app, page } = launched
      await app.evaluate(({ app: electronApp }, dir) => {
        electronApp.setPath('documents', dir)
      }, scratch)

      await expect(page.locator('.quick-card').nth(1)).toContainText('AI Sheets')
      await page.locator('.quick-card').nth(1).click()

      const sheets = await waitForPageWithUrl(app, 'sheets/out')
      await waitForSheetsGrid(sheets)

      const grid = await sheets.evaluate(() => {
        for (const canvas of document.querySelectorAll('canvas')) {
          const rect = canvas.getBoundingClientRect()
          if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
        }
        return null
      })
      if (!grid) throw new Error('worksheet canvas not found')
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)

      // source row A1:C1: two values and a formula referencing the row
      await sheets.evaluate(async () => {
        const debug = (window as unknown as Record<string, unknown>).__airyDebug as {
          univerAPI: {
            getActiveWorkbook(): {
              getActiveSheet(): {
                getRange(
                  row: number,
                  column: number,
                  rows: number,
                  columns: number,
                ): { setValues(v: unknown[][]): Promise<unknown>; activate(): unknown }
              }
            }
          }
        }
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet()
        await sheet.getRange(0, 0, 1, 3).setValues([[10, 'hi', '=A1&"-x"']])
        sheet.getRange(0, 0, 1, 3).activate()
      })
      await sheets.keyboard.press('Control+c')
      // the copy command writes the clipboard asynchronously and no DOM state
      // mirrors it — a short settle is the cheapest reliable gap (≤300ms)
      await sheets.waitForTimeout(300)

      // tile-paste into A2:C3 — row 3's formula cell becomes an si follower
      await sheets.evaluate(() => {
        const debug = (window as unknown as Record<string, unknown>).__airyDebug as {
          univerAPI: {
            getActiveWorkbook(): {
              getActiveSheet(): {
                getRange(
                  row: number,
                  column: number,
                  rows: number,
                  columns: number,
                ): { activate(): unknown }
              }
            }
          }
        }
        debug.univerAPI.getActiveWorkbook().getActiveSheet().getRange(1, 0, 2, 3).activate()
      })
      await sheets.keyboard.press('Control+v')
      // the paste lands in the edit journal asynchronously: poll the model
      // until both repetitions hold their values before saving
      await expect
        .poll(() =>
          sheets.evaluate(() => {
            const debug = (window as unknown as Record<string, unknown>).__airyDebug as {
              univerAPI: {
                getActiveWorkbook(): {
                  getActiveSheet(): {
                    getRange(
                      row: number,
                      column: number,
                      rows: number,
                      columns: number,
                    ): { getValues(): unknown[][] }
                  }
                }
              }
            }
            return debug.univerAPI
              .getActiveWorkbook()
              .getActiveSheet()
              .getRange(1, 0, 2, 1)
              .getValues()
          }),
        )
        .toEqual([[10], [10]])

      // The quick-created workbook is untitled-staged (4eb93d5): its first
      // plain Save opens the Save dialog anchored in the default save dir —
      // stub the native dialog to confirm under the untitled name there.
      const workbook = join(scratch, 'Airy', 'Untitled Spreadsheet.xlsx')
      await app.evaluate(({ dialog }, target) => {
        dialog.showSaveDialog = (async () => ({
          canceled: false,
          filePath: target,
        })) as typeof dialog.showSaveDialog
      }, workbook)

      await app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'save')
      })

      await expect(async () => {
        const xml = execSync(`unzip -p "${workbook}" xl/worksheets/sheet1.xml`).toString()
        // both pasted repetitions keep values and a row-shifted formula
        expect(xml).toContain('<c r="A2" s="1"><v>10</v></c>')
        expect(xml).toContain('<c r="A3" s="1"><v>10</v></c>')
        expect(xml).toMatch(/<c r="C2" s="1"><f>A2&amp;"-x"<\/f><\/c>/)
        expect(xml).toMatch(/<c r="C3" s="1"><f>A3&amp;"-x"<\/f><\/c>/)
      }).toPass({ timeout: 15_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-paste-shared-formula-save')
    }
  })
})
