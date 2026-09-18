import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, waitForSheetsGrid } from './helpers'

// the preload exposes window.__airyDebug only under this env var
process.env.AIRY_DEBUG_HOOKS = '1'

/** rows the active sheet's filter currently hides, through Univer's Facade */
function filteredOutRows(page: import('@playwright/test').Page): Promise<number[]> {
  return page.evaluate(() => {
    const debug = (window as unknown as Record<string, unknown>).__airyDebug as {
      univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
    }
    const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
      getFilter(): { getFilteredOutRows(): number[] } | null
    }
    return sheet.getFilter()?.getFilteredOutRows() ?? []
  })
}

/**
 * Regression for "filter dropdown selections vanish after reopening the
 * file" (alpha feedback): the save wrote each column's criteria into
 * the xlsx autoFilter, but reopening only restored the filter range — the
 * criteria were lost and the filtered-out rows came back as plain manual
 * hides, so the dropdown lost its checked values and other columns' lists
 * were no longer narrowed. The reopen must restore the criteria AND hand the
 * hidden rows back to the filter model, so a later criteria change can
 * unhide them.
 */
test.describe('sheets: filter criteria survive save and reopen', () => {
  test('criteria restore, rows stay filtered, and re-filtering unhides', async () => {
    test.setTimeout(180_000)
    const scratch = await mkdtemp(join(tmpdir(), 'airy-filter-reopen-'))
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-filter-reopen',
    })
    let savedPath = ''
    try {
      const { app, page } = launched
      await app.evaluate(({ app: electronApp }, dir) => {
        electronApp.setPath('documents', dir)
      }, scratch)

      await expect(page.locator('.quick-card').nth(1)).toContainText('AI Sheets')
      await page.locator('.quick-card').nth(1).click()

      const sheets = await waitForPageWithUrl(app, 'sheets/out')
      await waitForSheetsGrid(sheets)

      // A1:B5 — header row plus four data rows, then filter B to "keep"
      await sheets.evaluate(async () => {
        const debug = (window as unknown as Record<string, unknown>).__airyDebug as {
          univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
        }
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
          getRange(
            row: number,
            column: number,
            rows: number,
            columns: number,
          ): { setValues(values: unknown[][]): Promise<unknown>; createFilter(): unknown }
          getFilter(): {
            setColumnFilterCriteria(column: number, criteria: unknown): unknown
          } | null
        }
        await sheet.getRange(0, 0, 5, 2).setValues([
          ['Name', 'Cat'],
          ['a', 'keep'],
          ['b', 'drop'],
          ['c', 'keep'],
          ['d', 'drop'],
        ])
        sheet.getRange(0, 0, 5, 2).createFilter()
        sheet.getFilter()?.setColumnFilterCriteria(1, {
          colId: 1,
          filters: { filters: ['keep'] },
        })
      })
      // the criteria apply asynchronously — poll the filter model until the
      // rows are really filtered out before saving
      await expect.poll(() => filteredOutRows(sheets)).toEqual([2, 4])

      // The quick-created workbook is untitled-staged (4eb93d5): its first
      // plain Save opens the Save dialog anchored in the default save dir —
      // stub the native dialog to confirm under the untitled name there so
      // the reopen step gets a deterministic path.
      const saveDir = join(scratch, 'Airy')
      savedPath = join(saveDir, 'Untitled Spreadsheet.xlsx')
      await app.evaluate(({ dialog }, target) => {
        dialog.showSaveDialog = (async () => ({
          canceled: false,
          filePath: target,
        })) as typeof dialog.showSaveDialog
      }, savedPath)

      await app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'save')
      })

      await expect(async () => {
        const xml = execSync(`unzip -p "${savedPath}" xl/worksheets/sheet1.xml`).toString()
        expect(xml).toContain('<autoFilter ref="A1:B5">')
        expect(xml).toContain('<filterColumn colId="1"><filters><filter val="keep"/></filters>')
        expect(xml).toMatch(/<row r="3"[^>]* hidden="1"/)
        expect(xml).toMatch(/<row r="5"[^>]* hidden="1"/)
      }).toPass({ timeout: 15_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-filter-reopen')
    }

    // reopen the saved file in a fresh app instance
    const relaunched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-filter-reopen-reopened',
      openFile: savedPath,
    })
    try {
      const { app } = relaunched
      const sheets = await waitForPageWithUrl(app, 'sheets/out')
      await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
        timeout: 30_000,
      })

      // the restore runs once the sheet finishes indexing — poll for it
      await sheets.waitForFunction(
        () => {
          const debug = (window as unknown as Record<string, unknown>).__airyDebug as {
            univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
          }
          const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
            getFilter(): {
              getColumnFilterCriteria(column: number): unknown
            } | null
          }
          return sheet.getFilter()?.getColumnFilterCriteria(1) != null
        },
        null,
        { timeout: 30_000 },
      )

      const restored = await sheets.evaluate(() => {
        const debug = (window as unknown as Record<string, unknown>).__airyDebug as {
          univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
        }
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
          getFilter(): {
            getColumnFilterCriteria(column: number): { filters?: { filters?: string[] } } | null
            getFilteredOutRows(): number[]
          } | null
          getSheet(): { getRowRawVisible(row: number): boolean }
        }
        const filter = sheet.getFilter()
        return {
          criteria: filter?.getColumnFilterCriteria(1)?.filters?.filters ?? null,
          filteredOut: filter?.getFilteredOutRows() ?? null,
          // raw visibility ignores the filter: true = no manual hd flag left
          rawVisible: [2, 4].map((row) => sheet.getSheet().getRowRawVisible(row)),
        }
      })
      // the dropdown's checked values are back...
      expect(restored.criteria).toEqual(['keep'])
      // ...the filter (not manual hides) owns the hidden rows...
      expect(restored.filteredOut).toEqual([2, 4])
      expect(restored.rawVisible).toEqual([true, true])

      // ...so broadening the filter really unhides them
      await sheets.evaluate(() => {
        const debug = (window as unknown as Record<string, unknown>).__airyDebug as {
          univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
        }
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
          getFilter(): {
            setColumnFilterCriteria(column: number, criteria: unknown): unknown
          } | null
        }
        sheet.getFilter()?.setColumnFilterCriteria(1, {
          colId: 1,
          filters: { filters: ['keep', 'drop'] },
        })
      })
      // broadening unhide is asynchronous too — poll until nothing is
      // filtered out instead of a fixed settle
      await expect.poll(() => filteredOutRows(sheets)).toEqual([])
      const widened = await sheets.evaluate(() => {
        const debug = (window as unknown as Record<string, unknown>).__airyDebug as {
          univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
        }
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
          getFilter(): { getFilteredOutRows(): number[] } | null
        }
        return sheet.getFilter()?.getFilteredOutRows() ?? null
      })
      expect(widened).toEqual([])
    } finally {
      await closeAndSaveVideo(relaunched, 'sheets-filter-reopen-reopened')
    }
  })
})
