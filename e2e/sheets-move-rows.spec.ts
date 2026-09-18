import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { Page } from '@playwright/test'
import {
  launchShell,
  closeAndSaveVideo,
  waitForPageWithUrl,
  screenshotPath,
  waitForSheetsGrid,
} from './helpers'

// the preload exposes window.__airyDebug only under this env var; the spec
// reads row order and workbook identity through Univer's Facade
process.env.AIRY_DEBUG_HOOKS = '1'

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

/** column A values in order — the ground truth for move/undo verification */
function columnA(sheets: Page): Promise<unknown[][]> {
  return sheets.evaluate(() => {
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
    return debug.univerAPI.getActiveWorkbook().getActiveSheet().getRange(0, 0, 4, 1).getValues()
  })
}

/** Univer unit id — deterministic `file-<sha>`, so it changes when a save
 * reopens the session over the newly written file */
function workbookId(sheets: Page): Promise<string> {
  return sheets.evaluate(() => {
    const debug = (window as unknown as Record<string, unknown>).__airyDebug as {
      univerAPI: { getActiveWorkbook(): { getId(): string } }
    }
    return debug.univerAPI.getActiveWorkbook().getId()
  })
}

async function gridOrigin(page: Page): Promise<{ x: number; y: number }> {
  const grid = await page.evaluate(() => {
    for (const canvas of document.querySelectorAll('canvas')) {
      const rect = canvas.getBoundingClientRect()
      if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
    }
    return null
  })
  if (!grid) throw new Error('worksheet canvas not found')
  return grid
}

function cellPoint(origin: { x: number; y: number }, row: number, column: number) {
  return { x: origin.x + 46 + column * 86 + 43, y: origin.y + 24 + row * 23 + 11 }
}

function rowHeaderPoint(origin: { x: number; y: number }, row: number) {
  return { x: origin.x + 23, y: origin.y + 24 + row * 23 + 11 }
}

test.describe('sheets: whole-row move', () => {
  test('drag-moving a row persists the new order and undo restores it', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'airy-moverows-e2e-'))
    const workbook = join(scratch, 'move.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-move-rows',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForSheetsGrid(sheets)
      const origin = await gridOrigin(sheets)

      for (const [rowIndex, value] of ['one', 'two', 'three', 'four'].entries()) {
        const point = cellPoint(origin, rowIndex, 0)
        await sheets.mouse.click(point.x, point.y)
        await sheets.keyboard.type(value, { delay: 30 })
        await sheets.keyboard.press('Enter')
      }

      const dragRowDown = async () => {
        const header2 = rowHeaderPoint(origin, 1)
        await sheets.mouse.click(header2.x, header2.y)
        // let the header click commit its row selection before the drag
        // starts (a too-fast drag pairs into the click); no DOM signal
        // exists for a canvas selection commit — small settle ≤300ms
        await sheets.waitForTimeout(300)
        await sheets.mouse.move(header2.x, header2.y)
        await sheets.mouse.down()
        const target = rowHeaderPoint(origin, 3)
        await sheets.mouse.move(header2.x, header2.y + 12, { steps: 4 })
        await sheets.mouse.move(target.x, target.y + 8, { steps: 12 })
        await sheets.mouse.up()
        // the move command lands asynchronously — poll the model until the
        // rows actually changed before acting on the new order
        await expect.poll(() => columnA(sheets)).toEqual([['one'], ['three'], ['four'], ['two']])
      }
      const savedOrder = async (expected: readonly string[]) => {
        await launched.app.evaluate(({ webContents }) => {
          const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
          wc?.send('menu:action', 'save')
        })
        await expect(() => {
          const xml = execSync(`unzip -p "${workbook}" xl/worksheets/sheet1.xml`).toString()
          const order = [...xml.matchAll(/<is><t[^>]*>(one|two|three|four)<\/t><\/is>/g)].map(
            (m) => m[1],
          )
          expect(order).toEqual(expected)
        }).toPass({ timeout: 20_000 })
      }

      // Move, then undo before saving: the inverse move cancels in the
      // journal and the file keeps the original order.
      await dragRowDown()
      await sheets.screenshot({ path: screenshotPath('move-rows-after-drag') })
      const cell = cellPoint(origin, 0, 2)
      await sheets.mouse.click(cell.x, cell.y)
      await sheets.keyboard.press('ControlOrMeta+z')
      // undo applies asynchronously too — wait for the original order to be
      // back before the save snapshots the journal
      await expect.poll(() => columnA(sheets)).toEqual([['one'], ['two'], ['three'], ['four']])
      await sheets.screenshot({ path: screenshotPath('move-rows-after-undo') })
      const idBeforeSave = await workbookId(sheets)
      await savedOrder(['one', 'two', 'three', 'four'])

      // Saving reopens the session; move again and save the new order. The
      // reopen swaps the Univer unit (`file-<sha>` of the new content) —
      // poll for the id change so the second drag hits the reopened grid,
      // not the teardown of the old one.
      await expect.poll(() => workbookId(sheets), { timeout: 30_000 }).not.toBe(idBeforeSave)
      await dragRowDown()
      await savedOrder(['one', 'three', 'four', 'two'])
    } finally {
      await closeAndSaveVideo(launched, 'sheets-move-rows')
    }
  })
})
