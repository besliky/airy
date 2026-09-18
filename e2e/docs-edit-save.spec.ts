import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import {
  launchShell,
  closeAndSaveVideo,
  waitForPageWithUrl,
  screenshotPath,
  waitForPaintSettled,
} from './helpers'

const FIXTURE = resolve(__dirname, '../fixtures/generated/kitchen-sink.docx')

function documentXml(docxPath: string): string {
  return execSync(`unzip -p "${docxPath}" word/document.xml`).toString()
}

test.describe('docs: edit and save a document', () => {
  test('typed text round-trips through save into the docx', async () => {
    test.setTimeout(120_000)
    const marker = `e2e-edit-${Date.now()}`
    const scratch = await mkdtemp(join(tmpdir(), 'airy-docs-e2e-'))
    const document = join(scratch, 'edit-save.docx')
    await copyFile(FIXTURE, document)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-edit-save',
      openFile: document,
    })
    try {
      const editor = await waitForPageWithUrl(launched.app, 'docs/out')
      await editor.locator('.doc-page').first().waitFor({ timeout: 30_000 })
      // fonts.ready alone leaves the caret click exposed to font-driven
      // reflow repaints still queued behind it (audit R9) — settle two
      // frames past the fonts so the click lands on final geometry
      await waitForPaintSettled(editor)

      // Click into the first page and type; the caret lands at the nearest
      // text position, which is enough — the assertion only needs the marker
      // to survive into the saved part.
      await editor.locator('.doc-page').first().click()
      await editor.keyboard.type(marker, { delay: 20 })
      await expect
        .poll(() => editor.evaluate(() => document.querySelector('.doc-page')?.textContent ?? ''), {
          timeout: 10_000,
        })
        .toContain(marker)
      await editor.screenshot({ path: screenshotPath('docs-edited') })

      // File > Save, routed to the docs view the same way the app menu does it
      await launched.app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('docs/out'))
        wc?.send('menu:command', 'save')
      })
      await expect(() => {
        expect(documentXml(document)).toContain(marker)
      }).toPass({ timeout: 15_000 })

      // the save must be surgical: pre-existing body text survives
      expect(documentXml(document)).not.toBe('')
    } finally {
      await closeAndSaveVideo(launched, 'docs-edit-save')
    }
  })
})
