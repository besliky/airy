/**
 * Shared conversion harness for the html2docx suite.
 *
 * Every case pays roughly 1.4s of fixed browser waits inside the converter
 * (network idle, fonts.ready, settle sleeps), so the suite's wall time is
 * bound by how many cases run concurrently, not by CPU. Each test file gets
 * its own Chrome and its own fixture directory (so files run in parallel
 * vitest workers), and every case in a file shares them. Contexts and pages
 * stay per conversion: that isolation is PlaywrightDriver's product
 * behavior, not suite waste.
 */
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { afterAll, beforeAll } from 'vitest'
import JSZip from 'jszip'
import { chromium } from 'playwright-core'
import type { Browser } from 'playwright-core'

import { convertHtmlToDocx } from '../../src'
import { findChrome, PlaywrightDriver } from '../../src/drivers/playwright'

// Local launch instead of src/drivers' launchChrome(): that one still passes
// --force-device-scale-factor=2 (a leftover from the DSF=2 era), which is dead
// weight here — every context below is created with an explicit
// deviceScaleFactor (1, see VIEWPORT) that overrides the browser flag. The
// flag only misleads (and would double PNG encode cost for any context ever
// created without its own DSF), so the suite launches without it (PERF-702;
// src is left untouched for its CLI consumers).
async function launchChrome(): Promise<Browser> {
  return chromium.launch({
    executablePath: findChrome(),
    headless: true,
    args: ['--no-sandbox', '--disable-dev-shm-usage'],
  })
}

/**
 * A4 at 96dpi so layout (line wraps, column gaps) matches print.
 * deviceScaleFactor 1 (was 2): assertions are structural — XML parts, IR
 * geometry in CSS pixels, media counts — and never inspect screenshot
 * pixels, so Retina density only cost PNG encode time. Verified by running
 * the whole suite at both factors: identical results, DSF 1 slightly ahead.
 */
const VIEWPORT = { width: 794, height: 1123, deviceScaleFactor: 1 }

export interface ConvertedDocument {
  zip: JSZip
  xml: (file: string) => Promise<string | undefined>
  ir: any[]
  /** Text that only survives as pixels (inside screenshots). */
  screenshotText: string
}

/**
 * Call once at the top of a test file: registers the file-scoped Chrome and
 * fixture directory as beforeAll/afterAll hooks and returns a convertHtml
 * helper that renders an HTML string through the real converter pipeline.
 */
export function setupFileConversion(): (html: string, name: string) => Promise<ConvertedDocument> {
  let browser: Browser | undefined
  let fixtureDir: string | undefined
  beforeAll(async () => {
    browser = await launchChrome()
    fixtureDir = await fs.mkdtemp(path.join(os.tmpdir(), 'html2docx-'))
  })
  afterAll(async () => {
    await browser?.close()
    await fs.rm(fixtureDir, { recursive: true, force: true })
  })

  return async function convertHtml(html: string, name: string): Promise<ConvertedDocument> {
    const input = path.join(fixtureDir!, `${name}.html`)
    await fs.writeFile(input, html)
    const driver = await PlaywrightDriver.create(browser!, VIEWPORT)
    let result
    try {
      result = await convertHtmlToDocx({ url: pathToFileURL(input).href }, driver)
    } finally {
      await driver.close()
    }
    const zip = await JSZip.loadAsync(result.docx)
    const xml = (file: string) => zip.file(file)?.async('string')
    return { zip, xml, ir: result.ir, screenshotText: result.screenshotText }
  }
}
