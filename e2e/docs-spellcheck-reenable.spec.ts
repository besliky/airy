/**
 * Re-enabling Review → Spelling must restore red squiggles on EXISTING text
 * with no click or keystroke from the user (the earlier focus-cycle fix did
 * not actually do this). Squiggles are
 * native Chromium markers invisible to the DOM, so the assertion is
 * pixel-based: count red-ish pixels over the page. Every wait below polls a
 * terminal pixel (or DOM-attribute) condition — no fixed settle windows.
 */
import { test, expect } from '@playwright/test'
import { PNG } from 'pngjs'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

async function redCount(page: Page): Promise<number> {
  const buf = await page.locator('.doc-page').screenshot()
  const png = PNG.sync.read(buf)
  let n = 0
  for (let i = 0; i < png.data.length; i += 4) {
    const r = png.data[i]!,
      g = png.data[i + 1]!,
      b = png.data[i + 2]!
    if (r > 140 && g < 110 && b < 110) n++
  }
  return n
}

test('re-enabling spellcheck respells existing text without user input', async () => {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'spellcheck-reenable' })
  const { app, page } = launched
  try {
    await page.locator('.quick-card').first().click()
    const editor = await waitForPageWithUrl(app, 'docs/out')
    await editor.locator('.doc-page').waitFor()

    await editor.locator('.doc-page').click()
    await editor.keyboard.type('Je vais a la mison ce soir', { delay: 20 })
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('encore la mison demain matin', { delay: 20 })

    // BASELINE: squiggles appear asynchronously after word commits — poll the
    // pixel counter until the typos are really marked. The full budget
    // elapsing with no red pixel at all means the native spellchecker is
    // inactive in this environment and every assertion below would be
    // meaningless.
    const marked = await expect
      .poll(() => redCount(editor), { timeout: 20_000 })
      .toBeGreaterThan(100)
      .then(
        () => true,
        () => false,
      )
    test.skip(!marked, 'native spellchecker inactive in this environment')
    const baseline = await redCount(editor)
    const textBefore = await editor.evaluate(
      () => document.querySelector('.doc-page')?.textContent ?? '',
    )

    const spelling = editor.getByRole('button', { name: 'Spelling' })
    // the ribbon strip is a WAI-ARIA tablist — the Review entry is a role=tab
    await editor.getByRole('tab', { name: 'Review' }).click()
    await spelling.waitFor()

    // OFF: the pref routes through editorProps.attributes (setOptions applies
    // them synchronously), a fast localized signal that the toggle really
    // took — before waiting on the slower pixel outcome. The markers clear
    // asynchronously: poll to zero so the final comparison cannot inherit
    // stale pixels from the enabled phase.
    await spelling.click()
    await expect(editor.locator('.doc-page')).toHaveAttribute('spellcheck', 'false')
    await expect.poll(() => redCount(editor), { timeout: 15_000 }).toBe(0)

    // ON: re-enable via the ribbon only — no click into the text, no typing.
    // The respell chain (IPC kick → trusted space → scrub → Blink respell →
    // paint) has no DOM-visible signal; poll for the pixels to come back.
    await spelling.click()
    await expect(editor.locator('.doc-page')).toHaveAttribute('spellcheck', 'true')
    const on = await expect
      .poll(() => redCount(editor), { timeout: 30_000 })
      .toBeGreaterThan(100)
      .then(() => redCount(editor))

    const textAfter = await editor.evaluate(
      () => document.querySelector('.doc-page')?.textContent ?? '',
    )

    // off is pinned to 0 by the off-phase poll, so "off + 100" is 100 here
    expect(on).toBeGreaterThan(100) // squiggles came back…
    expect(on).toBeGreaterThanOrEqual(Math.floor(baseline * 0.8)) // …on the existing lines
    expect(textAfter).toBe(textBefore) // and the respell kick left no trace
    expect(textAfter).not.toContain('​')
    expect(textAfter).not.toContain('  ')
  } finally {
    await closeAndSaveVideo(launched, 'spellcheck-reenable')
  }
})
