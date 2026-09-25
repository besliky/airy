/**
 * UX-1696: the "Reopen with encoding" affordance must actually reach the
 * main process. The UX-1653 channel existed but was unreachable — no preload
 * pass-through, no renderer caller, no UI. These source pins (the repo's
 * wiring-test pattern, cf. apps/html/tests/save-staleness-wiring.test.ts)
 * keep the whole chain connected: renderer picker → MarkdownApi.setEncoding →
 * preload MARKDOWN_CHANNELS.setEncoding → main handler. The behavioral
 * channel→main-reopens side is covered by encoding-memory.test.ts against
 * the real handlers; the status-bar strings are checked for locale parity.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { SELECTABLE_ENCODINGS } from '../src/shared/ipc'
import { strings } from '../src/renderer/i18n/strings'

const here = dirname(fileURLToPath(import.meta.url))
const shared = readFileSync(join(here, '../src/shared/ipc.ts'), 'utf8')
const preload = readFileSync(join(here, '../src/preload/index.ts'), 'utf8')
const renderer = readFileSync(join(here, '../src/renderer/App.tsx'), 'utf8')
const main = readFileSync(join(here, '../src/main/markdown-main.ts'), 'utf8')

const NEW_KEYS = ['encodingAuto', 'reopenEncoding', 'openedAs', 'openedAuto', 'reopenDirty']

describe('markdown reopen-with-encoding wiring', () => {
  it('the channel lives in the pinned shared registry (preload-reachable)', () => {
    expect(shared).toContain("setEncoding: 'markdown:set-encoding'")
    expect(shared).toContain('setEncoding(path: string, encoding: string | null): Promise<boolean>')
    // the selectable list is shared, not a renderer-local copy
    expect(shared).toContain('export const SELECTABLE_ENCODINGS')
  })

  it('the preload passes the channel through', () => {
    expect(preload).toContain('setEncoding: (path, encoding) =>')
    expect(preload).toContain('ipcRenderer.invoke(MARKDOWN_CHANNELS.setEncoding, path, encoding)')
  })

  it('the renderer picker remembers the pick and re-reads the file', () => {
    expect(renderer).toContain(
      "window.markdownApi.setEncoding(path, pick === 'auto' ? null : pick)",
    )
    // the reopen goes back through the normal open path (readFile), not a side channel
    expect(renderer).toContain('const recovered = await loadDocument(path, () => false)')
    expect(renderer).toContain('reopenWithEncoding')
  })

  it('a dirty document is refused instead of silently discarded', () => {
    expect(renderer).toContain("showToast(moduleT('reopenDirty'), 'error')")
  })

  it('the main handler validates, remembers and forgets (null) through the settings queue', () => {
    expect(main).toContain('MARKDOWN_CHANNELS.setEncoding,')
    expect(main).toContain('if (encoding === null)')
    expect(main).toContain('await forgetFileEncoding(appSettingsPath(), path)')
    expect(main).toContain('await rememberFileEncoding(appSettingsPath(), path, encoding)')
  })
})

describe('markdown encoding picker truthfulness (BUG-1741)', () => {
  it('the read-only getEncoding channel is registered, wired and pick-aware', () => {
    expect(shared).toContain("getEncoding: 'markdown:get-encoding'")
    expect(shared).toContain('getEncoding(path: string): Promise<string | null>')
    expect(preload).toContain('getEncoding: (path) =>')
    expect(preload).toContain('ipcRenderer.invoke(MARKDOWN_CHANNELS.getEncoding, path)')
    expect(main).toContain('MARKDOWN_CHANNELS.getEncoding,')
    expect(main).toContain('readRememberedFileEncoding(appSettingsPath(), path) ?? null')
  })

  it('the picker mirrors the persisted pick at open and after every save', () => {
    // open: the persisted pick governs the decode just used — show it
    expect(renderer).toContain('window.markdownApi.getEncoding(path).catch(() => null)')
    // save: a Save As onto a fresh path has no pick, a fallback UTF-8 write
    // dropped the old one — re-read the truth for the resolved path
    expect(renderer).toContain('window.markdownApi.getEncoding(result.path).catch(() => null)')
    expect(renderer).toContain('setEncodingPick(asEncodingPick(remembered))')
  })

  it('the save encodes into the remembered charset instead of unconditional UTF-8', () => {
    expect(main).toContain('encodeForSave(textToWrite, target)')
    expect(main).toContain('readRememberedFileEncoding(appSettingsPath(), target)')
  })
})

describe('markdown reopen-with-encoding i18n', () => {
  const locales = Object.keys(strings) as Array<keyof typeof strings>

  it.each(locales)('locale %s carries the status-bar encoding strings', (locale) => {
    const table = strings[locale] as Record<string, unknown>
    for (const key of NEW_KEYS) {
      expect(typeof table[key]).toBe('string')
      expect((table[key] as string).trim().length).toBeGreaterThan(0)
    }
    // the success toast names the charset via the {encoding} placeholder
    expect(table.openedAs).toContain('{encoding}')
  })

  it('the picker offers auto-detect plus exactly the selectable charsets', () => {
    // 'auto' is the no-override value; every charset mirrors LEGACY_CHARSETS
    // + the UTF family (synchronization comment on SELECTABLE_ENCODINGS)
    expect(SELECTABLE_ENCODINGS.length).toBe(15)
    expect(SELECTABLE_ENCODINGS).toContain('windows-1251')
    expect(SELECTABLE_ENCODINGS).toContain('gb18030')
  })
})
