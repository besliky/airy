/**
 * UX-1696: the "Reopen with encoding" affordance must actually reach the
 * main process. The UX-1653 channel existed but was unreachable — no preload
 * pass-through, no renderer caller, no UI. These source pins (this app's
 * wiring-test pattern, cf. save-staleness-wiring.test.ts) keep the whole
 * chain connected: renderer picker → HtmlApi.setEncoding → preload
 * HTML_CHANNELS.setEncoding → main handler; the channel→main-reopens
 * mechanics are covered in encoding-memory.test.ts, and the status-bar
 * strings are checked for locale parity across the app shards.
 */
import { readFileSync, readdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { SELECTABLE_ENCODINGS } from '../src/shared/ipc'
import { appStrings } from '../src/renderer/i18n/strings-app'

const here = dirname(fileURLToPath(import.meta.url))
const shared = readFileSync(join(here, '../src/shared/ipc.ts'), 'utf8')
const preload = readFileSync(join(here, '../src/preload/index.ts'), 'utf8')
const renderer = readFileSync(join(here, '../src/renderer/App.tsx'), 'utf8')
const main = readFileSync(join(here, '../src/main/html-main.ts'), 'utf8')

const NEW_KEYS = ['encodingAuto', 'reopenEncoding', 'openedAs', 'openedAuto', 'reopenDirty']

describe('html reopen-with-encoding wiring', () => {
  it('the channel lives in the pinned shared registry (preload-reachable)', () => {
    expect(shared).toContain("setEncoding: 'html:set-encoding'")
    expect(shared).toContain('setEncoding(path: string, encoding: string | null): Promise<boolean>')
    // the selectable list is shared, not a renderer-local copy
    expect(shared).toContain('export const SELECTABLE_ENCODINGS')
  })

  it('the preload passes the channel through', () => {
    expect(preload).toContain('setEncoding: (path, encoding) =>')
    expect(preload).toContain('ipcRenderer.invoke(HTML_CHANNELS.setEncoding, path, encoding)')
  })

  it('the renderer picker remembers the pick and re-reads the file', () => {
    expect(renderer).toContain("window.htmlApi.setEncoding(current, pick === 'auto' ? null : pick)")
    // the reopen goes back through the normal open path (readFile), not a side channel
    expect(renderer).toContain('const opened = await window.htmlApi.readFile(current)')
    expect(renderer).toContain('reopenWithEncoding')
    // the source editor and the preview are both swapped to the new decoding
    expect(renderer).toContain('editorRef.current?.setDoc(doc.text)')
  })

  it('a dirty document is refused instead of silently discarded', () => {
    expect(renderer).toContain("setNotice(moduleT('reopenDirty'))")
  })

  it('the main handler validates, remembers and forgets (null) through the settings queue', () => {
    expect(main).toContain('HTML_CHANNELS.setEncoding,')
    expect(main).toContain('if (encoding === null)')
    expect(main).toContain('await forgetFileEncoding(appSettingsPath(), path)')
    expect(main).toContain('await rememberFileEncoding(appSettingsPath(), path, encoding)')
  })
})

describe('html reopen-with-encoding i18n', () => {
  const locales = Object.keys(appStrings) as Array<keyof typeof appStrings>

  it.each(locales)('locale %s carries the status-bar encoding strings', (locale) => {
    const table = appStrings[locale] as Record<string, unknown>
    for (const key of NEW_KEYS) {
      expect(typeof table[key]).toBe('string')
      expect((table[key] as string).trim().length).toBeGreaterThan(0)
    }
    // the success notice names the charset via the {encoding} placeholder
    expect(table.openedAs).toContain('{encoding}')
  })

  it('every app shard file on disk defines the new keys (no shard left behind)', () => {
    const shards = readdirSync(join(here, '../src/renderer/i18n/app')).filter((name) =>
      name.endsWith('.ts'),
    )
    expect(shards.length).toBe(20)
    for (const shard of shards) {
      const source = readFileSync(join(here, '../src/renderer/i18n/app', shard), 'utf8')
      for (const key of NEW_KEYS) expect(source).toContain(`${key}:`)
    }
  })

  it('the picker offers auto-detect plus exactly the selectable charsets', () => {
    // 'auto' is the no-override value; every charset mirrors LEGACY_CHARSETS
    // + the UTF family (synchronization comment on SELECTABLE_ENCODINGS)
    expect(SELECTABLE_ENCODINGS.length).toBe(15)
    expect(SELECTABLE_ENCODINGS).toContain('windows-1251')
    expect(SELECTABLE_ENCODINGS).toContain('gb18030')
  })
})
