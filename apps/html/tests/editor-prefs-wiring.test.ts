/**
 * UX-1704: the split-view scroll-sync and the word-wrap toggle must actually
 * reach every hop. This app's wiring-test pattern (cf.
 * encoding-ui-wiring.test.ts) pins the chains by source assertions:
 * - prefs: Ribbon toggle → App state → HtmlApi.setEditorPrefs → preload
 *   channel → main handler → app-settings.json single-writer queue, plus the
 *   hydration on open;
 * - scroll-sync: CM scroll event → rAF coalescing + echo gate → gx:scrollTo
 *   into the frame, and the frame's gx:scrolled back into applyScrollRatio;
 * - word-wrap: the Ribbon toggle → the lineWrapping compartment (the
 *   compartment mechanics are behavior-tested in word-wrap.test.ts);
 * and the status-bar… no — toolbar strings are checked for locale parity
 * across all 20 app shards.
 */
import { readFileSync, readdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { HTML_CHANNELS } from '../src/shared/ipc'
import { appStrings } from '../src/renderer/i18n/strings-app'

const here = dirname(fileURLToPath(import.meta.url))
const shared = readFileSync(join(here, '../src/shared/ipc.ts'), 'utf8')
const preload = readFileSync(join(here, '../src/preload/index.ts'), 'utf8')
const renderer = readFileSync(join(here, '../src/renderer/App.tsx'), 'utf8')
const main = readFileSync(join(here, '../src/main/html-main.ts'), 'utf8')
const cmSetup = readFileSync(join(here, '../src/renderer/source/cm-setup.ts'), 'utf8')
const sourceEditor = readFileSync(join(here, '../src/renderer/source/SourceEditor.tsx'), 'utf8')
const ribbon = readFileSync(join(here, '../src/renderer/components/Ribbon.tsx'), 'utf8')
const inspector = readFileSync(join(here, '../src/renderer/preview/inspector.js'), 'utf8')
const protocol = readFileSync(join(here, '../src/renderer/preview/inspector-protocol.ts'), 'utf8')

const NEW_KEYS = ['wordWrap', 'wordWrapTip', 'scrollSync', 'scrollSyncTip']

describe('editor prefs channel wiring', () => {
  it('the channels live in the pinned shared registry (preload-reachable)', () => {
    expect(HTML_CHANNELS.getEditorPrefs).toBe('html:get-editor-prefs')
    expect(HTML_CHANNELS.setEditorPrefs).toBe('html:set-editor-prefs')
    expect(shared).toContain('getEditorPrefs(): Promise<EditorPrefs>')
    expect(shared).toContain('setEditorPrefs(patch: EditorPrefsPatch): Promise<EditorPrefs>')
    // defaults are shared, not a main-side or renderer-side copy
    expect(shared).toContain(
      'DEFAULT_EDITOR_PREFS: EditorPrefs = { wordWrap: true, scrollSync: false }',
    )
  })

  it('the preload passes the channels through', () => {
    expect(preload).toContain(
      'getEditorPrefs: () => ipcRenderer.invoke(HTML_CHANNELS.getEditorPrefs)',
    )
    expect(preload).toContain(
      'setEditorPrefs: (patch) => ipcRenderer.invoke(HTML_CHANNELS.setEditorPrefs, patch)',
    )
  })

  it('the renderer hydrates the prefs before the document opens', () => {
    expect(renderer).toContain("typeof window.htmlApi.getEditorPrefs === 'function'")
    expect(renderer).toContain('setWordWrapState(prefs.wordWrap)')
    expect(renderer).toContain('setScrollSyncState(prefs.scrollSync)')
    // toggles must not fire before hydration (defaults would clobber the store)
    expect(renderer).toContain('prefsLoadedRef.current = true')
    expect(renderer).toContain('if (prefsLoadedRef.current && typeof window.htmlApi.setEditorPrefs')
  })

  it('both ribbon toggles persist through the settings queue', () => {
    expect(renderer).toContain('window.htmlApi.setEditorPrefs({ wordWrap: on })')
    expect(renderer).toContain('window.htmlApi.setEditorPrefs({ scrollSync: on })')
  })

  it('the main handler reads and writes through the shared settings file', () => {
    expect(main).toContain('HTML_CHANNELS.getEditorPrefs')
    expect(main).toContain('HTML_CHANNELS.setEditorPrefs')
    expect(main).toContain('readEditorPrefs(appSettingsPath())')
    expect(main).toContain('writeEditorPrefs(appSettingsPath()')
    // a wholesale patch is rejected instead of merged blindly
    expect(main).toContain('patch must be an object')
  })
})

describe('scroll-sync wiring', () => {
  it('the source pane forwards its scroll metrics', () => {
    expect(sourceEditor).toContain("scrollDOM.addEventListener('scroll', onScrollEvent)")
    expect(renderer).toContain('onScroll={onEditorScroll}')
    expect(renderer).toContain('const onEditorScroll = useCallback')
  })

  it('the forward direction coalesces, gates and drives the preview', () => {
    // scroll bursts collapse into one frame
    expect(renderer).toContain('window.requestAnimationFrame')
    expect(renderer).toContain('syncGateRef.current.shouldIgnore')
    expect(renderer).toContain("syncGateRef.current.drive('preview', Date.now())")
    expect(renderer).toContain("previewRef.current?.post({ type: 'gx:scrollTo', ratio })")
    // sub-rounding moves are dropped: no feedback hum
    expect(renderer).toContain('SYNC_RATIO_EPSILON')
  })

  it('the reverse direction drives the source pane, gated the same way', () => {
    expect(renderer).toContain("case 'gx:scrolled':")
    expect(renderer).toContain("syncGateRef.current.drive('source', Date.now())")
    expect(renderer).toContain('editorRef.current?.applyScrollRatio(msg.ratio)')
  })

  it('the protocol declares both directions', () => {
    expect(protocol).toContain("'gx:scrollTo'")
    expect(protocol).toContain("'gx:scrolled'")
  })

  it('the inspector reports the viewport and obeys host repositions', () => {
    expect(inspector).toContain("case 'gx:scrollTo'")
    expect(inspector).toContain("post({ type: 'gx:scrolled', ratio: viewportRatio() })")
    expect(inspector).toContain('window.scrollTo(0, ratio * max)')
    // its own programmatic reposition must not echo back as a user scroll
    expect(inspector).toContain('SYNC_ECHO_MS')
  })

  it('a reload of the frame re-applies the last known source position', () => {
    expect(renderer).toContain(
      "previewRef.current?.post({ type: 'gx:scrollTo', ratio: lastSyncRatioRef.current })",
    )
  })

  it('the sync only runs in split view, where both panes are on screen', () => {
    expect(renderer).toContain("const syncActive = (view: ViewMode) => view === 'split'")
    expect(renderer).toContain('!syncActive(viewRef.current)')
  })
})

describe('word-wrap wiring', () => {
  it('the wrap extension sits in a compartment, not as a static extension', () => {
    expect(cmSetup).toContain('export const wrapCompartment = new Compartment()')
    expect(cmSetup).toContain('wrapCompartment.of(wrap ? EditorView.lineWrapping : [])')
    // the old always-on static line is gone
    expect(cmSetup).not.toMatch(/^\s*EditorView\.lineWrapping,\s*$/m)
  })

  it('the editor applies the toggle in place', () => {
    expect(sourceEditor).toContain('setLineWrap(viewRef.current, wordWrap)')
    // the App passes the hydrated, persisted state down
    expect(renderer).toContain('wordWrap={wordWrap}')
  })

  it('the ribbon exposes both toggles as pressed-state buttons', () => {
    expect(ribbon).toContain('aria-pressed={p.wordWrap}')
    expect(ribbon).toContain('aria-pressed={p.scrollSync}')
    expect(ribbon).toContain('onToggleWordWrap(!p.wordWrap)')
    expect(ribbon).toContain('onToggleScrollSync(!p.scrollSync)')
    // scroll-sync only means something in split view; wrap only where the source pane shows
    expect(ribbon).toContain("disabled={off || p.view !== 'split'}")
    expect(ribbon).toContain("disabled={off || p.view === 'preview'}")
  })
})

describe('editor toggle i18n', () => {
  const locales = Object.keys(appStrings) as Array<keyof typeof appStrings>

  it.each(locales)('locale %s carries the toolbar toggle strings', (locale) => {
    const table = appStrings[locale] as Record<string, unknown>
    for (const key of NEW_KEYS) {
      expect(typeof table[key]).toBe('string')
      expect((table[key] as string).trim().length).toBeGreaterThan(0)
    }
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
})
