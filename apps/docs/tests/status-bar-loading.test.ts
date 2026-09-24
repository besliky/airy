import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { LANGS } from '@airy-office/i18n'
import { LocaleProvider, loadLocale, setModuleLang } from '../src/renderer/i18n/locale'
import { loadStrings } from '../src/renderer/i18n/strings'
import { StatusBarCounters } from '../src/renderer/App'

/**
 * UX-1612: while a phased open (PERF-1502/1639) streams the document tail the
 * status-bar counters grow chunk by chunk — "Page 1 of 4, 832 words" becomes
 * "Page 2 of 5, 1200 words" — which can read like a truncated file. The bar
 * must show a "Loading…" badge while the tail streams, drop it once it lands,
 * and leave the counters themselves untouched in both states.
 */

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // PERF-904: the en dictionary these assertions read must load first
  return loadLocale('en')
})

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })

setModuleLang('en')

function renderCounters(loading: boolean): { container: HTMLElement; root: Root } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() =>
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(StatusBarCounters, {
          current: 1,
          total: 4,
          wordCount: 832,
          loading,
          onOpenStats: vi.fn(),
        }),
      }),
    ),
  )
  return { container, root }
}

function counters(container: HTMLElement): { page: string; words: string } {
  const page = container.querySelector('.status-item')?.textContent ?? ''
  const words = container.querySelector('.status-wordcount')?.textContent ?? ''
  return { page, words }
}

describe('status bar loading badge during a phased open (UX-1612)', () => {
  it('shows a Loading… badge next to the counters while the tail streams', () => {
    const { container, root } = renderCounters(true)
    const badge = container.querySelector('.status-loading')
    expect(badge).not.toBeNull()
    expect(badge!.getAttribute('role')).toBe('status')
    expect(badge!.textContent).toBe('Loading…')
    // the counters themselves are intact while loading
    expect(counters(container)).toEqual({ page: 'Page 1 of 4', words: '832 words' })
    act(() => root.unmount())
    container.remove()
  })

  it('drops the badge once the phased open has finished', () => {
    const { container, root } = renderCounters(false)
    expect(container.querySelector('.status-loading')).toBeNull()
    // and the counters still render on their own
    expect(counters(container)).toEqual({ page: 'Page 1 of 4', words: '832 words' })
    act(() => root.unmount())
    container.remove()
  })

  it('removes the badge when loading settles on a re-render', () => {
    const container = document.createElement('div')
    document.body.appendChild(container)
    const root = createRoot(container)
    const render = (loading: boolean) =>
      act(() =>
        root.render(
          createElement(LocaleProvider, {
            initial: 'en',
            children: createElement(StatusBarCounters, {
              current: 2,
              total: 5,
              wordCount: 1200,
              loading,
              onOpenStats: vi.fn(),
            }),
          }),
        ),
      )
    render(true)
    expect(container.querySelector('.status-loading')).not.toBeNull()
    render(false)
    expect(container.querySelector('.status-loading')).toBeNull()
    // the counters keep their latest streamed values after completion
    expect(counters(container)).toEqual({ page: 'Page 2 of 5', words: '1200 words' })
    act(() => root.unmount())
    container.remove()
  })

  it('defines appDocLoading in every locale dictionary', async () => {
    for (const lang of LANGS) {
      const dict = await loadStrings(lang)
      expect(typeof dict.appDocLoading, `${lang} is missing appDocLoading`).toBe('string')
      expect(dict.appDocLoading.length, `${lang} has an empty appDocLoading`).toBeGreaterThan(0)
    }
  })
})
