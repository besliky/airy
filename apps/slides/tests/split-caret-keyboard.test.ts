import { describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { splitCaretKeyDown } from '../src/renderer/components/ribbon-shared'

/**
 * UX-11s2: the split-button carets (Slide Show, New Slide in Home and Insert)
 * toggled their panel from a nested `<span class="rb-caret-hit">` hit zone —
 * mouse-only, because a real <button> cannot nest inside the main-action
 * button, so the zone cannot take focus. The UX-11s log left exactly this as
 * the next-wave candidate: "Enter on the enclosing button does main-action"
 * and nothing opens the menu. The enclosing focusable button now opens the
 * caret menu on ArrowDown/ArrowUp (Enter/Space keep the main action).
 */

type FakeEvent = {
  key: string
  preventDefault: ReturnType<typeof vi.fn>
  stopPropagation: ReturnType<typeof vi.fn>
}

function fakeEvent(key: string, over: Record<string, unknown> = {}): FakeEvent {
  return { key, preventDefault: vi.fn(), stopPropagation: vi.fn(), ...over } as FakeEvent
}

describe('splitCaretKeyDown (UX-11s2)', () => {
  it.each([
    ['ArrowDown', {}],
    ['ArrowUp', {}],
    ['ArrowDown (Alt chord)', { altKey: true }],
    ['ArrowUp (Alt chord)', { altKey: true }],
  ] as const)('%s opens the menu and claims the key', (_chord, over) => {
    const open = vi.fn()
    const key = _chord.startsWith('ArrowUp') ? 'ArrowUp' : 'ArrowDown'
    const e = fakeEvent(key, over)
    splitCaretKeyDown(open)(e as never)
    expect(open).toHaveBeenCalledTimes(1)
    expect(e.preventDefault).toHaveBeenCalled()
    expect(e.stopPropagation).toHaveBeenCalled()
  })

  it.each(['Enter', ' ', 'Escape', 'ArrowLeft', 'a'] as const)(
    '%s stays the main-action key',
    (key) => {
      const open = vi.fn()
      const e = fakeEvent(key)
      splitCaretKeyDown(open)(e as never)
      expect(open).not.toHaveBeenCalled()
      expect(e.preventDefault).not.toHaveBeenCalled()
    },
  )

  it.each([
    ['Ctrl', 'ArrowDown', { ctrlKey: true }],
    ['Cmd', 'ArrowUp', { metaKey: true }],
  ] as const)('%s+arrows are left for system shortcuts', (mod, key, over) => {
    const open = vi.fn()
    splitCaretKeyDown(open)(fakeEvent(key, over) as never)
    expect(open).not.toHaveBeenCalled()
    expect(mod).toBeTruthy()
  })
})

describe('every split caret is keyboard-reachable (source contract)', () => {
  const SRC_DIR = join(__dirname, '../src/renderer/components')
  const FILES = ['RibbonHomeTab.tsx', 'RibbonInsertTab.tsx', 'Ribbon.tsx']

  /** each caret-hit span with its enclosing split-button opening tag */
  const carets = FILES.flatMap((file) => {
    const src = readFileSync(join(SRC_DIR, file), 'utf8')
    return [...src.matchAll(/className=\{`rb-caret-hit\$\{[^}]+\}`\}/g)].map((m) => {
      const start = src.lastIndexOf('<button', m.index)
      const tagEnd = src.indexOf('>', m.index)
      return { file, tag: src.slice(start, tagEnd + 1) } as const
    })
  })

  it('finds the split carets (scanner sanity, >= 3)', () => {
    expect(carets.length).toBeGreaterThanOrEqual(3)
    expect(new Set(carets.map((c) => c.file)).size).toBeGreaterThanOrEqual(2)
  })

  it.each(carets.map((c, i) => [c.file, i] as const))(
    '%s caret #%d opens on arrows',
    (file, i) => {
      expect(carets[i as number]!.tag).toContain('onKeyDown={splitCaretKeyDown(')
      expect(file).toBeTruthy()
    },
  )
})
