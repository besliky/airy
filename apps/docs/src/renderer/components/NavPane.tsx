import { memo, useEffect, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { collectHeadings } from '../editor/headings'
import { useI18n } from '../i18n/locale'
import { filterNavHeadings, splitNavLabel } from './nav-filter'

/** Word's navigation-pane search debounce: filter shortly after the last keystroke. */
const NAV_SEARCH_DEBOUNCE_MS = 150

/**
 * Word's navigation pane: heading outline with click-to-jump plus a
 * live-filtering search box (case-insensitive, Escape clears). Memoized on
 * the doc — caret moves skip the outline walk.
 */
export const NavPane = memo(function NavPane({ editor, doc }: { editor: Editor; doc: PmNode }) {
  const { t } = useI18n()
  const headings = collectHeadings(doc)
  const [query, setQuery] = useState('')
  const [debounced, setDebounced] = useState('')

  useEffect(() => {
    const id = setTimeout(() => setDebounced(query), NAV_SEARCH_DEBOUNCE_MS)
    return () => clearTimeout(id)
  }, [query])

  const visible = filterNavHeadings(headings, debounced)

  return (
    <aside className="nav-pane">
      <div className="nav-pane-title">{t('appNavTitle')}</div>
      <input
        type="search"
        className="nav-search"
        placeholder={t('appNavSearch')}
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Escape' && query) {
            e.stopPropagation()
            setQuery('')
            setDebounced('')
          }
        }}
      />
      <div className="nav-pane-list">
        {visible.map((h, i) => (
          <button
            key={`${h.pos}-${i}`}
            className={`nav-item nav-l${Math.min(h.level, 4)}`}
            data-tip={h.text}
            onClick={() => {
              const dom = editor.view.nodeDOM(h.pos) as HTMLElement | null
              dom?.scrollIntoView({ behavior: 'smooth', block: 'start' })
            }}
          >
            {splitNavLabel(h.text, debounced).map((part, j) =>
              part.hit ? (
                <span key={j} className="nav-hit">
                  {part.text}
                </span>
              ) : (
                part.text
              ),
            )}
          </button>
        ))}
        {visible.length === 0 && (
          <div className="nav-empty">
            {headings.length === 0 ? t('appNavNoHeadings') : t('appNavNoMatch')}
          </div>
        )}
      </div>
    </aside>
  )
})
