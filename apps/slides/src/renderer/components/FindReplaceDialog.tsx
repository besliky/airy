/**
 * Find/replace floating panel (⌘F) — modeled on PowerPoint "Home → Find/Replace".
 * Find works on render-tree text with per-hit offsets (see find-matches.ts):
 * match case / whole word, next + previous navigation, and a canvas overlay
 * highlighting every hit on the shown slide (the active hit outlined; elements
 * whose runs cannot be boxed — vertical or WordArt-warped text — get an element
 * outline instead). Replace goes through the main-process model layer (in-run
 * matching, byte-faithful patches) under the same conditions, and on success
 * the whole RenderSlide set refreshes.
 */
import React, { useEffect, useMemo, useRef, useState } from 'react'
import type { RefObject } from 'react'
import { createPortal } from 'react-dom'
import { useI18n } from '../i18n/locale'
import type { RenderSlide } from '@airy-office/pptx-render'
import {
  buildMatches,
  matchFocusBox,
  matchStageOutline,
  matchStageRects,
  stepMatchIndex,
  type FindMatch,
} from '../find-matches'

const SCAN_DEBOUNCE_MS = 120

/** one highlighted rect on the overlay (slide px) */
interface OverlayHit {
  key: string
  x: number
  y: number
  w: number
  h: number
  active: boolean
  /** rotate about the origin — the accumulated ancestor-group chain plus the
   *  element's own rotation (see matchStageRects) */
  rotation: number
  originX: number
  originY: number
}

export function FindReplaceDialog({
  slides,
  current,
  stageRel,
  onNavigate,
  onReplaced,
  onClose,
}: {
  slides: RenderSlide[]
  /** slide the canvas is showing (drives which hits get the overlay before navigation) */
  current: number
  /** .stage-rel element the highlight overlay portals into (inherits the canvas zoom) */
  stageRel: RefObject<HTMLDivElement | null>
  onNavigate: (
    slideIndex: number,
    sourceId: string,
    focus?: { x: number; y: number; w: number; h: number },
  ) => void
  onReplaced: (slides: RenderSlide[]) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [query, setQuery] = useState('')
  const [replaceText, setReplaceText] = useState('')
  const [matchCase, setMatchCase] = useState(false)
  const [wholeWord, setWholeWord] = useState(false)
  const [cursor, setCursor] = useState(-1)
  const [status, setStatus] = useState('')
  const [scanned, setScanned] = useState<FindMatch[]>([])
  const findRef = useRef<HTMLInputElement>(null)

  useEffect(() => findRef.current?.focus(), [])

  // Debounced rescan when the deck or the conditions change (large decks stay
  // smooth while typing); the cursor is invalidated
  useEffect(() => {
    const id = window.setTimeout(
      () => setScanned(buildMatches(slides, query, { matchCase, wholeWord })),
      SCAN_DEBOUNCE_MS,
    )
    return () => window.clearTimeout(id)
  }, [slides, query, matchCase, wholeWord])
  const matches = scanned

  const findWith = (dir: 1 | -1) => {
    if (!matches.length) {
      setStatus(t('paneFrNotFound'))
      return
    }
    const next = stepMatchIndex(cursor, dir, matches.length)
    setCursor(next)
    const m = matches[next]!
    onNavigate(m.slideIndex, m.sourceId, matchFocusBox(slides[m.slideIndex], m))
    setStatus(t('paneFrMatchPos', { i: String(next + 1), n: String(matches.length) }))
  }

  const doReplace = async (all: boolean) => {
    if (!query) return
    const cur = !all && cursor >= 0 ? matches[cursor] : undefined
    if (!all && !cur) {
      findWith(1)
      return
    }
    const r = await window.slidesApi.findReplace({
      find: query,
      replace: replaceText,
      matchCase,
      wholeWord,
      ...(all ? {} : { firstOnly: true, slideIndex: cur!.slideIndex, elementId: cur!.sourceId }),
    })
    if (!r || !r.count || !r.slides) {
      setStatus(t('paneFrNotFound'))
      return
    }
    onReplaced(r.slides)
    setStatus(t('paneFrReplaced', { n: String(r.count) }))
    if (!all) setCursor(cursor - 1) // The current item is consumed; the next findNext lands on the following one
  }

  // overlay hits live on the navigated slide, or on the shown slide before navigation
  const overlaySlide = cursor >= 0 ? matches[cursor]!.slideIndex : current
  const overlay = useMemo<OverlayHit[]>(() => {
    const slide = slides[overlaySlide]
    if (!slide || !matches.length) return []
    const hits: OverlayHit[] = []
    matches.forEach((m, i) => {
      if (m.slideIndex !== overlaySlide) return
      // run rects projected through the ancestor-group rotation chain
      const rects = matchStageRects(slide, m)
      if (rects.length) {
        for (const r of rects) {
          hits.push({ key: `${i}-${r.x.toFixed(1)}-${r.y.toFixed(1)}`, ...r, active: i === cursor })
        }
        return
      }
      // layout without boxable runs (vertical / warped): outline the element,
      // rotated like the canvas draws it
      const outline = matchStageOutline(slide, m)
      if (outline) {
        hits.push({ key: `${i}-outline`, ...outline, active: i === cursor })
      }
    })
    return hits
  }, [slides, matches, overlaySlide, cursor])

  return (
    <>
      <div className="find-panel" onKeyDown={(e) => e.key === 'Escape' && onClose()}>
        <div className="find-panel-head">
          <span>{t('paneFrTitle')}</span>
          <button className="find-panel-close" onClick={onClose} data-tip="Esc" aria-label="Esc">
            ×
          </button>
        </div>
        <div className="find-panel-row">
          <input
            ref={findRef}
            placeholder={t('paneFrFind')}
            value={query}
            onChange={(e) => {
              setQuery(e.target.value)
              setCursor(-1)
              setStatus('')
            }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') findWith(e.shiftKey ? -1 : 1)
            }}
          />
        </div>
        <div className="find-panel-row">
          <input
            placeholder={t('paneFrReplaceWith')}
            value={replaceText}
            onChange={(e) => setReplaceText(e.target.value)}
          />
        </div>
        <div className="find-panel-row find-panel-opts">
          <label>
            <input
              type="checkbox"
              checked={matchCase}
              onChange={(e) => {
                setMatchCase(e.target.checked)
                setCursor(-1)
              }}
            />
            {t('paneFrMatchCase')}
          </label>
          <label>
            <input
              type="checkbox"
              checked={wholeWord}
              onChange={(e) => {
                setWholeWord(e.target.checked)
                setCursor(-1)
              }}
            />
            {t('paneFrWholeWord')}
          </label>
          <span className="find-panel-status">
            {status ||
              (query
                ? t('paneFrMatchPos', {
                    i: String(Math.max(cursor + 1, 0)),
                    n: String(matches.length),
                  })
                : '')}
          </span>
        </div>
        <div className="find-panel-row find-panel-actions">
          <button onClick={() => findWith(-1)} disabled={!query}>
            {t('paneFrFindPrev')}
          </button>
          <button onClick={() => findWith(1)} disabled={!query}>
            {t('paneFrFindNext')}
          </button>
          <button onClick={() => void doReplace(false)} disabled={!query}>
            {t('paneFrReplace')}
          </button>
          <button onClick={() => void doReplace(true)} disabled={!query || !matches.length}>
            {t('paneFrReplaceAll')}
          </button>
        </div>
      </div>
      {stageRel.current &&
        createPortal(
          <div className="find-hits" aria-hidden="true">
            {overlay.map((h) => (
              <div
                key={h.key}
                className={`find-hit${h.active ? ' active' : ''}`}
                style={{
                  left: h.x,
                  top: h.y,
                  width: h.w,
                  height: h.h,
                  ...(h.rotation
                    ? {
                        transform: `rotate(${h.rotation}deg)`,
                        transformOrigin: `${h.originX}px ${h.originY}px`,
                      }
                    : {}),
                }}
              />
            ))}
          </div>,
          stageRel.current,
        )}
    </>
  )
}
