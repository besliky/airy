/// The outline +/- gutter: a DOM overlay pinned over the row- and
/// column-header strips, showing one collapse/expand button per outline
/// group (level lanes 1-7, deeper levels further from the header numbers).
/// Clicks route through the shared outline toggle (hidden pipeline +
/// collapsed flag + visual undo). The overlay repositions on scroll, zoom,
/// and any command (row sizes, sheet switches) and hides itself whenever
/// the active sheet has no outline.

import { ICommandService } from '@univerjs/core'
import { IRenderManagerService, SHEET_VIEWPORT_KEY } from '@univerjs/engine-render'

import { computeOutlineGroups, maxOutlineLevel } from './outline'
import { lazySheetScreenExtent, type LazyWorkbookState, type UniverRuntime } from './univer-state'
import { sheetOutline } from './univer-sync'

export interface OutlineGutterHandle {
  refresh(): void
  dispose(): void
}

export interface OutlineGutterHost {
  /// The sheet's summary placement (outlinePr over Excel's defaults).
  placement: () => { summaryBelow: boolean; summaryRight: boolean }
  /// Executes one collapse/expand (the data-tools toggle).
  onToggle: (
    axis: 'rows' | 'cols',
    detail: { start: number; end: number },
    summary: number,
    collapse: boolean,
  ) => void
}

/// Buttons are 12 px squares in 14 px lanes; a gutter with the maximum
/// seven levels needs ~100 px of header strip.
const BUTTON_PX = 12
const LANE_PX = 14
const MAX_BUTTONS_PER_AXIS = 300

type OutlineEntries = ReadonlyMap<number, { level: number; collapsed: boolean }>

export function installOutlineGutter(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  host: OutlineGutterHost,
): OutlineGutterHandle | null {
  if (!runtime.univerAPI.getActiveWorkbook()) return null
  const rowsLayer = document.createElement('div')
  rowsLayer.className = 'outline-gutter outline-gutter-rows'
  const colsLayer = document.createElement('div')
  colsLayer.className = 'outline-gutter outline-gutter-cols'
  document.body.append(rowsLayer, colsLayer)

  const buttons = new Map<string, HTMLButtonElement>()
  let disposed = false
  let raf = 0

  const schedule = (): void => {
    if (disposed || raf !== 0) return
    raf = requestAnimationFrame(() => {
      raf = 0
      try {
        update()
      } catch {
        hide()
      }
    })
  }

  const hide = (): void => {
    rowsLayer.style.display = 'none'
    colsLayer.style.display = 'none'
    for (const button of buttons.values()) button.remove()
    buttons.clear()
  }

  const update = (): void => {
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const worksheet = workbook?.getActiveSheet()
    const state = lazyWorkbookRef.current
    if (!workbook || !worksheet || !state) {
      hide()
      return
    }
    const injector = (
      runtime.univer as unknown as {
        __getInjector(): { get<T>(token: unknown): T }
      }
    ).__getInjector()
    const render = injector
      .get<RenderServiceLike>(IRenderManagerService)
      .getRenderById(workbook.getId())
    const viewport = render?.scene?.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN)
    const surface = mainSurfaceRect()
    if (!viewport || !surface) {
      hide()
      return
    }
    const sheetId = worksheet.getSheetId()
    const outline = sheetOutline(state, sheetId)
    const extent = lazySheetScreenExtent(state, sheetId)
    const placement = host.placement()
    const zoom = worksheet.getZoom() || 1

    const cellOrigin = (row: number, column: number): { x: number; y: number } | null => {
      try {
        const rect = worksheet.getRange(row, column, 1, 1).getCellRect()
        if (!rect) return null
        return {
          x: surface.x + (rect.x - viewport.viewportScrollX) * zoom,
          y: surface.y + (rect.y - viewport.viewportScrollY) * zoom,
        }
      } catch {
        return null
      }
    }

    const rowStrip = headerStrip(surface, 'rows')
    const colStrip = headerStrip(surface, 'cols')
    const wanted = new Set<string>()

    const place = (
      axis: 'rows' | 'cols',
      layer: HTMLDivElement,
      entries: OutlineEntries,
      summaryAfter: boolean,
      extentLines: number,
      strip: DOMRect | null,
    ): boolean => {
      if (!strip || maxOutlineLevel(entries) === 0) return false
      const groups = computeOutlineGroups(entries, axis, summaryAfter, extentLines - 1)
      let placed = 0
      for (const group of groups) {
        const origin = cellOrigin(
          axis === 'rows' ? group.summary : 0,
          axis === 'rows' ? 0 : group.summary,
        )
        if (!origin) continue
        const size =
          axis === 'rows'
            ? worksheet.getRowHeight(group.summary) * zoom
            : worksheet.getColumnWidth(group.summary) * zoom
        const leading = axis === 'rows' ? origin.y : origin.x
        const stripStart = axis === 'rows' ? surface.y : surface.x
        const stripEnd = axis === 'rows' ? surface.y + surface.height : surface.x + surface.width
        // Summary lines fully off-screen draw nothing.
        if (leading + size < stripStart || leading > stripEnd) continue
        if (placed >= MAX_BUTTONS_PER_AXIS) break
        placed += 1
        const key = `${axis}:${group.summary}:${group.level}`
        wanted.add(key)
        let button = buttons.get(key)
        if (!button) {
          button = document.createElement('button')
          button.type = 'button'
          // Resolve the group fresh at click time: row/column edits may
          // have moved its span since the button was placed.
          button.addEventListener('click', () => {
            const current = computeOutlineGroups(entries, axis, summaryAfter, extentLines - 1).find(
              (candidate) => candidate.summary === group.summary,
            )
            if (!current) return
            host.onToggle(
              axis,
              { start: current.start, end: current.end },
              current.summary,
              !(entries.get(current.summary)?.collapsed ?? false),
            )
            schedule()
          })
          layer.append(button)
          buttons.set(key, button)
        }
        const collapsed = entries.get(group.summary)?.collapsed ?? false
        button.textContent = collapsed ? '+' : '−'
        button.className = 'outline-gutter-button'
        if (axis === 'rows') {
          button.style.left = `${surface.x - 3 - group.level * LANE_PX}px`
          button.style.top = `${leading + Math.max((size - BUTTON_PX) / 2, 0)}px`
        } else {
          button.style.top = `${surface.y - 3 - group.level * LANE_PX}px`
          button.style.left = `${leading + Math.max((size - BUTTON_PX) / 2, 0)}px`
        }
      }
      return placed > 0
    }

    const rowsShown = place(
      'rows',
      rowsLayer,
      outline.rows,
      placement.summaryBelow,
      Math.max(extent?.rows ?? 0, worksheet.getMaxRows()),
      rowStrip,
    )
    const colsShown = place(
      'cols',
      colsLayer,
      outline.cols,
      placement.summaryRight,
      Math.max(extent?.columns ?? 0, worksheet.getMaxColumns()),
      colStrip,
    )
    rowsLayer.style.display = rowsShown ? 'block' : 'none'
    colsLayer.style.display = colsShown ? 'block' : 'none'

    for (const [key, button] of buttons) {
      if (wanted.has(key)) continue
      button.remove()
      buttons.delete(key)
    }
  }

  const disposers: Array<() => void> = []
  try {
    const injector = (
      runtime.univer as unknown as {
        __getInjector(): { get<T>(token: unknown): T }
      }
    ).__getInjector()
    const render = injector
      .get<RenderServiceLike>(IRenderManagerService)
      .getRenderById(runtime.univerAPI.getActiveWorkbook()?.getId() ?? '')
    const viewport = render?.scene?.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN)
    if (viewport) {
      const subscription = viewport.onScrollAfter$.subscribe(() => schedule())
      disposers.push(() => subscription.unsubscribe())
    }
    const commandSubscription = injector
      .get<{ onCommandExecuted(cb: () => void): { dispose(): void } }>(ICommandService)
      .onCommandExecuted(() => schedule())
    disposers.push(() => commandSubscription.dispose())
  } catch {
    // Without the observers the gutter stays at its last position; the
    // explicit refresh() calls from outline edits still reposition it.
  }
  const onResize = (): void => schedule()
  window.addEventListener('resize', onResize)
  disposers.push(() => window.removeEventListener('resize', onResize))

  schedule()
  return {
    refresh: schedule,
    dispose(): void {
      disposed = true
      if (raf !== 0) cancelAnimationFrame(raf)
      for (const dispose of disposers) dispose()
      hide()
      rowsLayer.remove()
      colsLayer.remove()
    },
  }
}

/// The slice of the render manager the gutter reads.
interface RenderServiceLike {
  getRenderById(id: string):
    | {
        scene?:
          | {
              getViewport(key: string): {
                viewportScrollX: number
                viewportScrollY: number
                onScrollAfter$: { subscribe(cb: () => void): { unsubscribe(): void } }
              } | null
            }
          | undefined
      }
    | undefined
}

/// The main grid surface: the largest canvas in the Univer container (the
/// container holds extra chrome above the grid).
function mainSurfaceRect(): DOMRect | null {
  const host = document.getElementById('univer-container')
  if (!host) return null
  let surface: DOMRect | null = null
  for (const canvas of host.querySelectorAll('canvas')) {
    const rect = canvas.getBoundingClientRect()
    if (!surface || rect.width * rect.height > surface.width * surface.height) surface = rect
  }
  return surface
}

/// The row/column header strip: the canvas sitting immediately left of /
/// above the main surface with a strip-like aspect. Its presence decides
/// whether the gutter can draw at all.
function headerStrip(surface: DOMRect, axis: 'rows' | 'cols'): DOMRect | null {
  const host = document.getElementById('univer-container')
  if (!host) return null
  let best: DOMRect | null = null
  for (const canvas of host.querySelectorAll('canvas')) {
    const rect = canvas.getBoundingClientRect()
    if (rect.width * rect.height >= surface.width * surface.height) continue
    if (axis === 'rows') {
      if (Math.abs(rect.right - surface.left) > 2) continue
      if (rect.height < surface.height * 0.5 || rect.width > 200) continue
    } else {
      if (Math.abs(rect.bottom - surface.top) > 2) continue
      if (rect.width < surface.width * 0.5 || rect.height > 200) continue
    }
    if (!best || rect.width * rect.height > best.width * best.height) best = rect
  }
  return best
}
