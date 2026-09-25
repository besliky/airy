/**
 * PERF-1727 regression guard: the Format Pane fields must keep showing the
 * selected element's real geometry — position/size derived from the node box —
 * and follow a selection switch (same pane instance re-rendered with another
 * node). The perf fix lives in the canvas click path, so this pins the consumer
 * side: the pane values the unblocked selection commit lands in.
 */
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { ShapeRenderNode } from '@airy-office/pptx-render'

import { FormatPane } from '../src/renderer/components/FormatPane'
import { LocaleProvider, setModuleLang } from '../src/renderer/i18n/locale'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { slidesApi: { onLanguageChanged: () => () => undefined } })
setModuleLang('en')

const PX_PER_CM = 96 / 2.54

let seq = 0
function shapeNode(x: number, y: number, w: number, h: number): ShapeRenderNode {
  seq += 1
  return {
    id: `n${seq}`,
    type: 'shape',
    sourceId: `s${seq}`,
    box: { x, y, w, h, rotationDeg: 0 },
    fill: { kind: 'solid', color: '#FF0000' },
  } as unknown as ShapeRenderNode
}

function renderPane(node: ShapeRenderNode): {
  container: HTMLElement
  rerender: (next: ShapeRenderNode) => void
  unmount: () => void
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  const props = {
    viewScale: 1,
    slideSizePx: { w: 1280, h: 720 },
    onTransform: vi.fn(),
    onFill: vi.fn(),
    onImageFill: vi.fn(),
    onTextAnchor: vi.fn(),
    onTextBodyProps: vi.fn(),
    onEffects: vi.fn(),
    onStroke: vi.fn(),
    onCollapse: vi.fn(),
    onPictureCrop: vi.fn(),
    onPictureCutout: vi.fn(),
    pictureCanCutout: false,
    chartData: null,
    onChartPointColor: vi.fn(),
    onAltText: vi.fn(),
  }
  const draw = (n: ShapeRenderNode) =>
    createElement(LocaleProvider, {
      initial: 'en',
      children: createElement(FormatPane, { node: n, ...props }),
    })
  act(() => root.render(draw(node)))
  return {
    container,
    rerender: (next) => act(() => root.render(draw(next))),
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

function openSizeTab(container: HTMLElement) {
  const tab = [...container.querySelectorAll('button')].find(
    (b) => b.textContent === 'Size & Properties',
  )
  if (!tab) throw new Error('Size & Properties sub-tab not found')
  act(() => tab.click())
}

function field(container: HTMLElement, label: string): string {
  for (const row of container.querySelectorAll('label.fp-prow')) {
    const span = row.querySelector('span')
    const inp = row.querySelector('input')
    if (span?.textContent === label && inp) return inp.value
  }
  throw new Error(`pane field not found: ${label}`)
}

const cm = (px: number) => `${Math.round((px / PX_PER_CM) * 100) / 100} cm`

describe('FormatPane geometry fields follow the selected node (PERF-1727)', () => {
  it('shows the selected shape position and size in cm', () => {
    const pane = renderPane(shapeNode(100, 50, 200, 80))
    try {
      openSizeTab(pane.container)
      expect(field(pane.container, 'Horizontal position')).toBe(cm(100))
      expect(field(pane.container, 'Vertical position')).toBe(cm(50))
      expect(field(pane.container, 'Width')).toBe(cm(200))
      expect(field(pane.container, 'Height')).toBe(cm(80))
    } finally {
      pane.unmount()
    }
  })

  it('a selection switch re-reads every field from the new node', () => {
    const pane = renderPane(shapeNode(100, 50, 200, 80))
    try {
      openSizeTab(pane.container)
      const second = shapeNode(333, 21, 90, 140)
      pane.rerender(second)
      expect(field(pane.container, 'Horizontal position')).toBe(cm(333))
      expect(field(pane.container, 'Vertical position')).toBe(cm(21))
      expect(field(pane.container, 'Width')).toBe(cm(90))
      expect(field(pane.container, 'Height')).toBe(cm(140))
    } finally {
      pane.unmount()
    }
  })
})
