import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { LocaleProvider, setModuleLang, t } from '../src/renderer/i18n/locale'
import { ExportVideoDialog } from '../src/renderer/components/ExportVideoDialog'
import type { VideoExportPhase, VideoExportSettings } from '../src/renderer/file-actions'

/**
 * UX-1201 focus retention: starting the export used to unmount the Export
 * button, dropping focus to <body> for the whole (minutes-long, real-time)
 * recording — the Tab trap only sees keys bubbling through the backdrop, so
 * the "modal" was keyboard-transparent behind its own aria-modal. The buttons
 * now stay mounted (Export disables itself), focus is handed to Cancel when
 * the run starts, and a focus sentinel pulls any escaping focus back in.
 *
 * UX-1202 cancel acknowledgement: the mid-run Cancel button flags the
 * cooperative box (honored by BOTH the render and record loops) and shows a
 * disabled "Cancelling…" state until the pipeline unwinds and the dialog
 * closes — no dead button on a minutes-long run.
 */
beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { slidesApi: { onLanguageChanged: () => () => undefined } })

setModuleLang('en')

/** The dialog probes MediaRecorder for a container before enabling Export. */
class FakeMediaRecorder {
  static isTypeSupported(): boolean {
    return true
  }
}
;(globalThis as Record<string, unknown>).MediaRecorder = FakeMediaRecorder

type Progress = (phase: VideoExportPhase, done: number, total: number) => void
type ExportFn = (
  settings: VideoExportSettings,
  onProgress: Progress,
  cancel: { current: boolean },
) => Promise<boolean>

function renderDialog(onExport: ExportFn): {
  container: HTMLElement
  onClose: ReturnType<typeof vi.fn>
  unmount: () => void
} {
  const onClose = vi.fn()
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() =>
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(ExportVideoDialog, {
          slides: [{}, {}],
          advanceMs: [null, null],
          transitions: [
            { kind: 'none', durationMs: null },
            { kind: 'none', durationMs: null },
          ],
          onExport,
          onClose,
        }),
      }),
    ),
  )
  return {
    container,
    onClose,
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

/** Export control that never resolves (the run is "under way"). */
function pendingExport(): ExportFn {
  return () => new Promise<boolean>(() => undefined)
}

/** The mid-run buttons: [Cancel, Export]. */
function runButtons(container: HTMLElement): [HTMLButtonElement, HTMLButtonElement] {
  const buttons = [...container.querySelectorAll<HTMLButtonElement>('.modal-actions button')]
  return [buttons[0]!, buttons[1]!]
}

describe('ExportVideoDialog focus retention (UX-1201)', () => {
  it('keeps the Export button mounted (disabled) and hands focus to Cancel', () => {
    const { container, unmount } = renderDialog(pendingExport())
    const [cancelBtn, exportBtn] = [
      ...container.querySelectorAll<HTMLButtonElement>('.modal-actions button'),
    ]
    exportBtn.focus()
    expect(document.activeElement).toBe(exportBtn)
    act(() => exportBtn.click())
    expect(exportBtn.disabled).toBe(true) // disabled, not unmounted
    expect(container.contains(exportBtn)).toBe(true)
    expect(document.activeElement).toBe(cancelBtn) // not <body>
    unmount()
  })

  it('pulls focus back into the dialog when it escapes mid-run (sentinel)', () => {
    const { container, unmount } = renderDialog(pendingExport())
    const exportBtn = container.querySelectorAll<HTMLButtonElement>(
      '.modal-actions button.primary',
    )[0]!
    act(() => exportBtn.click())
    const cancelBtn = container.querySelectorAll<HTMLButtonElement>('.modal-actions button')[0]!
    expect(document.activeElement).toBe(cancelBtn)
    act(() => {
      document.body.focus() // focus fell out (element removed / clicked away)
    })
    expect(document.activeElement).toBe(cancelBtn)
    unmount()
  })

  it('wraps Tab inside the dialog while exporting (trap stays honest)', () => {
    const { container, unmount } = renderDialog(pendingExport())
    const exportBtn = container.querySelectorAll<HTMLButtonElement>(
      '.modal-actions button.primary',
    )[0]!
    act(() => exportBtn.click())
    const cancelBtn = container.querySelectorAll<HTMLButtonElement>('.modal-actions button')[0]!
    expect(document.activeElement).toBe(cancelBtn)
    // Cancel is the only focusable control mid-run: Tab wraps to itself
    const tab = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true })
    act(() => {
      cancelBtn.dispatchEvent(tab)
    })
    expect(tab.defaultPrevented).toBe(true)
    expect(document.activeElement).toBe(cancelBtn)
    unmount()
  })

  it('labels the mid-run Cancel control in the dialog', () => {
    const { container, unmount } = renderDialog(pendingExport())
    const [, exportBtn] = runButtons(container)
    act(() => exportBtn.click())
    const [cancelBtn] = runButtons(container)
    expect(cancelBtn.textContent).toBe(t('appSettingsCancel'))
    expect(cancelBtn.disabled).toBe(false) // cancelling stays available
    unmount()
  })
})

describe('ExportVideoDialog cancel acknowledgement (UX-1202)', () => {
  it('flags the cooperative box, disables itself and relabels to Cancelling…', async () => {
    let box: { current: boolean } | undefined
    const onExport: ExportFn = (_settings, _onProgress, cancel) => {
      box = cancel
      return new Promise<boolean>(() => undefined)
    }
    const { container, unmount } = renderDialog(onExport)
    const [, exportBtn] = runButtons(container)
    act(() => exportBtn.click())
    const [cancelBtn] = runButtons(container)
    act(() => cancelBtn.click())
    expect(box?.current).toBe(true) // the render loop reads this between slides
    expect(cancelBtn.disabled).toBe(true)
    expect(cancelBtn.textContent).toBe(t('appExportVideoCancelling'))
    unmount()
  })

  it('closes the dialog once the cancelled run unwinds', async () => {
    let release: (ok: boolean) => void = () => undefined
    const onExport: ExportFn = () =>
      new Promise<boolean>((resolve) => {
        release = resolve
      })
    const { container, onClose, unmount } = renderDialog(onExport)
    const [, exportBtn] = runButtons(container)
    act(() => exportBtn.click())
    const [cancelBtn] = runButtons(container)
    act(() => cancelBtn.click())
    expect(onClose).not.toHaveBeenCalled()
    await act(async () => {
      release(false) // cancelled pipelines report failure — cancel closes anyway
    })
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
  })
})
