/**
 * UX-1007: Compress Pictures re-encodes asynchronously, but Apply stayed
 * clickable through the encode — a double click compressed the picture twice
 * and screen readers heard nothing while it ran. The dialog now runs the
 * pdf insert dialog's UX-910 busy pattern: aria-busy on the modal, every
 * control disabled, and an always-mounted polite live region with a spinner
 * while the encoder works. A second Apply during the encode must not start a
 * second encode.
 */
import { beforeAll, describe, expect, it, vi, type Mock } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { LocaleProvider, loadLocale, setModuleLang, t } from '../src/renderer/i18n/locale'
import { CompressPicturesDialog } from '../src/renderer/components/PictureDialogs'
import {
  compressPictureDataUrl,
  imageSizeOf,
  type CompressResult,
} from '../src/renderer/editor/compress-picture'

vi.mock('../src/renderer/editor/compress-picture', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/renderer/editor/compress-picture')>()
  return {
    ...actual,
    // jsdom decodes no images: the probe answers a fixed natural size
    imageSizeOf: vi.fn(async () => ({ widthPx: 1200, heightPx: 900 })),
    compressPictureDataUrl: vi.fn(),
  }
})

const encode = vi.mocked(compressPictureDataUrl)

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // PERF-904: the en dictionary these assertions read must load first
  return loadLocale('en')
})

Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })
setModuleLang('en')

/** a controllable encode promise */
function deferredEncode() {
  let resolve!: (out: CompressResult | null) => void
  encode.mockImplementation(
    () =>
      new Promise((res) => {
        resolve = res
      }),
  )
  return (out: CompressResult | null) => resolve(out)
}

function mount(onApply: Mock<(result: { dataUrl: string; deleteCropped: boolean }) => void>): {
  container: HTMLElement
  root: Root
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() =>
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(CompressPicturesDialog, {
          dataUrl: 'data:image/png;base64,AAAA',
          displayWidthPx: 300,
          displayHeightPx: 200,
          crop: null,
          onApply,
          onCancel: vi.fn(),
        }),
      }),
    ),
  )
  return { container, root }
}

/** wait for the size probe effect and render it */
async function settle(): Promise<void> {
  await act(async () => {})
}

const byText = (container: HTMLElement, text: string): HTMLButtonElement =>
  [...container.querySelectorAll('button')].find((b) => b.textContent === text) as HTMLButtonElement

describe('Compress Pictures busy state (UX-1007)', () => {
  it('disables and announces while the encoder runs; a second Apply is a no-op', async () => {
    const onApply = vi.fn()
    const { container, root } = mount(onApply)
    await settle()
    expect(imageSizeOf).toHaveBeenCalled()
    const apply = byText(container, t('ribbonApply'))
    expect(apply.disabled).toBe(false)

    const finish = deferredEncode()
    await act(async () => {
      apply.click()
    })

    // busy: modal aria-busy, live region with the spinner line, all disabled
    const modal = container.querySelector('.compress-modal')!
    expect(modal.getAttribute('aria-busy')).toBe('true')
    const status = container.querySelector('[role="status"]')!
    expect(status.getAttribute('aria-live')).toBe('polite')
    expect(status.textContent).toContain(t('ribbonProcessing'))
    expect(status.querySelector('.compress-busy-spin')).not.toBeNull()
    expect(byText(container, t('ribbonApply')).disabled).toBe(true)
    expect(byText(container, t('ribbonCancel')).disabled).toBe(true)
    expect(
      (container.querySelector('input[name="compress-ppi"]') as HTMLInputElement).disabled,
    ).toBe(true)

    // a second click while encoding (disabled buttons still receive dispatched
    // events) must not start a second encode
    await act(async () => {
      apply.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(encode).toHaveBeenCalledTimes(1)

    finish({
      dataUrl: 'data:image/jpeg;base64,BBBB',
      widthPx: 625,
      heightPx: 438,
      mime: 'image/jpeg',
    })
    await settle()
    expect(encode).toHaveBeenCalledTimes(1)
    expect(modal.getAttribute('aria-busy')).toBe('false')
    expect(status.textContent).not.toContain(t('ribbonProcessing'))
    // the applied payload is the encode result (deleteCropped defaulted off)
    expect(onApply).toHaveBeenCalledTimes(1)
    expect(onApply).toHaveBeenCalledWith({
      dataUrl: 'data:image/jpeg;base64,BBBB',
      deleteCropped: false,
    })
    act(() => root.unmount())
  })

  it('clears busy and shows the no-gain error when the encode shrinks nothing', async () => {
    const { container, root } = mount(vi.fn())
    await settle()
    const finish = deferredEncode()
    await act(async () => {
      byText(container, t('ribbonApply')).click()
    })
    finish(null)
    await settle()
    expect(container.querySelector('.compress-modal')!.getAttribute('aria-busy')).toBe('false')
    expect(container.textContent).toContain(t('ribbonCompressNoGain'))
    // error state re-disables Apply until the dialog is reopened
    expect(byText(container, t('ribbonApply')).disabled).toBe(true)
    act(() => root.unmount())
  })
})
