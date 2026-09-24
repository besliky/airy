import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest'

import { exportPdf, type FileActionContext } from '../src/renderer/file-actions'
import { loadLocale, setModuleLang, t } from '../src/renderer/i18n/locale'

// UX-1613: every export exit branch (success / cancel / error / rejected IPC)
// must settle the status bar — a failure may never leave "Exporting PDF…"
// hanging as if the export were still running.

/** Minimal FileActionContext: only the fields the export path reads. */
function makeCtx(statuses: string[]): FileActionContext {
  return {
    doc: { fileName: 'test.docx', parsed: {}, filePath: null, hash: '' },
    sections: [],
    section: { pageWidth: 12240, pageHeight: 15840 },
    header: null,
    footer: null,
    titlePg: false,
    evenOddHf: false,
    hfVariants: {},
    sectionHfEdits: {},
    pendingMixedExportRef: { current: false },
    bumpPendingExportTick: () => {},
    setShowPagePreview: () => {},
    setStatus: (s: string) => statuses.push(s),
  } as unknown as FileActionContext
}

/** One preview page: sends the export down the preview-open branch. */
function addPreviewPage(): void {
  const page = document.createElement('div')
  page.className = 'pv-page'
  page.style.width = '816px'
  page.style.height = '1056px'
  document.body.appendChild(page)
}

const w = window as unknown as { desktop?: unknown }
const savedDesktop = w.desktop

beforeAll(async () => {
  await loadLocale('en')
  setModuleLang('en')
})

afterEach(() => {
  document.querySelectorAll('.pv-page').forEach((el) => el.remove())
})

afterAll(() => {
  w.desktop = savedDesktop
})

describe('PDF export status handling (UX-1613)', () => {
  it('success replaces the exporting status with the exported message', async () => {
    const statuses: string[] = []
    w.desktop = { exportPdf: vi.fn().mockResolvedValue({ ok: true, path: '/out/doc.pdf' }) }
    const ok = await exportPdf(makeCtx(statuses))
    expect(ok).toBe(true)
    expect(statuses[0]).toBe(t('appExportingPdf'))
    expect(statuses[statuses.length - 1]).toBe(t('appExportedPdf', { path: '/out/doc.pdf' }))
    expect(statuses[statuses.length - 1]).not.toContain('Exporting')
  })

  it('cancel (no error payload) settles with the canceled message', async () => {
    const statuses: string[] = []
    w.desktop = { exportPdf: vi.fn().mockResolvedValue({ ok: false }) }
    const ok = await exportPdf(makeCtx(statuses))
    expect(ok).toBe(false)
    expect(statuses[0]).toBe(t('appExportingPdf'))
    expect(statuses[statuses.length - 1]).toBe(t('appExportPdfCanceled'))
    expect(statuses[statuses.length - 1]).not.toContain('Exporting')
  })

  it('an unauthorized-path refusal settles with the failure, not "Exporting…"', async () => {
    const statuses: string[] = []
    // lawful canPdfWrite refusal from the main process (authorized-path check)
    w.desktop = {
      exportPdf: vi
        .fn()
        .mockResolvedValue({ ok: false, error: 'export target is not an authorized path' }),
    }
    // single preview page: the refusal cannot retry through the merge path
    addPreviewPage()
    const ok = await exportPdf(makeCtx(statuses))
    expect(ok).toBe(false)
    expect(statuses[0]).toBe(t('appExportingPdf'))
    expect(statuses[statuses.length - 1]).toBe(
      t('appExportPdfFailed', { error: 'export target is not an authorized path' }),
    )
    expect(statuses[statuses.length - 1]).not.toContain('Exporting')
  })

  it('a rejected export IPC settles with the failure instead of hanging', async () => {
    const statuses: string[] = []
    w.desktop = { exportPdf: vi.fn().mockRejectedValue(new Error('ipc handler crashed')) }
    const ok = await exportPdf(makeCtx(statuses))
    expect(ok).toBe(false)
    expect(statuses[0]).toBe(t('appExportingPdf'))
    expect(statuses[statuses.length - 1]).toBe(
      t('appExportPdfFailed', { error: 'ipc handler crashed' }),
    )
    expect(statuses[statuses.length - 1]).not.toContain('Exporting')
  })

  it('a rejected merge IPC after printed parts settles with the failure too', async () => {
    const statuses: string[] = []
    // 11 preview pages: chunks into >1 print group, then merges via saveMergedPdf
    for (let i = 0; i < 11; i++) addPreviewPage()
    w.desktop = {
      printPdfBuffer: vi.fn().mockResolvedValue({ ok: true, base64: 'cGFydA==' }),
      saveMergedPdf: vi.fn().mockRejectedValue(new Error('merge crashed')),
    }
    const ok = await exportPdf(makeCtx(statuses))
    expect(ok).toBe(false)
    expect(statuses[0]).toBe(t('appExportingPdf'))
    expect(statuses[statuses.length - 1]).toBe(t('appExportPdfFailed', { error: 'merge crashed' }))
    expect(statuses[statuses.length - 1]).not.toContain('Exporting')
  })
})
