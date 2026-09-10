import { describe, expect, it, vi } from 'vitest'

import {
  DEFAULT_START_DELAY_MS,
  RELEASES_LATEST_URL,
  UpdaterController,
  detectUpdatePolicy,
  releaseNotesUrl,
  type UpdaterClient,
  type UpdaterStatus,
  type UpdaterUi,
} from '../src/main/updater/core'

/**
 * Updater state machine (src/main/updater/core.ts) against mock feeds: the
 * in-app flow (check → offer → download → install), the quiet no-update and
 * network-error paths, the deb notify-only fallback, and the macOS noop.
 * core.ts is pure — no Electron, no electron-updater.
 */

type MockClient = ReturnType<typeof createMockClient>

function createMockClient() {
  const progressCallbacks: Array<(percent: number) => void> = []
  const errorCallbacks: Array<(error: unknown) => void> = []
  return {
    checkForUpdates: vi.fn<UpdaterClient['checkForUpdates']>(),
    downloadUpdate: vi.fn<UpdaterClient['downloadUpdate']>(),
    quitAndInstall: vi.fn<UpdaterClient['quitAndInstall']>(),
    onDownloadProgress: vi.fn((callback: (percent: number) => void) => {
      progressCallbacks.push(callback)
    }),
    onError: vi.fn((callback: (error: unknown) => void) => {
      errorCallbacks.push(callback)
    }),
    emitProgress(percent: number) {
      for (const callback of progressCallbacks) callback(percent)
    },
    emitError(error: unknown) {
      for (const callback of errorCallbacks) callback(error)
    },
  }
}

function createMockUi() {
  return {
    offerUpdate: vi.fn<UpdaterUi['offerUpdate']>(async () => 'download' as const),
    offerInstall: vi.fn<UpdaterUi['offerInstall']>(async () => 'now' as const),
    notifyAboutRelease: vi.fn<UpdaterUi['notifyAboutRelease']>(),
  }
}

/** flush pending promise chains (each await in the controller is a microtask) */
const flush = () => new Promise<void>((resolve) => setTimeout(resolve, 0))

/** downloadUpdate mock that stays pending until the test resolves it */
function pendingDownload(client: MockClient): () => void {
  let finish!: () => void
  client.downloadUpdate.mockImplementation(
    () =>
      new Promise((resolve) => {
        finish = () => resolve(undefined)
      }),
  )
  return () => finish()
}

interface Harness {
  client: MockClient
  ui: ReturnType<typeof createMockUi>
  log: ReturnType<typeof vi.fn>
  statuses: UpdaterStatus[]
  controller: UpdaterController
}

function createHarness(policy: 'in-app' | 'notify-only' | 'off' = 'in-app'): Harness {
  const client = createMockClient()
  const ui = createMockUi()
  const log = vi.fn()
  const statuses: UpdaterStatus[] = []
  const controller = new UpdaterController({
    client,
    ui,
    policy,
    log,
    transientStatusMs: 5,
    onStatusChange: (status) => statuses.push(status),
  })
  return { client, ui, log, statuses, controller }
}

describe('detectUpdatePolicy', () => {
  it('windows and AppImage linux self-update in-app', () => {
    expect(detectUpdatePolicy('win32')).toBe('in-app')
    expect(detectUpdatePolicy('linux', { APPIMAGE: '/opt/Airy.AppImage' })).toBe('in-app')
  })

  it('linux without the APPIMAGE marker is a deb-style install (notify only)', () => {
    expect(detectUpdatePolicy('linux', {})).toBe('notify-only')
  })

  it('macOS (and anything unknown) keeps the updater off', () => {
    expect(detectUpdatePolicy('darwin')).toBe('off')
    expect(detectUpdatePolicy('freebsd')).toBe('off')
  })
})

describe('in-app policy (nsis / appimage)', () => {
  it('checks, offers, downloads with progress, and installs after confirmation', async () => {
    const h = createHarness('in-app')
    const finishDownload = pendingDownload(h.client)
    h.client.checkForUpdates.mockResolvedValue({ available: true, version: '0.10.0' })
    h.controller.checkNow()
    await flush()

    // available dialog carries the version and the release-notes URL
    expect(h.ui.offerUpdate).toHaveBeenCalledWith({
      version: '0.10.0',
      notesUrl: releaseNotesUrl('0.10.0'),
    })
    expect(h.ui.offerUpdate).toHaveBeenCalledTimes(1)

    // the dialog confirmation starts the download (still in flight)
    expect(h.client.downloadUpdate).toHaveBeenCalledTimes(1)
    expect(h.controller.status).toEqual({ phase: 'downloading', percent: 0 })

    // progress ticks update the downloading status shown in the menu
    h.client.emitProgress(41.6)
    expect(h.controller.status).toEqual({ phase: 'downloading', percent: 42 })
    h.client.emitProgress(100)
    expect(h.controller.status).toEqual({ phase: 'downloading', percent: 100 })

    // finishing the download offers install; "now" quits into the installer
    finishDownload()
    await vi.waitFor(() => expect(h.ui.offerInstall).toHaveBeenCalledTimes(1))
    expect(h.ui.offerInstall).toHaveBeenCalledWith({ version: '0.10.0' })
    expect(h.controller.status).toEqual({ phase: 'ready' })
    expect(h.client.quitAndInstall).toHaveBeenCalledTimes(1)
    expect(h.log).not.toHaveBeenCalled()
  })

  it('defers the startup check instead of checking at once', () => {
    vi.useFakeTimers()
    try {
      const h = createHarness('in-app')
      h.client.checkForUpdates.mockResolvedValue({ available: false, version: null })
      h.controller.start()
      expect(h.client.checkForUpdates).not.toHaveBeenCalled()
      vi.advanceTimersByTime(DEFAULT_START_DELAY_MS - 1)
      expect(h.client.checkForUpdates).not.toHaveBeenCalled()
      vi.advanceTimersByTime(1)
      expect(h.client.checkForUpdates).toHaveBeenCalledTimes(1)
    } finally {
      vi.useRealTimers()
    }
  })

  it('"install on quit" keeps the app running and the ready status', async () => {
    const h = createHarness('in-app')
    h.client.checkForUpdates.mockResolvedValue({ available: true, version: '0.10.0' })
    h.client.downloadUpdate.mockResolvedValue(undefined)
    h.ui.offerInstall.mockResolvedValue('on-quit')
    h.controller.checkNow()
    await vi.waitFor(() => expect(h.ui.offerInstall).toHaveBeenCalledTimes(1))
    expect(h.client.quitAndInstall).not.toHaveBeenCalled()
    expect(h.controller.status).toEqual({ phase: 'ready' })
  })

  it('declining the available dialog returns to idle without downloading', async () => {
    const h = createHarness('in-app')
    h.client.checkForUpdates.mockResolvedValue({ available: true, version: '0.10.0' })
    h.ui.offerUpdate.mockResolvedValue('later')
    h.controller.checkNow()
    await flush()
    expect(h.client.downloadUpdate).not.toHaveBeenCalled()
    expect(h.controller.status).toEqual({ phase: 'idle' })
  })
})

describe('no update available', () => {
  it('stays silent: transient up-to-date menu status, no dialogs', async () => {
    const h = createHarness('in-app')
    h.client.checkForUpdates.mockResolvedValue({ available: false, version: null })
    h.controller.checkNow()
    await flush()
    expect(h.ui.offerUpdate).not.toHaveBeenCalled()
    expect(h.controller.status).toEqual({ phase: 'up-to-date' })
    await vi.waitFor(() => expect(h.controller.status).toEqual({ phase: 'idle' }))
    expect(h.log).not.toHaveBeenCalled()
  })
})

describe('network failure', () => {
  it('is swallowed: logged to stderr, transient failed status, never a dialog', async () => {
    const h = createHarness('in-app')
    h.client.checkForUpdates.mockRejectedValue(new Error('getaddrinfo ENOTFOUND api.github.com'))
    h.controller.checkNow()
    await flush()
    expect(h.ui.offerUpdate).not.toHaveBeenCalled()
    expect(h.controller.status).toEqual({ phase: 'failed' })
    expect(h.log).toHaveBeenCalledWith(
      'update check failed',
      expect.objectContaining({ message: 'getaddrinfo ENOTFOUND api.github.com' }),
    )
    await vi.waitFor(() => expect(h.controller.status).toEqual({ phase: 'idle' }))
  })

  it('a failed download marks failed without surfacing an error dialog', async () => {
    const h = createHarness('in-app')
    h.client.checkForUpdates.mockResolvedValue({ available: true, version: '0.10.0' })
    h.client.downloadUpdate.mockRejectedValue(new Error('404 blockmap'))
    h.controller.checkNow()
    await vi.waitFor(() => expect(h.statuses).toContainEqual({ phase: 'failed' }))
    expect(h.ui.offerInstall).not.toHaveBeenCalled()
    expect(h.log).toHaveBeenCalledWith('update download failed', expect.any(Error))
    await vi.waitFor(() => expect(h.controller.status).toEqual({ phase: 'idle' }))
  })

  it('a spurious autoUpdater error outside a download is logged only', async () => {
    const h = createHarness('in-app')
    h.client.checkForUpdates.mockResolvedValue({ available: false, version: null })
    h.controller.checkNow()
    await flush()
    h.client.emitError(new Error('race'))
    // up-to-date (the transient status), not flipped to failed
    expect(h.controller.status).toEqual({ phase: 'up-to-date' })
    expect(h.log).toHaveBeenCalledWith('autoUpdater error', expect.any(Error))
    expect(h.ui.offerUpdate).not.toHaveBeenCalled()
  })
})

describe('deb installs (notify-only)', () => {
  it('notifies with the releases URL and never downloads', async () => {
    const h = createHarness('notify-only')
    h.client.checkForUpdates.mockResolvedValue({ available: true, version: '0.10.0' })
    h.controller.checkNow()
    await flush()
    expect(h.ui.notifyAboutRelease).toHaveBeenCalledWith({
      version: '0.10.0',
      releasesUrl: RELEASES_LATEST_URL,
    })
    expect(h.ui.offerUpdate).not.toHaveBeenCalled()
    expect(h.client.downloadUpdate).not.toHaveBeenCalled()
    expect(h.client.quitAndInstall).not.toHaveBeenCalled()
    expect(h.controller.status).toEqual({ phase: 'idle' })
  })
})

describe('macOS (off policy)', () => {
  it('is a noop: no scheduled check, manual check refused, idle forever', () => {
    vi.useFakeTimers()
    try {
      const h = createHarness('off')
      h.controller.start()
      vi.advanceTimersByTime(DEFAULT_START_DELAY_MS * 2)
      h.controller.checkNow()
      expect(h.client.checkForUpdates).not.toHaveBeenCalled()
      expect(h.controller.status).toEqual({ phase: 'idle' })
      expect(h.controller.active).toBe(false)
      expect(h.log).not.toHaveBeenCalled()
    } finally {
      vi.useRealTimers()
    }
  })
})

describe('re-check guards', () => {
  it('ignores manual checks while a flow is busy', async () => {
    const h = createHarness('in-app')
    pendingDownload(h.client)
    h.client.checkForUpdates.mockResolvedValue({ available: true, version: '0.10.0' })
    h.controller.checkNow()
    await flush()
    expect(h.controller.status).toEqual({ phase: 'downloading', percent: 0 })
    h.controller.checkNow() // downloading: refused
    expect(h.client.checkForUpdates).toHaveBeenCalledTimes(1)
  })
})
