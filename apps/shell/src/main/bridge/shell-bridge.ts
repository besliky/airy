// Electron glue for the Airy live bridge: wires the transport/server modules
// (pure Node) to the shell's TabManager and forwards document commands into
// the active docs renderer over an IPC round-trip modeled on docs:close-check
// (docs-main.ts queryCloseState) — but keyed by requestId so concurrent calls
// to different tabs cannot strand each other's waiters. This module is the
// only bridge file that imports Electron.
import { ipcMain } from 'electron'
import type { WebContents } from 'electron'

import type { TabManager } from '../tab-manager'
import type { BridgeCommandResult } from '../../../../docs/src/shared/ipc'
import { BRIDGE_INVOKE_CHANNEL, BRIDGE_RESULT_CHANNEL } from '../../../../docs/src/shared/ipc'
import { BridgeMethodError, isBridgeErrorCode } from './protocol'
import { startBridgeServer, type BridgeServerHandle } from './server'

/** default per-call timeout; the env var AIRY_DISABLE_BRIDGE=1 turns the bridge off entirely */
export const DEFAULT_BRIDGE_TIMEOUT_MS = 30_000

interface PendingRendererCall {
  resolve: (result: BridgeCommandResult) => void
  webContentsId: number
}

const pendingCalls = new Map<number, PendingRendererCall>()
let nextRequestId = 1
let ipcInstalled = false

function normalizeRendererResult(raw: unknown): BridgeCommandResult {
  if (typeof raw === 'object' && raw !== null && 'ok' in raw) {
    const r = raw as {
      ok?: unknown
      result?: unknown
      error?: { code?: unknown; message?: unknown }
    }
    if (r.ok === true) return { ok: true, result: r.result }
    if (
      r.ok === false &&
      typeof r.error === 'object' &&
      r.error !== null &&
      typeof r.error.code === 'string' &&
      typeof r.error.message === 'string'
    ) {
      return { ok: false, error: { code: r.error.code, message: r.error.message } }
    }
  }
  return { ok: false, error: { code: 'internal', message: 'malformed renderer reply' } }
}

function ensureResultChannel(): void {
  if (ipcInstalled) return
  ipcInstalled = true
  ipcMain.on(BRIDGE_RESULT_CHANNEL, (event, requestId: unknown, result: unknown) => {
    const waiter = pendingCalls.get(Number(requestId))
    // a late reply after timeout/abort has no waiter; a mismatched sender is dropped
    if (!waiter || waiter.webContentsId !== event.sender.id) return
    pendingCalls.delete(Number(requestId))
    waiter.resolve(normalizeRendererResult(result))
  })
}

/**
 * Send one command to a docs renderer and await its envelope. The signal
 * aborts on bridge timeout; a destroyed tab resolves as tab_closed so no
 * waiter can hang past the call's lifetime (SA3 §orphan-tab risk).
 */
function callRenderer(
  webContents: WebContents,
  method: string,
  params: Record<string, unknown>,
  signal: AbortSignal,
): Promise<BridgeCommandResult> {
  return new Promise<BridgeCommandResult>((resolve) => {
    if (webContents.isDestroyed()) {
      resolve({ ok: false, error: { code: 'tab_closed', message: 'the document tab is gone' } })
      return
    }
    const requestId = nextRequestId++
    const settle = (result: BridgeCommandResult) => {
      pendingCalls.delete(requestId)
      webContents.off('destroyed', onDestroyed)
      signal.removeEventListener('abort', onAbort)
      resolve(result)
    }
    const onDestroyed = () =>
      settle({
        ok: false,
        error: { code: 'tab_closed', message: 'the document tab was closed mid-call' },
      })
    const onAbort = () =>
      settle({ ok: false, error: { code: 'timeout', message: 'the renderer call was aborted' } })
    pendingCalls.set(requestId, { resolve: settle, webContentsId: webContents.id })
    webContents.once('destroyed', onDestroyed)
    signal.addEventListener('abort', onAbort, { once: true })
    webContents.send(BRIDGE_INVOKE_CHANNEL, requestId, method, params)
  })
}

/** Route a document command to the ACTIVE docs tab; v1 targets docs only. */
async function callActiveDocs(
  getTabManager: () => TabManager | null,
  method: string,
  params: Record<string, unknown>,
  signal: AbortSignal,
): Promise<unknown> {
  const tab = getTabManager()?.activeDocsTab()
  if (!tab) {
    throw new BridgeMethodError(
      'not_docs_tab',
      'the active tab is not a docs document (live bridge v1 targets docs tabs only)',
    )
  }
  if (tab.webContents.isDestroyed()) {
    throw new BridgeMethodError('tab_closed', 'the active docs tab is gone')
  }
  const reply = await callRenderer(tab.webContents, method, params, signal)
  if (!reply.ok) {
    // renderer codes are strings on the wire; unknown codes surface as internal
    const code = isBridgeErrorCode(reply.error.code) ? reply.error.code : 'internal'
    throw new BridgeMethodError(code, reply.error.message)
  }
  return reply.result
}

let server: BridgeServerHandle | null = null

/**
 * Start the live bridge in the shell main process (on by default;
 * AIRY_DISABLE_BRIDGE=1 turns it off). Called from app.whenReady next to
 * startSheetsCaptureServer.
 */
export async function startShellBridge(options: {
  userDataDir: string
  getTabManager: () => TabManager | null
  timeoutMs?: number
  log?: (message: string) => void
}): Promise<BridgeServerHandle | null> {
  if (process.env.AIRY_DISABLE_BRIDGE === '1') return null
  if (server) return server
  ensureResultChannel()
  const { userDataDir, getTabManager, timeoutMs, log } = options
  server = await startBridgeServer({
    userDataDir,
    timeoutMs,
    log,
    methods: {
      ping: () => ({
        pong: true,
        protocolVersion: 1,
        pid: process.pid,
      }),
      list: () => ({ documents: getTabManager()?.docsTabSummaries() ?? [] }),
      get_context: (params, signal) => callActiveDocs(getTabManager, 'get_context', params, signal),
      apply_ops: (params, signal) => {
        if (!Array.isArray(params.ops)) {
          throw new BridgeMethodError('invalid_params', 'params.ops must be an array of ops')
        }
        return callActiveDocs(getTabManager, 'apply_ops', params, signal)
      },
      insert_content: (params, signal) => {
        if (typeof params.html !== 'string' || params.html.trim() === '') {
          throw new BridgeMethodError(
            'invalid_params',
            'params.html must be a non-empty restricted-HTML string',
          )
        }
        return callActiveDocs(getTabManager, 'insert_content', params, signal)
      },
      undo: (params, signal) => callActiveDocs(getTabManager, 'undo', params, signal),
    },
  })
  return server
}

/** stop the bridge (app quit): closes the listener and removes socket + token file */
export async function stopShellBridge(): Promise<void> {
  const handle = server
  server = null
  await handle?.stop()
}
