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
import type { BridgeCallContext } from './dispatcher'
import { generateBridgeToken, startBridgeServer, type BridgeServerHandle } from './server'

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
 * aborts on bridge timeout; a dying tab settles the call as tab_closed so no
 * waiter can hang past the call's lifetime (SA3 §orphan-tab risk). Both death
 * signals are watched: 'destroyed' covers a graceful teardown, while
 * 'render-process-gone' covers a killed renderer, whose webContents is NOT
 * destroyed — without it every call into a corpse burned the full dispatcher
 * timeout (BUG-1697).
 */
function callRenderer(
  webContents: WebContents,
  method: string,
  params: Record<string, unknown>,
  signal: AbortSignal,
  context: BridgeCallContext,
): Promise<BridgeCommandResult> {
  return new Promise<BridgeCommandResult>((resolve) => {
    if (webContents.isDestroyed()) {
      resolve({ ok: false, error: { code: 'tab_closed', message: 'the document tab is gone' } })
      return
    }
    const requestId = nextRequestId++
    const settle = (result: BridgeCommandResult) => {
      pendingCalls.delete(requestId)
      webContents.off('destroyed', onDeath)
      webContents.off('render-process-gone', onDeath)
      signal.removeEventListener('abort', onAbort)
      resolve(result)
    }
    const onDeath = () =>
      settle({
        ok: false,
        error: { code: 'tab_closed', message: 'the document tab renderer is gone' },
      })
    const onAbort = () =>
      settle({ ok: false, error: { code: 'timeout', message: 'the renderer call was aborted' } })
    pendingCalls.set(requestId, { resolve: settle, webContentsId: webContents.id })
    webContents.once('destroyed', onDeath)
    webContents.once('render-process-gone', onDeath)
    signal.addEventListener('abort', onAbort, { once: true })
    // the connection identity rides along so the renderer can attribute
    // bridge turns to the client that made them
    webContents.send(BRIDGE_INVOKE_CHANNEL, requestId, method, params, context.clientId)
  })
}

/** Route a document command to the ACTIVE docs tab; v1 targets docs only. */
async function callActiveDocs(
  getTabManager: () => TabManager | null,
  method: string,
  params: Record<string, unknown>,
  signal: AbortSignal,
  context: BridgeCallContext,
): Promise<unknown> {
  const tab = getTabManager()?.activeDocsTab()
  if (!tab) {
    throw new BridgeMethodError(
      'not_docs_tab',
      'the active tab is not a docs document (live bridge v1 targets docs tabs only)',
    )
  }
  if (tab.dead) {
    // BUG-1697: fast typed failure — reopening the file (openPath) builds a
    // fresh webContents and revives the document; calling into the corpse
    // would otherwise wait out the whole dispatcher timeout for a renderer
    // that no longer exists
    throw new BridgeMethodError(
      'tab_closed',
      'the active docs tab renderer is gone; open the file again to restore the tab',
    )
  }
  if (tab.webContents.isDestroyed()) {
    throw new BridgeMethodError('tab_closed', 'the active docs tab is gone')
  }
  const reply = await callRenderer(tab.webContents, method, params, signal, context)
  if (!reply.ok) {
    // renderer codes are strings on the wire; unknown codes surface as internal
    const code = isBridgeErrorCode(reply.error.code) ? reply.error.code : 'internal'
    throw new BridgeMethodError(code, reply.error.message)
  }
  return reply.result
}

let server: BridgeServerHandle | null = null

/**
 * BUG-1313: the bridge token lives for the PROCESS, not per server start. An
 * aborted quit restarts the bridge (stop→start); regenerating would rewrite
 * the published info file and invalidate every connected client's token —
 * "cancel quit" must leave the app as if nothing happened. Generated once
 * per process, reused across restarts; a fresh process generates anew.
 */
let processToken: string | null = null

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
    token: (processToken ??= generateBridgeToken()),
    methods: {
      ping: () => ({
        pong: true,
        protocolVersion: 1,
        pid: process.pid,
      }),
      list: () => ({ documents: getTabManager()?.docsTabSummaries() ?? [] }),
      get_context: (params, signal, context) =>
        callActiveDocs(getTabManager, 'get_context', params, signal, context),
      apply_ops: (params, signal, context) => {
        if (!Array.isArray(params.ops)) {
          throw new BridgeMethodError('invalid_params', 'params.ops must be an array of ops')
        }
        return callActiveDocs(getTabManager, 'apply_ops', params, signal, context)
      },
      insert_content: (params, signal, context) => {
        if (typeof params.html !== 'string' || params.html.trim() === '') {
          throw new BridgeMethodError(
            'invalid_params',
            'params.html must be a non-empty restricted-HTML string',
          )
        }
        return callActiveDocs(getTabManager, 'insert_content', params, signal, context)
      },
      undo: (params, signal, context) =>
        callActiveDocs(getTabManager, 'undo', params, signal, context),
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
