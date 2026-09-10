// Discovery of the Airy live bridge info file (airy-bridge.json): the running
// app writes it next to its socket in userData with mode 0600, carrying
// {socketPath, token, pid, protocolVersion}. The client rereads the file on
// every connect (lazily) — an app restart changes the token, a stale file must
// never authorize a reconnect. Pure Node, no Electron imports.
import { readFile } from 'node:fs/promises'
import { homedir } from 'node:os'
import { join } from 'node:path'

/** absolute path to the bridge info file; overrides all default locations */
export const BRIDGE_INFO_FILE_ENV = 'AIRY_BRIDGE_FILE'

/** the info file name the bridge server writes into userData */
export const BRIDGE_INFO_NAME = 'airy-bridge.json'

/**
 * userData directory names searched for the info file, in order:
 * - 'Airy'          — productName of the shell app (apps/shell/package.json);
 *                     packaged runs keep userData in appData/Airy
 * - 'Airy Dev'      — dev/unpacked runs (the main process redirects userData
 *                     to appData/'Airy Dev' when !app.isPackaged, unless
 *                     AIRY_USER_DATA points elsewhere — dev setups like
 *                     that need AIRY_BRIDGE_FILE)
 * - 'GenOffice' / 'GenOffice Dev' — legacy pre-rename userData layouts, kept
 *                     so a bridge file from an older install is still found
 */
export const BRIDGE_APP_NAME_CANDIDATES: readonly string[] = [
  'Airy',
  'Airy Dev',
  'GenOffice',
  'GenOffice Dev',
]

/** the shape published in the info file (twin of the server's BridgeEndpointInfo) */
export interface BridgeEndpointInfo {
  socketPath: string
  token: string
  pid: number
  protocolVersion: number
}

/** Electron's app.getPath('appData') base directory per platform */
function appDataBase(platform: NodeJS.Platform, env: NodeJS.ProcessEnv, home: string): string {
  if (platform === 'win32') return env.APPDATA || join(home, 'AppData', 'Roaming')
  if (platform === 'darwin') return join(home, 'Library', 'Application Support')
  // linux & the bsds: Electron honors XDG_CONFIG_HOME, falling back to ~/.config
  return env.XDG_CONFIG_HOME || join(home, '.config')
}

/**
 * Ordered candidate paths for the info file. With AIRY_BRIDGE_FILE set, that
 * single path is the whole list (explicit configuration wins — no silent
 * fallback to a differently-configured app instance); otherwise the per-
 * platform userData defaults for every app-name candidate are searched in
 * order.
 */
export function candidateBridgeInfoPaths(
  options: { env?: NodeJS.ProcessEnv; platform?: NodeJS.Platform; homeDir?: string } = {},
): string[] {
  const env = options.env ?? process.env
  const platform = options.platform ?? process.platform
  const home = options.homeDir ?? homedir()
  const fromEnv = env[BRIDGE_INFO_FILE_ENV]
  if (fromEnv && fromEnv.trim() !== '') return [fromEnv]
  const base = appDataBase(platform, env, home)
  return BRIDGE_APP_NAME_CANDIDATES.map((name) => join(base, name, BRIDGE_INFO_NAME))
}

/** parse + validate one info file payload; null when malformed */
function parseInfoFile(text: string): BridgeEndpointInfo | null {
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch {
    return null
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) return null
  const obj = parsed as Record<string, unknown>
  if (typeof obj.socketPath !== 'string' || obj.socketPath === '') return null
  if (typeof obj.token !== 'string' || obj.token === '') return null
  if (typeof obj.pid !== 'number' || !Number.isInteger(obj.pid)) return null
  if (typeof obj.protocolVersion !== 'number' || !Number.isInteger(obj.protocolVersion)) return null
  return {
    socketPath: obj.socketPath,
    token: obj.token,
    pid: obj.pid,
    protocolVersion: obj.protocolVersion,
  }
}

export interface DiscoveredBridge {
  info: BridgeEndpointInfo
  /** which candidate file the info came from (diagnostics) */
  path: string
}

/**
 * The first readable info file among the candidates, or null when the bridge
 * is not discoverable. Missing/unreadable/malformed candidates are skipped.
 */
export async function discoverBridgeInfo(
  options: { env?: NodeJS.ProcessEnv; platform?: NodeJS.Platform; homeDir?: string } = {},
): Promise<DiscoveredBridge | null> {
  for (const path of candidateBridgeInfoPaths(options)) {
    try {
      const info = parseInfoFile(await readFile(path, 'utf8'))
      if (info) return { info, path }
    } catch {
      // missing or unreadable — try the next candidate
    }
  }
  return null
}
