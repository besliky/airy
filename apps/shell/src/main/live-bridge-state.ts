/**
 * Live-bridge effective-state decisions (pure, unit-tested).
 *
 * AIRY_DISABLE_BRIDGE=1 is a hard environment override (locked-down
 * machines, sandboxed runs): when set, the bridge server stays off no
 * matter what app-settings.json says, and a toggle request must not
 * persist — flipping the switch writes nothing, so removing the override
 * later restores the user's real preference instead of whatever the last
 * no-op click sent. The UI shows the effective state plus a note.
 */

/** env surface the decisions read (loose so tests and partial envs fit) */
type BridgeEnv = { AIRY_DISABLE_BRIDGE?: string | undefined }

export function bridgeEnvDisabled(env: BridgeEnv = process.env): boolean {
  return env.AIRY_DISABLE_BRIDGE === '1'
}

/**
 * Effective state: the env override wins over the stored preference, and an
 * absent stored value means enabled (the bridge ships on by default).
 */
export function effectiveLiveBridgeEnabled(stored: unknown, env: BridgeEnv = process.env): boolean {
  if (bridgeEnvDisabled(env)) return false
  return stored !== false
}

/** Whether a toggle request may persist and drive the server (false = visible no-op). */
export function liveBridgeToggleAllowed(env: BridgeEnv = process.env): boolean {
  return !bridgeEnvDisabled(env)
}
