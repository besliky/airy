/**
 * Main-side consent for the sheets Insert > Screenshot picker.
 *
 * The enumeration channel (sheets:capture-screen-sources) returns the
 * capturable surfaces for the visible picker grid and hands the renderer a
 * short-lived, single-use capture token bound to its webContents. The
 * full-resolution channel (sheets:capture-screen-source) must present that
 * token and name one source id from the same enumeration. Each capture
 * consumes the grant, and issuing a new one replaces it, so a compromised
 * renderer cannot silently pull every screen at full resolution — it gets
 * at most one frame per enumeration round of the (visible) picker flow.
 * The logic is pure and unit-tested; sheets-main owns the per-tab map.
 */
import { randomBytes } from 'node:crypto'

/** Long enough for a user to browse the picker, short enough to bound a leak window. */
export const CAPTURE_TOKEN_TTL_MS = 90_000

export interface CaptureGrant {
  /** Unguessable token presented by the renderer for the full-res capture */
  token: string
  /** Source ids that were part of this enumeration */
  sourceIds: ReadonlySet<string>
  /** Epoch-ms deadline; the grant is invalid afterwards */
  expiresAt: number
}

export function randomCaptureToken(bytes = 16): string {
  return randomBytes(bytes).toString('hex')
}

/**
 * Issue a grant for an enumeration. `now` and the token generator are
 * injectable for tests.
 */
export function newCaptureGrant(
  sourceIds: readonly string[],
  now: number,
  ttlMs: number = CAPTURE_TOKEN_TTL_MS,
  token: string = randomCaptureToken(),
): CaptureGrant {
  return { token, sourceIds: new Set(sourceIds), expiresAt: now + ttlMs }
}

/** Pure decision: may the renderer redeem this grant for `sourceId`? */
export function captureGrantValid(
  grant: CaptureGrant | null | undefined,
  token: unknown,
  sourceId: unknown,
  now: number,
): boolean {
  if (!grant || typeof token !== 'string' || typeof sourceId !== 'string') return false
  if (now >= grant.expiresAt) return false
  // constant-time compare for the token (it is a capability)
  if (!timingSafeEqualCompat(Buffer.from(token), Buffer.from(grant.token))) return false
  return grant.sourceIds.has(sourceId)
}

function timingSafeEqualCompat(a: Buffer, b: Buffer): boolean {
  if (a.length !== b.length) return false
  let diff = 0
  for (let i = 0; i < a.length; i++) diff |= a[i]! ^ b[i]!
  return diff === 0
}
