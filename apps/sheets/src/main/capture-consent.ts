/**
 * Main-side consent for the sheets Insert > Screenshot picker.
 *
 * The enumeration channel (sheets:capture-screen-sources) returns the
 * capturable surfaces for the visible picker grid and hands the renderer a
 * short-lived, single-use capture token bound to its webContents. The
 * full-resolution channel (sheets:capture-screen-source) must present that
 * token and name one source id from the same enumeration. Each capture
 * consumes the grant, and issuing a new one replaces it.
 *
 * Electron has no native multi-source picker on Windows/Linux, so full
 * user-consent per enumeration cannot be enforced main-side. The guarantee is
 * therefore bounded + session-bound instead of "none": a per-tab state machine
 * only enumerates while the renderer has signaled an open picker session
 * (sheets:capture-picker-state), and sliding-window rate limits bound a
 * lying renderer that cycles sessions — max sessions per minute and max
 * enumerations per hour per tab. RESIDUAL: a compromised renderer can still
 * capture at the rate limit (one full-res frame per granted enumeration);
 * macOS additionally enforces the OS screen-recording permission, Windows and
 * Linux do not. Every enumeration is logged for an audit trail.
 * The logic is pure and unit-tested; sheets-main owns the per-tab tracker.
 */
import { randomBytes } from 'node:crypto'

/** Long enough for a user to browse the picker, short enough to bound a leak window. */
export const CAPTURE_TOKEN_TTL_MS = 90_000

/** Sessions = picker-open signals; a lying renderer must cycle these to enumerate. */
export const CAPTURE_SESSIONS_PER_MINUTE = 6
/** Harder cap on the thing that leaks (thumbnails + one full-res grant each). */
export const CAPTURE_ENUMERATIONS_PER_HOUR = 12

const SESSION_WINDOW_MS = 60_000
const ENUMERATION_WINDOW_MS = 60 * 60_000
/** Memory bound for the sliding windows (rate caps prune well below this). */
const MAX_TIMESTAMPS_KEPT = 1_000

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

// ── per-tab consent state machine ─────────────────────────────────────────
//
// idle → picker-open → armed (one enumeration with token) → captured → idle
//
// picker-open is signaled when the renderer's ScreenshotDialog opens and
// picker-close when it closes (sheets:capture-picker-state). Enumeration is
// only allowed inside an open session; each enumeration counts against the
// hourly cap and issues a fresh single-use token that replaces the previous
// one. Closing the picker (or tab teardown) drops any live token.

export type CapturePhase = 'idle' | 'picker-open' | 'armed' | 'captured'

export type CaptureRefusal = 'no-session' | 'too-many-sessions' | 'too-many-enumerations'

export type CaptureDecision = { ok: true } | { ok: false; refusal: CaptureRefusal }

interface CaptureTabState {
  phase: CapturePhase
  grant: CaptureGrant | null
  /** epoch-ms timestamps of picker sessions started inside the window */
  sessions: number[]
  /** epoch-ms timestamps of enumerations served inside the window */
  enumerations: number[]
}

export interface CaptureConsentLimits {
  sessionsPerMinute: number
  enumerationsPerHour: number
}

export const DEFAULT_CAPTURE_CONSENT_LIMITS: CaptureConsentLimits = {
  sessionsPerMinute: CAPTURE_SESSIONS_PER_MINUTE,
  enumerationsPerHour: CAPTURE_ENUMERATIONS_PER_HOUR,
}

/** Sliding-window check: keep only in-window stamps, then count them. */
function countInWindow(stamps: number[], now: number, windowMs: number): number {
  while (stamps.length > 0 && now - stamps[0]! > windowMs) stamps.shift()
  return stamps.length
}

function pruneStamps(stamps: number[]): void {
  while (stamps.length > MAX_TIMESTAMPS_KEPT) stamps.shift()
}

/**
 * Per-tab consent tracker: phase state machine plus sliding-window rate
 * limits. Pure (no Electron imports, clock and logger injected) so the whole
 * policy is unit-testable; sheets-main feeds it event.sender.id.
 */
export class CaptureConsentTracker {
  private readonly tabs = new Map<number, CaptureTabState>()

  constructor(
    private readonly limits: CaptureConsentLimits = DEFAULT_CAPTURE_CONSENT_LIMITS,
    private readonly log: (message: string) => void = () => {},
  ) {}

  private tab(wcId: number): CaptureTabState {
    let state = this.tabs.get(wcId)
    if (!state) {
      state = { phase: 'idle', grant: null, sessions: [], enumerations: [] }
      this.tabs.set(wcId, state)
    }
    return state
  }

  /** Renderer signaled its screenshot picker opening; starts (and counts) a session. */
  pickerOpened(wcId: number, now: number): CaptureDecision {
    const tab = this.tab(wcId)
    if (countInWindow(tab.sessions, now, SESSION_WINDOW_MS) >= this.limits.sessionsPerMinute) {
      this.log(`capture-consent: picker session for tab ${wcId} refused (session rate limit)`)
      return { ok: false, refusal: 'too-many-sessions' }
    }
    tab.sessions.push(now)
    pruneStamps(tab.sessions)
    // a re-open without a close signal (or after a capture) restarts cleanly
    tab.phase = 'picker-open'
    tab.grant = null
    this.log(`capture-consent: picker session opened for tab ${wcId}`)
    return { ok: true }
  }

  /** Renderer signaled its screenshot picker closing; drops any live token. */
  pickerClosed(wcId: number): void {
    const tab = this.tabs.get(wcId)
    if (!tab) return
    tab.phase = 'idle'
    tab.grant = null
  }

  /**
   * Gate an enumeration attempt: must run inside an open picker session and
   * inside the hourly enumeration budget. Call BEFORE desktopCapturer work;
   * a refusal leaves the phase (and any earlier grant) untouched.
   */
  beginEnumeration(wcId: number, now: number): CaptureDecision {
    const tab = this.tab(wcId)
    if (tab.phase === 'idle') {
      this.log(`capture-consent: enumeration for tab ${wcId} refused (no open picker session)`)
      return { ok: false, refusal: 'no-session' }
    }
    if (
      countInWindow(tab.enumerations, now, ENUMERATION_WINDOW_MS) >= this.limits.enumerationsPerHour
    ) {
      this.log(`capture-consent: enumeration for tab ${wcId} refused (enumeration rate limit)`)
      return { ok: false, refusal: 'too-many-enumerations' }
    }
    tab.enumerations.push(now)
    pruneStamps(tab.enumerations)
    return { ok: true }
  }

  /**
   * Store the grant for a completed enumeration (replacing any earlier one)
   * and arm the tab for one full-res capture.
   */
  issueGrant(
    wcId: number,
    sourceIds: readonly string[],
    now: number,
    ttlMs: number = CAPTURE_TOKEN_TTL_MS,
    token: string = randomCaptureToken(),
  ): CaptureGrant {
    const tab = this.tab(wcId)
    const grant = newCaptureGrant(sourceIds, now, ttlMs, token)
    tab.phase = 'armed'
    tab.grant = grant
    this.log(
      `capture-consent: enumeration served for tab ${wcId} ` +
        `(${sourceIds.length} sources listed, ${tab.enumerations.length} in window)`,
    )
    return grant
  }

  /**
   * Validate a full-res capture attempt against the tab's live grant and
   * consume it on success (armed → captured). A failed attempt (wrong token,
   * unlisted source, expiry, other tab) consumes nothing.
   */
  redeem(wcId: number, token: unknown, sourceId: unknown, now: number): boolean {
    const tab = this.tabs.get(wcId)
    if (!tab || !captureGrantValid(tab.grant, token, sourceId, now)) return false
    tab.grant = null
    tab.phase = 'captured'
    this.log(`capture-consent: full-res capture redeemed for tab ${wcId}`)
    return true
  }

  /** Tab teardown: forget the phase, grant, and rate history. */
  forget(wcId: number): void {
    this.tabs.delete(wcId)
  }
}
