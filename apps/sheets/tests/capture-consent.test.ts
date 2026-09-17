import { describe, expect, it } from 'vitest'

import {
  CAPTURE_TOKEN_TTL_MS,
  CaptureConsentTracker,
  captureGrantValid,
  newCaptureGrant,
  randomCaptureToken,
} from '../src/main/capture-consent'

const NOW = 1_000_000
const TOKEN = 'tok1234567890abcd'
const MINUTE_MS = 60_000
const HOUR_MS = 60 * MINUTE_MS

describe('newCaptureGrant', () => {
  it('remembers the listed source ids and the deadline', () => {
    const grant = newCaptureGrant(['screen:0', 'window:1'], NOW, 5_000, 'tok')
    expect(grant.token).toBe('tok')
    expect(grant.sourceIds.has('screen:0')).toBe(true)
    expect(grant.sourceIds.has('window:1')).toBe(true)
    expect(grant.sourceIds.has('window:9')).toBe(false)
    expect(grant.expiresAt).toBe(NOW + 5_000)
  })

  it('issues unguessable tokens by default', () => {
    const a = randomCaptureToken()
    const b = randomCaptureToken()
    expect(a).toMatch(/^[0-9a-f]{32}$/)
    expect(a).not.toBe(b)
  })
})

describe('captureGrantValid', () => {
  const grant = newCaptureGrant(['screen:0', 'window:1'], NOW, CAPTURE_TOKEN_TTL_MS, TOKEN)

  it('accepts the matching token for a listed source within the ttl', () => {
    expect(captureGrantValid(grant, TOKEN, 'screen:0', NOW)).toBe(true)
    expect(captureGrantValid(grant, TOKEN, 'window:1', NOW + CAPTURE_TOKEN_TTL_MS - 1)).toBe(true)
  })

  it('rejects after expiry', () => {
    expect(captureGrantValid(grant, TOKEN, 'screen:0', NOW + CAPTURE_TOKEN_TTL_MS)).toBe(false)
  })

  it('rejects wrong tokens and unlisted sources', () => {
    expect(captureGrantValid(grant, `${TOKEN.slice(0, -1)}e`, 'screen:0', NOW)).toBe(false)
    expect(captureGrantValid(grant, 'tok', 'screen:0', NOW)).toBe(false)
    expect(captureGrantValid(grant, TOKEN, 'window:42', NOW)).toBe(false)
  })

  it('rejects malformed input and missing grants', () => {
    expect(captureGrantValid(null, TOKEN, 'screen:0', NOW)).toBe(false)
    expect(captureGrantValid(undefined, TOKEN, 'screen:0', NOW)).toBe(false)
    expect(captureGrantValid(grant, 42, 'screen:0', NOW)).toBe(false)
    expect(captureGrantValid(grant, TOKEN, undefined, NOW)).toBe(false)
    expect(captureGrantValid(grant, '', 'screen:0', NOW)).toBe(false)
  })
})

describe('CaptureConsentTracker (session state machine)', () => {
  const sources = ['screen:0', 'window:1']

  function tracker() {
    const logs: string[] = []
    return { tracker: new CaptureConsentTracker(undefined, (m) => logs.push(m)), logs }
  }

  it('refuses enumeration without an announced picker session', () => {
    const t = tracker()
    expect(t.tracker.beginEnumeration(1, NOW)).toEqual({
      ok: false,
      refusal: 'no-session',
    })
  })

  it('arms on enumeration inside an open session and redeems exactly once', () => {
    const t = tracker()
    expect(t.tracker.pickerOpened(1, NOW)).toEqual({ ok: true })
    expect(t.tracker.beginEnumeration(1, NOW)).toEqual({ ok: true })
    const grant = t.tracker.issueGrant(1, sources, NOW, CAPTURE_TOKEN_TTL_MS, TOKEN)
    expect(grant.token).toBe(TOKEN)
    expect(t.tracker.redeem(1, TOKEN, 'screen:0', NOW)).toBe(true)
    // token consumed: a second attempt with the same token fails
    expect(t.tracker.redeem(1, TOKEN, 'screen:0', NOW)).toBe(false)
  })

  it('re-arms after a capture within the same session (refresh after failure)', () => {
    const t = tracker()
    t.tracker.pickerOpened(1, NOW)
    t.tracker.beginEnumeration(1, NOW)
    t.tracker.issueGrant(1, sources, NOW, CAPTURE_TOKEN_TTL_MS, TOKEN)
    expect(t.tracker.redeem(1, TOKEN, 'screen:0', NOW)).toBe(true)
    expect(t.tracker.beginEnumeration(1, NOW + 1_000)).toEqual({ ok: true })
    const second = t.tracker.issueGrant(1, sources, NOW + 1_000)
    expect(t.tracker.redeem(1, second.token, 'window:1', NOW + 2_000)).toBe(true)
  })

  it('closing the picker drops any live token', () => {
    const t = tracker()
    t.tracker.pickerOpened(1, NOW)
    t.tracker.beginEnumeration(1, NOW)
    t.tracker.issueGrant(1, sources, NOW, CAPTURE_TOKEN_TTL_MS, TOKEN)
    t.tracker.pickerClosed(1)
    expect(t.tracker.redeem(1, TOKEN, 'screen:0', NOW)).toBe(false)
    expect(t.tracker.beginEnumeration(1, NOW + 1)).toEqual({ ok: false, refusal: 'no-session' })
  })

  it('keeps grants tab-bound (cross-tab tokens rejected)', () => {
    const t = tracker()
    t.tracker.pickerOpened(1, NOW)
    t.tracker.pickerOpened(2, NOW)
    t.tracker.beginEnumeration(1, NOW)
    const grant = t.tracker.issueGrant(1, sources, NOW, CAPTURE_TOKEN_TTL_MS, TOKEN)
    expect(t.tracker.redeem(2, grant.token, 'screen:0', NOW)).toBe(false)
    expect(t.tracker.redeem(1, grant.token, 'screen:0', NOW)).toBe(true)
  })

  it('forgets everything about a torn-down tab', () => {
    const t = tracker()
    t.tracker.pickerOpened(1, NOW)
    t.tracker.beginEnumeration(1, NOW)
    t.tracker.issueGrant(1, sources, NOW, CAPTURE_TOKEN_TTL_MS, TOKEN)
    t.tracker.forget(1)
    t.tracker.pickerOpened(1, NOW + HOUR_MS + 1)
    t.tracker.beginEnumeration(1, NOW + HOUR_MS + 1)
    t.tracker.issueGrant(1, sources, NOW + HOUR_MS + 1)
    // rate history was dropped with the tab, not carried into the new one
    t.tracker.forget(1)
    t.tracker.pickerOpened(1, NOW)
    expect(t.tracker.beginEnumeration(1, NOW)).toEqual({ ok: true })
  })
})

describe('CaptureConsentTracker (rate limits)', () => {
  const sources = ['screen:0']
  const limits = { sessionsPerMinute: 3, enumerationsPerHour: 4 }
  const minute = 60_000

  function tracker() {
    const logs: string[] = []
    return { tracker: new CaptureConsentTracker(limits, (m) => logs.push(m)), logs }
  }

  it('caps picker sessions per sliding minute', () => {
    const t = tracker()
    for (let i = 0; i < 3; i++) {
      expect(t.tracker.pickerOpened(7, i * 1_000)).toEqual({ ok: true })
      t.tracker.pickerClosed(7)
    }
    expect(t.tracker.pickerOpened(7, 3 * 1_000)).toEqual({
      ok: false,
      refusal: 'too-many-sessions',
    })
    // the window slides: a session older than a minute falls out
    expect(t.tracker.pickerOpened(7, minute + 1_000)).toEqual({ ok: true })
  })

  it('caps enumerations per sliding hour regardless of sessions', () => {
    const t = tracker()
    let now = 0
    for (let round = 0; round < 4; round++) {
      now += minute
      expect(t.tracker.pickerOpened(7, now)).toEqual({ ok: true })
      expect(t.tracker.beginEnumeration(7, now)).toEqual({ ok: true })
      t.tracker.issueGrant(7, sources, now)
      t.tracker.pickerClosed(7)
    }
    now += minute
    expect(t.tracker.pickerOpened(7, now)).toEqual({ ok: true })
    expect(t.tracker.beginEnumeration(7, now)).toEqual({
      ok: false,
      refusal: 'too-many-enumerations',
    })
    // the oldest enumeration ages out after an hour
    expect(t.tracker.beginEnumeration(7, 4 * minute + 60 * minute + 1)).toEqual({ ok: true })
  })

  it('rate limits are per tab, not global', () => {
    const t = tracker()
    for (let i = 0; i < 3; i++) {
      expect(t.tracker.pickerOpened(1, i)).toEqual({ ok: true })
      t.tracker.pickerClosed(1)
    }
    expect(t.tracker.pickerOpened(1, 3)).toEqual({ ok: false, refusal: 'too-many-sessions' })
    t.tracker.pickerOpened(2, 1)
    expect(t.tracker.beginEnumeration(2, 1)).toEqual({ ok: true })
  })

  it('logs every enumeration and every refusal (audit trail)', () => {
    const t = tracker()
    t.tracker.pickerOpened(5, NOW)
    for (let i = 0; i < 4; i++) {
      expect(t.tracker.beginEnumeration(5, NOW + i)).toEqual({ ok: true })
      t.tracker.issueGrant(5, sources, NOW + i)
    }
    t.tracker.beginEnumeration(5, NOW + 4)
    const joined = t.logs.join('\n')
    expect(joined).toContain('picker session opened for tab 5')
    expect(joined).toContain('enumeration served for tab 5')
    expect(joined).toContain('enumeration rate limit')
    expect(t.logs.filter((l) => l.includes('enumeration served')).length).toBe(4)
  })
})
