import { describe, expect, it } from 'vitest'

import {
  CAPTURE_TOKEN_TTL_MS,
  captureGrantValid,
  newCaptureGrant,
  randomCaptureToken,
} from '../src/main/capture-consent'

const NOW = 1_000_000
const TOKEN = 'tok1234567890abcd'

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
