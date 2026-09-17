import { describe, expect, it } from 'vitest'

/**
 * Friendly error mapping (src/shared/error-codes.ts): common Node filesystem
 * errno codes resolve to a stable key that the renderer toast and the
 * main-process error dialog localize. Codes come from `err.code` or the
 * leading errno token of a raw message string (IPC results carry the text).
 */
import { friendlyErrorKey } from '../src/shared/error-codes'

describe('friendlyErrorKey', () => {
  it('maps err.code from Error objects', () => {
    const err = (code: string) => Object.assign(new Error('boom'), { code })
    expect(friendlyErrorKey(err('ENOENT'))).toBe('enoent')
    expect(friendlyErrorKey(err('enoent'))).toBe('enoent')
    expect(friendlyErrorKey(err('EACCES'))).toBe('eperm')
    expect(friendlyErrorKey(err('EPERM'))).toBe('eperm')
    expect(friendlyErrorKey(err('EBUSY'))).toBe('ebusy')
    expect(friendlyErrorKey(err('ETXTBSY'))).toBe('ebusy')
    expect(friendlyErrorKey(err('EMFILE'))).toBe('emfile')
  })

  it('maps errno tokens embedded in raw message strings', () => {
    expect(friendlyErrorKey('ENOENT: no such file or directory, open /tmp/x.docx')).toBe('enoent')
    expect(friendlyErrorKey('EPERM: operation not permitted')).toBe('eperm')
    expect(friendlyErrorKey('EBUSY: resource busy or locked')).toBe('ebusy')
  })

  it('returns null for unrelated and malformed errors', () => {
    expect(friendlyErrorKey(null)).toBeNull()
    expect(friendlyErrorKey(undefined)).toBeNull()
    expect(friendlyErrorKey(new Error('Something else went wrong'))).toBeNull()
    expect(friendlyErrorKey(Object.assign(new Error('x'), { code: 'UNKNOWN' }))).toBeNull()
    expect(friendlyErrorKey(42)).toBeNull()
    expect(friendlyErrorKey('plain failure text')).toBeNull()
    expect(friendlyErrorKey('')).toBeNull()
  })

  it('does not match errno tokens outside Node print format', () => {
    // the regex anchors on the errno token followed by ':' — free text mentioning
    // a code name (or a bare tail like "spawn EACCES") is not a mappable error;
    // such real errors carry the code on the Error object instead
    expect(friendlyErrorKey('the file was EBUSY at the time')).toBeNull()
    expect(friendlyErrorKey('spawn EACCES')).toBeNull()
  })
})
