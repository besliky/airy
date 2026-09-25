import { describe, expect, it } from 'vitest'

import { excelLegacyPasswordHash, legacyPasswordMatches } from '../src/shared/legacy-password'

describe('excelLegacyPasswordHash', () => {
  // Ground-truth vectors from openpyxl's hash_password (the same public
  // ECMA-376 algorithm Excel uses for the `password=` attribute).
  it('matches openpyxl hash_password vectors', () => {
    const vectors: readonly [string, string][] = [
      ['password', '83AF'],
      ['secret', 'DAA7'],
      ['x', 'CEBA'],
      ['', 'CE4B'],
      ['abc123', 'C58F'],
      ['Password1', 'E1EE'],
      ['пароль', 'A370'],
      ['密码', '99C3'],
    ]
    for (const [password, hash] of vectors) {
      expect(excelLegacyPasswordHash(password)).toBe(hash)
    }
  })

  it('writes 4 uppercase hex digits', () => {
    for (const password of ['a', 'zz', '0', 'long-password-with-digits-123']) {
      expect(excelLegacyPasswordHash(password)).toMatch(/^[0-9A-F]{4}$/)
    }
  })

  it('verifies candidates case-insensitively against a stored hash', () => {
    const hash = excelLegacyPasswordHash('S3cret!')
    expect(legacyPasswordMatches('S3cret!', hash)).toBe(true)
    expect(legacyPasswordMatches('S3cret!', hash.toLowerCase())).toBe(true)
    expect(legacyPasswordMatches('s3cret!', hash)).toBe(false)
    expect(legacyPasswordMatches('', hash)).toBe(false)
  })
})
