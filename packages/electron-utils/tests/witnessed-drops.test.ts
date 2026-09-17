import { beforeEach, describe, expect, it } from 'vitest'

import {
  forgetWitnessedDrops,
  mayGrantAttachmentRead,
  MAX_ATTACHMENT_ADD_PATHS,
  MAX_ATTACHMENT_PATH_CHARS,
  parseAttachmentPaths,
  recordWitnessedDrops,
  resetWitnessedDrops,
  witnessedDroppedPath,
  WITNESS_DROP_CHANNEL,
} from '../src/witnessed-drops'
import { grantRendererDir, resetRendererFileGrants } from '../src/renderer-file-access'

const SENDER = 7
const OTHER = 8

beforeEach(() => {
  resetWitnessedDrops()
  resetRendererFileGrants()
})

describe('recordWitnessedDrops / witnessedDroppedPath', () => {
  it('records and queries dropped paths per sender', () => {
    recordWitnessedDrops(SENDER, ['/tmp/notes.txt', '/tmp/img.png'], 1_000)
    expect(witnessedDroppedPath(SENDER, '/tmp/notes.txt', 1_500)).toBe(true)
    expect(witnessedDroppedPath(SENDER, '/tmp/missing.txt', 1_500)).toBe(false)
    expect(witnessedDroppedPath(OTHER, '/tmp/notes.txt', 1_500)).toBe(false)
  })

  it('trims, ignores junk entries, and caps one event at 20 paths', () => {
    recordWitnessedDrops(SENDER, ['  /a  ', 42, '', null, '/b'], 1_000)
    expect(witnessedDroppedPath(SENDER, '/a', 1_000)).toBe(true)
    expect(witnessedDroppedPath(SENDER, '/b', 1_000)).toBe(true)
    recordWitnessedDrops(
      SENDER,
      Array.from({ length: 40 }, (_, i) => `/tmp/f${i}`),
      1_000,
    )
    expect(witnessedDroppedPath(SENDER, '/tmp/f19', 1_000)).toBe(true)
    expect(witnessedDroppedPath(SENDER, '/tmp/f20', 1_000)).toBe(false)
  })

  it('expires entries after the TTL and forgets torn-down senders', () => {
    recordWitnessedDrops(SENDER, ['/tmp/old.txt'], 1_000)
    expect(witnessedDroppedPath(SENDER, '/tmp/old.txt', 1_000 + 30 * 60_000 + 1)).toBe(false)
    recordWitnessedDrops(SENDER, ['/tmp/gone.txt'], 1_000)
    forgetWitnessedDrops(SENDER)
    expect(witnessedDroppedPath(SENDER, '/tmp/gone.txt', 1_000)).toBe(false)
  })
})

describe('mayGrantAttachmentRead policy', () => {
  it('allows witnessed paths and paths already inside granted directories', () => {
    recordWitnessedDrops(SENDER, ['/tmp/dropped/report.pdf'], 1_000)
    expect(mayGrantAttachmentRead(SENDER, '/tmp/dropped/report.pdf', 1_500)).toBe(true)
    grantRendererDir('/home/user/docs', SENDER)
    expect(mayGrantAttachmentRead(SENDER, '/home/user/docs/secret.zip', 1_500)).toBe(true)
  })

  it('refuses renderer-named paths with no user-driven origin (the self-grant)', () => {
    expect(mayGrantAttachmentRead(SENDER, '/home/user/secret', 1_000)).toBe(false)
    expect(mayGrantAttachmentRead(SENDER, '/etc/passwd', 1_000)).toBe(false)
    // a witness for one sender must not bless another sender's claim
    recordWitnessedDrops(SENDER, ['/tmp/x.txt'], 1_000)
    expect(mayGrantAttachmentRead(OTHER, '/tmp/x.txt', 1_000)).toBe(false)
  })

  it('a grant held by another sender does not bless this sender (per-sender allowlist)', () => {
    grantRendererDir('/home/user/docs', OTHER)
    expect(mayGrantAttachmentRead(OTHER, '/home/user/docs/secret.zip', 1_000)).toBe(true)
    expect(mayGrantAttachmentRead(SENDER, '/home/user/docs/secret.zip', 1_000)).toBe(false)
  })
})

describe('channel name', () => {
  it('is the drop-open witness channel', () => {
    expect(WITNESS_DROP_CHANNEL).toBe('app:witnessed-dropped-files')
  })
})

describe('parseAttachmentPaths (files:add shape policy)', () => {
  it('accepts a bounded list of non-empty strings, trimmed', () => {
    expect(parseAttachmentPaths([' /a/b.docx ', 'c.pdf'])).toEqual(['/a/b.docx', 'c.pdf'])
  })

  it('rejects non-array and empty payloads', () => {
    expect(parseAttachmentPaths(null)).toBeNull()
    expect(parseAttachmentPaths('/a/b.docx')).toBeNull()
    expect(parseAttachmentPaths({})).toBeNull()
    expect(parseAttachmentPaths([])).toBeNull()
  })

  it('rejects non-string, empty, and over-long entries', () => {
    expect(parseAttachmentPaths(['/ok.docx', 42])).toBeNull()
    expect(parseAttachmentPaths(['/ok.docx', null])).toBeNull()
    expect(parseAttachmentPaths(['/ok.docx', '   '])).toBeNull()
    expect(
      parseAttachmentPaths(['/ok.docx', `${'x'.repeat(MAX_ATTACHMENT_PATH_CHARS + 1)}`]),
    ).toBeNull()
    expect(parseAttachmentPaths(['x'.repeat(MAX_ATTACHMENT_PATH_CHARS)])).not.toBeNull()
  })

  it('caps the list length like the sheets zod gate', () => {
    const fifty = Array.from({ length: MAX_ATTACHMENT_ADD_PATHS }, (_, i) => `/f${i}`)
    expect(parseAttachmentPaths(fifty)).not.toBeNull()
    expect(parseAttachmentPaths([...fifty, '/one-too-many'])).toBeNull()
  })
})
