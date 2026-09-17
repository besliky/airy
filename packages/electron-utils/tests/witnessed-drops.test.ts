import { beforeEach, describe, expect, it } from 'vitest'

import {
  forgetWitnessedDrops,
  mayGrantAttachmentRead,
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
