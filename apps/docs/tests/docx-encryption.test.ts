// Crypto half of the docx-encryption suite (PERF-1002 split): container
// detection, wrapper error mapping and real interop against Office-produced
// files. These tests deliberately run the real officecrypto-tool — the agile
// key setup costs ~0.5s of 100k-iteration hashing per operation, so the whole
// file shares one encrypted fixture instead of each test re-encrypting, and
// the in-memory password/intent state machine lives in the sibling
// docx-encryption-password-store.test.ts with a cheap backend injected.
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  decryptDocx,
  DocxDecryptError,
  encryptDocx,
  isEncryptedDocx,
} from '../src/main/docx-encryption'

const plain = readFileSync(join(__dirname, 'pagination-corpus/docx/kitchen-sink.docx'))

// one real agile encrypt (~0.5s) shared by every test below
const encrypted = encryptDocx(plain, 'S3cret!密码')

describe('isEncryptedDocx', () => {
  it('is false for a plain docx (zip) and random bytes', () => {
    expect(isEncryptedDocx(plain)).toBe(false)
    expect(isEncryptedDocx(Buffer.from('not a docx at all'))).toBe(false)
    expect(isEncryptedDocx(Buffer.alloc(0))).toBe(false)
  })

  it('is true for an ECMA-376 encrypted docx', () => {
    expect(isEncryptedDocx(encrypted)).toBe(true)
  })

  it('is false for a CFB container without an EncryptedPackage stream (legacy .doc shape)', () => {
    const cfbNoPackage = Buffer.concat([
      Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
      Buffer.alloc(512),
    ])
    expect(isEncryptedDocx(cfbNoPackage)).toBe(false)
  })
})

describe('encryptDocx / decryptDocx', () => {
  it('round-trips byte-exact with the right password', async () => {
    expect(encrypted.subarray(0, 4)).not.toEqual(plain.subarray(0, 4))
    const decrypted = await decryptDocx(encrypted, 'S3cret!密码')
    expect(Buffer.compare(decrypted, plain)).toBe(0)
  })

  it('rejects a wrong password with reason wrong-password', async () => {
    const err = await decryptDocx(encrypted, 'wrong').catch((e) => e)
    expect(err).toBeInstanceOf(DocxDecryptError)
    expect((err as DocxDecryptError).reason).toBe('wrong-password')
  })

  it('reports unrecognized containers as unsupported', async () => {
    // CFB magic but no usable encryption streams (e.g. proprietary/account encryption)
    const bogus = Buffer.concat([
      Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
      Buffer.alloc(4096),
    ])
    const err = await decryptDocx(bogus, 'any').catch((e) => e)
    expect(err).toBeInstanceOf(DocxDecryptError)
    expect((err as DocxDecryptError).reason).toBe('unsupported')
  })
})

// Interop lock: fixtures produced by real Microsoft Office, vendored from the
// reference implementation's test suite (nolze/msoffcrypto-tool, MIT). The agile
// one is >4096 bytes, covering multi-segment package decryption.
describe('real Office files', () => {
  const fixture = (name: string) => readFileSync(join(__dirname, 'encrypted-fixtures', name))

  it('decrypts a real Office agile-encrypted docx byte-exact (multi-segment)', async () => {
    const agile = fixture('office-agile-password.docx')
    expect(isEncryptedDocx(agile)).toBe(true)
    const decrypted = await decryptDocx(agile, 'Password1234_')
    expect(Buffer.compare(decrypted, fixture('office-agile-plain.docx'))).toBe(0)
  })

  it('decrypts a real Office standard-encrypted docx byte-exact', async () => {
    const standard = fixture('office-standard-password.docx')
    expect(isEncryptedDocx(standard)).toBe(true)
    const decrypted = await decryptDocx(standard, 'Password1234_')
    expect(Buffer.compare(decrypted, fixture('office-standard-plain.docx'))).toBe(0)
  })
})
