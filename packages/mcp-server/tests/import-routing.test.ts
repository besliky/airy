// Import routing for open_document (S6): extension dispatch, the .odt
// without-LibreOffice error, and the .doc read-only text fallback driven by
// the real word-extractor over the committed legacy sample fixture.
// LibreOffice discovery is mocked to "absent" so routing is deterministic on
// machines that do have soffice installed.
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

vi.mock('../src/import/soffice.js', () => ({
  SOFFICE_FILTERS: { docx: 'MS Word 2007 XML', doc: 'MS Word 97', odt: 'writer8', ods: 'calc8' },
  findSoffice: vi.fn(async () => null),
  sofficeMissingError: (reason: string) => new Error(`${reason} Install LibreOffice.`),
  convertViaSoffice: vi.fn(async () => {
    throw new Error('not available in unit tests')
  }),
}))

import { extensionOf, openDocument } from '../src/import/open.js'
import { TextSession } from '../src/sessions/text.js'

// committed binary Word 97-2003 sample from @airy-office/file-parse
const DOC_SAMPLE = fileURLToPath(
  new URL('../../../packages/file-parse/tests/fixtures/legacy-sample.doc', import.meta.url),
)

// committed password-protected OOXML fixtures (ECMA-376 encrypted: an OLE2
// container holding an EncryptedPackage stream) from the docs app suite
const ENCRYPTED_FIXTURES = ['office-agile-password.docx', 'office-standard-password.docx'].map(
  (name) =>
    fileURLToPath(new URL(`../../../apps/docs/tests/encrypted-fixtures/${name}`, import.meta.url)),
)

let root: string

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'airy-routing-'))
})

afterEach(async () => {
  await rm(root, { recursive: true, force: true })
})

describe('extensionOf', () => {
  it('lowercases and strips the dot', () => {
    expect(extensionOf('/a/b/REPORT.DOCX')).toBe('docx')
    expect(extensionOf('/a/b/legacy.Xls')).toBe('xls')
    expect(extensionOf('noext')).toBe('')
  })
})

describe('openDocument routing', () => {
  it('rejects unsupported extensions with the supported list', async () => {
    await expect(openDocument(join(root, 'archive.zip'), root)).rejects.toThrow(
      /Unsupported file type "\.zip".*\.docx, \.xlsx/,
    )
  })

  it('refuses legacy .ppt with a conversion hint (.pptx itself is supported)', async () => {
    await expect(openDocument(join(root, 'deck.ppt'), root)).rejects.toThrow(
      /Convert the deck to \.pptx/,
    )
  })

  it('rejects .odt without LibreOffice with an install hint', async () => {
    // routing fails before the file is ever read, so no fixture is needed
    await expect(openDocument('doc.odt', root)).rejects.toThrow(/requires LibreOffice/)
  })

  it('opens .doc read-only via word-extractor when LibreOffice is absent', async () => {
    const doc = join(root, 'legacy.doc')
    await copyFile(DOC_SAMPLE, doc)
    const session = await openDocument(doc, root)
    expect(session).toBeInstanceOf(TextSession)
    const textSession = session as TextSession
    const meta = textSession.meta()
    expect(meta.format).toBe('doc')
    expect(meta.converted).toBe(false)
    expect(meta.editable).toBe(false)
    expect(meta.warnings[0]).toMatch(/read-only/)
    const text = textSession.readDocument()
    expect(text).toContain('Legacy Report')
    expect(text).toContain('Second paragraph from Word 97-2003.')
    expect(text).toContain('editable: false')
    // read-only sessions close cleanly with nothing to clean
    expect(await textSession.close()).toEqual([])
  })

  it('confines every route to the workspace root', async () => {
    await expect(openDocument('../../etc/hosts.docx', root)).rejects.toThrow(
      /outside the workspace root/,
    )
  })
})

describe('encrypted Office containers refuse with a password hint (BUG-1504)', () => {
  it('refuses an encrypted .docx naming password protection, not the raw-zip parse error', async () => {
    for (const [index, fixture] of ENCRYPTED_FIXTURES.entries()) {
      const docx = join(root, `protected-${String(index)}.docx`)
      await copyFile(fixture, docx)
      await expect(openDocument(docx, root)).rejects.toThrow(/password-protected/)
      await expect(openDocument(docx, root)).rejects.not.toThrow(/central directory/)
    }
  })

  it('refuses an encrypted .doc (no LibreOffice) instead of the word-extractor crash', async () => {
    const doc = join(root, 'protected.doc')
    await copyFile(ENCRYPTED_FIXTURES[0]!, doc)
    await expect(openDocument(doc, root)).rejects.toThrow(
      /Cannot open ".*protected\.doc": the file is password-protected/,
    )
    await expect(openDocument(doc, root)).rejects.not.toThrow(/memory outside buffer bounds/)
  })

  it('refuses an encrypted container on the xlsx route before the sidecar is spawned', async () => {
    const book = join(root, 'protected.xlsx')
    await copyFile(ENCRYPTED_FIXTURES[1]!, book)
    await expect(openDocument(book, root)).rejects.toThrow(/password-protected/)
  })

  it('detects the legacy .doc FIB fEncrypted flag (no EncryptedPackage stream)', async () => {
    // CFB magic + the fEncrypted bit in the FIB base flags (offset 0x0A,
    // bit 0x0100) — a password-protected Word 97 document has no
    // EncryptedPackage stream, so this signal must be checked separately
    const bytes = new Uint8Array(4096)
    bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0)
    bytes[0x0b] = 0x01
    const doc = join(root, 'legacy-encrypted.doc')
    await writeFile(doc, bytes)
    await expect(openDocument(doc, root)).rejects.toThrow(/password-protected/)
  })

  it('refuses a plain OLE document renamed to .docx as a wrong container, not a corrupt zip', async () => {
    // the committed sample is a REAL unencrypted Word 97 OLE file
    const docx = join(root, 'actually-doc.docx')
    await copyFile(DOC_SAMPLE, docx)
    await expect(openDocument(docx, root)).rejects.toThrow(/OLE2 compound document.*not a \.docx/)
  })
})
