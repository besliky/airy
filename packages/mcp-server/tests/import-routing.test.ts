// Import routing for open_document (S6): extension dispatch, the .odt
// without-LibreOffice error, and the .doc read-only text fallback driven by
// the real word-extractor over the committed legacy sample fixture.
// LibreOffice discovery is mocked to "absent" so routing is deterministic on
// machines that do have soffice installed.
import { copyFile, mkdtemp, rm } from 'node:fs/promises'
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

// committed binary Word 97-2003 sample from @genoffice/file-parse
const DOC_SAMPLE = fileURLToPath(
  new URL('../../../packages/file-parse/tests/fixtures/legacy-sample.doc', import.meta.url),
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
    await expect(openDocument(join(root, 'deck.pptx'), root)).rejects.toThrow(
      /Unsupported file type "\.pptx".*\.docx, \.xlsx/,
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
