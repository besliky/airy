// Integration with a REAL LibreOffice installation (soffice on PATH or
// AIRY_SOFFICE): .doc/.odt -> temp .docx -> editable session -> save, and
// format:'origin' export back through the canonical filters. Skipped with an
// explicit reason when LibreOffice is absent — the S6 environment has none,
// so discovery/args/failures are covered by tests/soffice.test.ts against a
// fake binary and routing by tests/import-routing.test.ts.
import { readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { afterAll, beforeAll, describe, expect, it } from 'vitest'

import { findSoffice } from '../src/import/soffice.js'
import { DocxSession } from '../src/docx/session.js'
import { openDocument } from '../src/import/open.js'
import { buildOdtFixture } from './helpers/odf-fixture.js'

const tool = await findSoffice()
const describeWithSoffice = describe.skipIf(!tool)

const DOC_SAMPLE = fileURLToPath(
  new URL('../../../packages/file-parse/tests/fixtures/legacy-sample.doc', import.meta.url),
)

let root: string

beforeAll(async () => {
  if (!tool) return
  const { mkdtemp } = await import('node:fs/promises')
  root = await mkdtemp(join(tmpdir(), 'airy-doc-legacy-'))
})

afterAll(async () => {
  if (root) await rm(root, { recursive: true, force: true })
})

// OLE2 compound document magic (legacy .doc)
const DOC_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])

describeWithSoffice(
  'legacy word documents via LibreOffice ' +
    '(skipped when absent: soffice not found on PATH and AIRY_SOFFICE unset — install LibreOffice ' +
    'to run; the conversion wrapper itself is covered against a fake soffice in soffice.test.ts)',
  () => {
    it('opens .doc as an editable converted session and saves a sibling .docx', async () => {
      const docPath = join(root, 'legacy.doc')
      await writeFile(docPath, await readFile(DOC_SAMPLE))
      const session = await openDocument(docPath, root)
      expect(session).toBeInstanceOf(DocxSession)
      const docxSession = session as DocxSession
      const meta = docxSession.meta()
      expect(meta.kind).toBe('docx')
      expect(meta.converted).toBe(true)
      expect(meta.originPath).toBe(docPath)
      expect(meta.editable).toBe(true)

      const overview = docxSession.readDocument()
      expect(overview).toContain('Legacy Report')

      docxSession.insertContent('<h2>Agent Section</h2><p>Added by the copilot.</p>', 0)
      const saved = await docxSession.save()
      expect(saved.path).toBe(join(root, 'legacy.docx'))
      expect(saved.warnings?.[0]).toMatch(/original \.doc file .* left untouched/)
      // original legacy bytes untouched
      expect(Buffer.compare(await readFile(docPath), await readFile(DOC_SAMPLE))).toBe(0)
      await docxSession.close()
    })

    it('exports back to .doc with format origin (best-effort)', async () => {
      const docPath = join(root, 'roundtrip.doc')
      await writeFile(docPath, await readFile(DOC_SAMPLE))
      const session = await openDocument(docPath, root)
      expect(session).toBeInstanceOf(DocxSession)
      const docxSession = session as DocxSession
      docxSession.insertContent('<p>Origin export paragraph.</p>', -1)
      const saved = await docxSession.save(undefined, 'origin')
      expect(saved.path).toBe(docPath)
      expect(saved.format).toBe('origin')
      // the exported file is a real OLE2 Word 97 document
      const bytes = await readFile(docPath)
      expect(bytes.subarray(0, 8).equals(DOC_MAGIC)).toBe(true)
      await docxSession.close()
    })

    it('round-trips .odt: open, edit, export back to odt (zip magic)', async () => {
      const odtPath = join(root, 'source.odt')
      const odtBytes = await buildOdtFixture()
      await writeFile(odtPath, odtBytes)
      const session = await openDocument(odtPath, root)
      expect(session).toBeInstanceOf(DocxSession)
      const docxSession = session as DocxSession
      const meta = docxSession.meta()
      expect(meta.kind).toBe('docx')
      expect(meta.converted).toBe(true)
      expect(docxSession.readDocument()).toContain('ODF Source Document')

      docxSession.insertContent('<p>Edited through the copilot.</p>', 1)
      const saved = await docxSession.save(undefined, 'origin')
      expect(saved.path).toBe(odtPath)
      expect(saved.format).toBe('origin')
      const bytes = await readFile(odtPath)
      // a valid ODF package is a zip (PK)
      expect(bytes[0]).toBe(0x50)
      expect(bytes[1]).toBe(0x4b)
      await docxSession.close()
    })
  },
)
