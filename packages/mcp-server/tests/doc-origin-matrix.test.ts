// Origin/save matrix for docx sessions (S6 Phase 3b): default targets for
// native vs converted sessions, byte-preserving behavior, and the
// format:'origin' guards. LibreOffice is mocked absent — the real soffice
// round-trips live in doc-legacy-integration.test.ts.
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

vi.mock('../src/import/soffice.js', () => ({
  SOFFICE_FILTERS: { docx: 'MS Word 2007 XML', doc: 'MS Word 97', odt: 'writer8', ods: 'calc8' },
  findSoffice: vi.fn(async () => null),
  sofficeMissingError: (reason: string) => new Error(`${reason} Install LibreOffice.`),
  convertViaSoffice: vi.fn(async () => {
    throw new Error('not available in unit tests')
  }),
}))

import { DocxSession, FencingError, type SessionOrigin } from '../src/docx/session.js'
import { buildFixtureDocx } from './helpers/docx-fixture.js'

let root: string
let fixture: Uint8Array

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'airy-origin-'))
  fixture = await buildFixtureDocx()
  await writeFile(join(root, 'report.docx'), fixture)
})

afterEach(async () => {
  await rm(root, { recursive: true, force: true })
})

describe('native docx sessions', () => {
  it('reports native meta (not converted, editable, no warnings)', async () => {
    const session = await DocxSession.open('report.docx', root)
    const meta = session.meta()
    expect(meta.kind).toBe('docx')
    expect(meta.format).toBe('docx')
    expect(meta.converted).toBe(false)
    expect(meta.editable).toBe(true)
    expect(meta.warnings).toEqual([])
  })

  it('rejects format origin without a conversion origin', async () => {
    const session = await DocxSession.open('report.docx', root)
    await expect(session.save(undefined, 'origin')).rejects.toThrow(
      /only valid for sessions converted from \.doc\/\.odt/,
    )
  })
})

describe('converted docx sessions (origin remembered)', () => {
  async function openConverted(format: 'doc' | 'odt'): Promise<DocxSession> {
    // simulate a soffice conversion: the temp docx is a copy of the fixture,
    // and the origin stamp is captured from disk like the real flow does
    const tempDir = join(root, 'temp-import')
    const originPath = join(root, format === 'doc' ? 'letter.doc' : 'letter.odt')
    await writeFile(originPath, 'legacy-bytes')
    const tempDocx = join(tempDir, 'letter.docx')
    const { mkdir, stat } = await import('node:fs/promises')
    await mkdir(tempDir, { recursive: true })
    await writeFile(tempDocx, fixture)
    const info = await stat(originPath)
    const origin: SessionOrigin = {
      path: originPath,
      format,
      stamp: { mtimeMs: info.mtimeMs, size: info.size },
      tempDir,
    }
    return DocxSession.open(tempDocx, root, origin)
  }

  it('meta reports the origin path, converted flag and warning', async () => {
    const session = await openConverted('doc')
    const meta = session.meta()
    expect(meta.converted).toBe(true)
    expect(meta.originPath).toBe(join(root, 'letter.doc'))
    expect(meta.path).toBe(join(root, 'letter.doc'))
    expect(meta.warnings[0]).toMatch(/Converted from \.doc via LibreOffice/)
  })

  it('default save writes a sibling .docx and leaves the origin untouched', async () => {
    const session = await openConverted('doc')
    session.insertContent('<p>Added by the agent.</p>', 0)
    const result = await session.save()
    expect(result.path).toBe(join(root, 'letter.docx'))
    expect(result.warnings?.[0]).toMatch(/original \.doc file .* was left untouched/)
    expect(await readFile(join(root, 'letter.doc'), 'utf8')).toBe('legacy-bytes')
    // the saved sibling reopens as a normal native session
    const reopened = await DocxSession.open('letter.docx', root)
    expect(reopened.meta().converted).toBe(false)
  })

  it('format origin requires LibreOffice when absent', async () => {
    const session = await openConverted('odt')
    await expect(session.save(undefined, 'origin')).rejects.toThrow(/Install LibreOffice/)
  })

  it('format origin fences when the origin changed since conversion', async () => {
    const session = await openConverted('doc')
    // rewrite the origin after the session captured its stamp — an external
    // writer touched the file since conversion
    await writeFile(join(root, 'letter.doc'), 'changed-by-someone-else')
    await expect(session.save(undefined, 'origin')).rejects.toBeInstanceOf(FencingError)
  })

  it('close removes the conversion temp dir', async () => {
    const session = await openConverted('doc')
    const cleaned = await session.close()
    expect(cleaned).toEqual([join(root, 'temp-import')])
    const { access } = await import('node:fs/promises')
    await expect(access(join(root, 'temp-import'))).rejects.toThrow()
  })
})
