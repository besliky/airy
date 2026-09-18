/**
 * Html app exports (docx via html2docx, pdf via the print pipeline) write
 * into the path the user picked in a save dialog — typically over an existing
 * file. The write must stage the bytes and rename them into place
 * (atomicWriteFile): a plain writeFile truncates the target the moment it
 * opens it, so a crash or disk error mid-write destroys the previous file.
 * This pins the wiring of both export handlers; the stage+rename mechanics
 * are covered by packages/electron-utils/tests/atomic-write.test.ts. The
 * remaining plain writeFile writes into the handler's own mkdtemp staging
 * dir, where a torn file is discarded with the dir anyway.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const source = readFileSync(join(here, '../src/main/html-main.ts'), 'utf8')

describe('html export atomic writes', () => {
  it('the docx export writes the user-picked path atomically', () => {
    expect(source).toContain('await atomicWriteFile(picked.filePath, docx)')
  })

  it('the pdf export writes the user-picked path atomically', () => {
    expect(source).toContain('await atomicWriteFile(picked.filePath, await renderPrintPdf(')
  })

  it('no export writes straight into a user-picked path', () => {
    expect(source).not.toMatch(/writeFile\(picked\.filePath/)
  })
})
