/**
 * Markdown exports (docx via conversion, pdf via the hidden print window,
 * html standalone via PAR-315) write into the path the user picked in a save
 * dialog — typically over an existing file. The write must stage the bytes
 * and rename them into place (atomicWriteFile): a plain writeFile truncates
 * the target the moment it opens it, so a crash or disk error mid-write
 * destroys the previous file. This pins the wiring of all export handlers;
 * the stage+rename mechanics are covered by
 * packages/electron-utils/tests/atomic-write.test.ts. The remaining plain
 * writeFile writes into the handler's own mkdtemp staging dir, where a torn
 * file is discarded with the dir anyway.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const source = readFileSync(join(here, '../src/main/markdown-main.ts'), 'utf8')

describe('markdown export atomic writes', () => {
  it('the docx export writes the user-picked path atomically', () => {
    expect(source).toContain('await atomicWriteFile(picked.filePath, bytes)')
  })

  it('the pdf export writes the user-picked path atomically', () => {
    expect(source).toContain('await atomicWriteFile(picked.filePath, pdf)')
  })

  it('the html export stages the document and writes the user-picked path atomically', () => {
    expect(source).toContain("const htmlBytes = Buffer.from(request.html, 'utf8')")
    expect(source).toContain('await atomicWriteFile(picked.filePath, htmlBytes)')
  })

  it('no export writes straight into a user-picked path', () => {
    expect(source).not.toMatch(/writeFile\(picked\.filePath/)
  })
})
