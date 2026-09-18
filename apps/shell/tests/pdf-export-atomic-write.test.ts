/**
 * The pdf→Word/PowerPoint/Excel exports write converted bytes into the path
 * the user picked in a save dialog — typically over an existing file. The
 * write must stage the bytes and rename them into place (atomicWriteFile): a
 * plain writeFileSync truncates the target the moment it opens it, so a crash
 * or disk error mid-write destroys the previous file. These tests pin the
 * wiring of all three export handlers; the stage+rename mechanics themselves
 * are covered by packages/electron-utils/tests/atomic-write.test.ts.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const source = readFileSync(join(here, '../src/main/index.ts'), 'utf8')

function functionBody(name: string): string {
  const start = source.indexOf(`async function ${name}(`)
  expect(start, `${name} not found in shell main`).toBeGreaterThan(-1)
  return source.slice(start, source.indexOf('\n}\n', start))
}

describe('pdf export atomic writes', () => {
  for (const [name, payload] of [
    ['exportPdfAsDocxLocal', 'result.docx'],
    ['exportPdfAsPptxLocal', 'result.pptx'],
    ['exportPdfAsXlsxLocal', 'result.xlsx'],
  ] as const) {
    it(`${name} renames the converted bytes into place`, () => {
      const body = functionBody(name)
      expect(body).toContain(`await atomicWriteFile(picked.filePath, ${payload})`)
      expect(body).not.toMatch(/writeFileSync\(/)
    })
  }
})
