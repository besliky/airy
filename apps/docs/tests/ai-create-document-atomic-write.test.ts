/**
 * AI create_document writes the generated pdf/md/html straight into the
 * default save folder (docs-main createAiDocument) — a file the user may
 * already have from an earlier run, since uniquePathIn only avoids names in
 * use at pick time. The write must stage the bytes and rename them into place
 * (atomicWriteFile): a plain writeFile truncates the target the moment it
 * opens it, so a crash or disk error mid-write destroys the previous file.
 * This pins the wiring of the direct-write branches; the stage+rename
 * mechanics are covered by packages/electron-utils/tests/atomic-write.test.ts.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const source = readFileSync(join(here, '../src/main/docs-main.ts'), 'utf8')

function functionBody(name: string): string {
  const start = source.indexOf(`export async function ${name}(`)
  expect(start, `${name} not found in docs main`).toBeGreaterThan(-1)
  return source.slice(start, source.indexOf('\n}\n', start))
}

describe('createAiDocument atomic writes', () => {
  it('the pdf branch renames the rendered bytes into place', () => {
    const body = functionBody('createAiDocument')
    expect(body).toContain('await atomicWriteFile(filePath, bytes)')
    expect(body).toContain("await atomicWriteFile(filePath, Buffer.from(content, 'utf8'))")
    expect(body).not.toMatch(/\bwriteFile\(/)
  })
})
