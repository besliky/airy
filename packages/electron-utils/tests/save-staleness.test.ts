import { dirname, join } from 'node:path'
import { mkdtemp, rename, rm, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { afterEach, describe, expect, it } from 'vitest'

import { checkSaveStaleness, statFileStamp } from '../src/save-staleness'

const temporaryDirectories: string[] = []

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  )
})

async function createFile(content: string): Promise<string> {
  const directory = await mkdtemp(join(tmpdir(), 'save-staleness-'))
  temporaryDirectories.push(directory)
  const path = join(directory, 'doc.md')
  await writeFile(path, content)
  return path
}

describe('statFileStamp', () => {
  it('captures mtime and size of an existing file', async () => {
    const path = await createFile('hello')
    expect(statFileStamp(path)).toEqual({ mtimeMs: expect.any(Number), size: 5 })
  })

  it('returns null for a missing file', () => {
    expect(statFileStamp(join(tmpdir(), 'save-staleness-does-not-exist.md'))).toBeNull()
  })
})

describe('checkSaveStaleness', () => {
  it('reports fresh when nothing touched the file since the stamp', async () => {
    const path = await createFile('# Saved')
    expect(checkSaveStaleness(path, statFileStamp(path))).toBe('fresh')
  })

  it('reports changed after an external rewrite', async () => {
    const path = await createFile('# Saved')
    const stamp = statFileStamp(path)
    await new Promise((resolve) => setTimeout(resolve, 20))
    await writeFile(path, '# Evil')
    expect(checkSaveStaleness(path, stamp)).toBe('changed')
  })

  it('reports changed on a size change even within the same mtime tick', async () => {
    const path = await createFile('# Saved')
    const stamp = statFileStamp(path)
    await writeFile(path, '# Saved longer')
    utimes(path, new Date(), new Date(stamp!.mtimeMs))
    expect(checkSaveStaleness(path, stamp)).toBe('changed')
  })

  it('reports missing after an external rename', async () => {
    const path = await createFile('# Saved')
    const stamp = statFileStamp(path)
    await rename(path, join(dirname(path), 'moved.md'))
    expect(checkSaveStaleness(path, stamp)).toBe('missing')
  })

  it('treats an absent baseline as fresh (no fence yet)', async () => {
    const path = await createFile('# Saved')
    expect(checkSaveStaleness(path, null)).toBe('fresh')
    expect(checkSaveStaleness(path, undefined)).toBe('fresh')
  })
})
