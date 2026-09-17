import { afterEach, describe, expect, it } from 'vitest'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { loadMediaReference } from '../src/media-tools'

describe('loadMediaReference local-path policy', () => {
  let dir: string | undefined
  afterEach(async () => {
    if (dir) await rm(dir, { recursive: true, force: true })
    dir = undefined
  })

  async function withLocalPng(fn: (path: string) => Promise<void>): Promise<void> {
    dir = await mkdtemp(join(tmpdir(), 'airy-media-policy-'))
    // minimal PNG header is enough: the loader sniffs nothing locally, it
    // trusts the media extension for the mime and reads the bytes verbatim
    const path = join(dir, 'pic.png')
    await writeFile(path, Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))
    await fn(path)
  }

  it('reads a local media file when no policy is injected (pure package default)', async () => {
    await withLocalPng(async (path) => {
      const blob = await loadMediaReference(path)
      expect(blob.mime).toBe('image/png')
      expect(blob.name).toBe('pic.png')
      expect(blob.bytes.byteLength).toBe(8)
    })
  })

  it('reads a local media file the policy allows', async () => {
    await withLocalPng(async (path) => {
      const blob = await loadMediaReference(path, { mayReadFile: () => true })
      expect(blob.bytes.byteLength).toBe(8)
    })
  })

  it('refuses a local media file the policy rejects, without reading it', async () => {
    await withLocalPng(async (path) => {
      const mayReadFile = (candidate: string): boolean => candidate !== path
      await expect(loadMediaReference(path, { mayReadFile })).rejects.toThrow(
        /Not an accessible media path/,
      )
    })
  })

  it('still rejects non-media extensions before consulting the policy', async () => {
    await withLocalPng(async (_path) => {
      const secret = join(dir!, 'secret.zip')
      await writeFile(secret, Buffer.from('x'))
      let consulted = false
      await expect(
        loadMediaReference(secret, {
          mayReadFile: () => {
            consulted = true
            return true
          },
        }),
      ).rejects.toThrow(/Unsupported media file/)
      expect(consulted).toBe(false)
    })
  })
})
