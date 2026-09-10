// Binary discovery for the xlsx sidecar: env override, repo-checkout walk
// and the missing-binary error. All pure/fs checks, no Rust toolchain
// needed.
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it } from 'vitest'

import {
  findSidecarBinary,
  repoRootOf,
  SIDECAR_ENV,
  sidecarCandidates,
  sidecarMissingError,
} from '../src/xlsx/discovery.js'

const TESTS_DIR = fileURLToPath(new URL('.', import.meta.url))

let dir: string
let previousEnv: string | undefined

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'airy-discovery-'))
})

afterAll(async () => {
  await rm(dir, { recursive: true, force: true })
})

beforeEach(() => {
  previousEnv = process.env[SIDECAR_ENV]
})

afterEach(() => {
  if (previousEnv === undefined) delete process.env[SIDECAR_ENV]
  else process.env[SIDECAR_ENV] = previousEnv
})

describe('repo root detection', () => {
  it('walks up from a nested package dir to the checkout root', () => {
    // this test file lives at <repo>/packages/mcp-server/tests
    const root = repoRootOf(TESTS_DIR)
    expect(root).toBe(join(TESTS_DIR, '..', '..', '..'))
    expect(repoRootOf(join(root!, 'apps', 'sheets', 'src'))).toBe(root)
  })

  it('returns null outside a checkout', () => {
    expect(repoRootOf(dir)).toBeNull()
  })
})

describe('sidecar candidate resolution', () => {
  it('puts the env override first', async () => {
    process.env[SIDECAR_ENV] = join(dir, 'custom-sidecar')
    // default fromDir is this module inside the checkout, so the repo build
    // path is offered as the second candidate
    const candidates = await sidecarCandidates()
    expect(candidates[0]).toBe(join(dir, 'custom-sidecar'))
    expect(candidates[1]).toBe(
      join(
        TESTS_DIR,
        '..',
        '..',
        '..',
        'apps',
        'sheets',
        'native',
        'xlsx-engine',
        'target',
        'release',
        process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar',
      ),
    )
  })

  it('resolves an existing binary from the env override', async () => {
    const binary = join(dir, 'custom-sidecar')
    await writeFile(binary, 'stub')
    process.env[SIDECAR_ENV] = binary
    await expect(findSidecarBinary(dir)).resolves.toBe(binary)
  })

  it('returns null when nothing exists (clear error available)', async () => {
    delete process.env[SIDECAR_ENV]
    await expect(findSidecarBinary(dir)).resolves.toBeNull()
    const error = sidecarMissingError()
    expect(error.message).toContain('native:build')
    expect(error.message).toContain(SIDECAR_ENV)
  })
})
