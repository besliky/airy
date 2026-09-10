// LibreOffice discovery + conversion wrapper: env override, PATH/platform
// candidates, and the spawn surface (headless, isolated UserInstallation
// profile, --convert-to ext:filter, --outdir) exercised against a fake
// soffice implemented as an executable Node script. Timeout and no-output
// failures covered too. No real LibreOffice needed.
import { chmod, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it } from 'vitest'

import {
  convertViaSoffice,
  findSoffice,
  resetSofficeCache,
  SOFFICE_ENV,
  sofficeCandidates,
  sofficeMissingError,
} from '../src/import/soffice.js'

// Fake soffice: records argv, creates <outdir>/<input-base>.<ext>; can be
// told to fail (exit 3) or hang (never reply) through env vars.
const FAKE_SOFFICE = `#!/usr/bin/env node
const fs = require('node:fs')
const path = require('node:path')
const args = process.argv.slice(2)
if (process.env.ARGS_OUT) fs.appendFileSync(process.env.ARGS_OUT, JSON.stringify(args) + '\\n')
if (process.env.FAKE_SOFFICE_HANG === '1') {
  setInterval(() => {}, 1000)
} else {
  if (process.env.FAKE_SOFFICE_FAIL !== '1') {
    const outDir = args[args.indexOf('--outdir') + 1]
    const input = args[args.length - 1]
    const base = path.basename(input).replace(/\\.[^.]+$/, '')
    const ext = args[args.indexOf('--convert-to') + 1].split(':')[0]
    fs.writeFileSync(path.join(outDir, base + '.' + ext), 'converted-content')
  }
  process.exit(process.env.FAKE_SOFFICE_FAIL === '1' ? 3 : 0)
}
`

let dir: string
let fakeSoffice: string
let argsDump: string
let previousEnv: string | undefined

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'airy-soffice-test-'))
  fakeSoffice = join(dir, 'soffice')
  await writeFile(fakeSoffice, FAKE_SOFFICE, 'utf8')
  await chmod(fakeSoffice, 0o755)
  argsDump = join(dir, 'args.jsonl')
})

afterAll(async () => {
  await rm(dir, { recursive: true, force: true })
})

beforeEach(() => {
  previousEnv = process.env[SOFFICE_ENV]
  process.env[SOFFICE_ENV] = fakeSoffice
  process.env.ARGS_OUT = argsDump
  resetSofficeCache()
})

afterEach(async () => {
  if (previousEnv === undefined) delete process.env[SOFFICE_ENV]
  else process.env[SOFFICE_ENV] = previousEnv
  delete process.env.ARGS_OUT
  delete process.env.FAKE_SOFFICE_FAIL
  delete process.env.FAKE_SOFFICE_HANG
  resetSofficeCache()
  await writeFile(argsDump, '')
})

describe('soffice discovery', () => {
  it('resolves the env override and caches it', async () => {
    const tool = await findSoffice()
    expect(tool?.binaryPath).toBe(fakeSoffice)
  })

  it('falls back to PATH and platform install locations', async () => {
    process.env[SOFFICE_ENV] = '/nonexistent/soffice'
    resetSofficeCache()
    const candidates = await sofficeCandidates()
    expect(candidates[0]).toBe('/nonexistent/soffice')
    if (process.platform === 'linux') {
      expect(candidates).toContain('/usr/bin/soffice')
    }
    if (process.platform === 'darwin') {
      expect(candidates).toContain('/Applications/LibreOffice.app/Contents/MacOS/soffice')
    }
    if (process.platform === 'win32') {
      expect(candidates.some((candidate) => candidate.endsWith('soffice.com'))).toBe(true)
    }
  })

  it('produces an actionable error when missing', () => {
    const error = sofficeMissingError('Editing .odt failed.')
    expect(error.message).toContain('Editing .odt failed.')
    expect(error.message).toContain('Install LibreOffice')
    expect(error.message).toContain(SOFFICE_ENV)
  })
})

describe('convertViaSoffice against a fake soffice', () => {
  it('invokes headless with an isolated profile, filter and outdir', async () => {
    const tool = { binaryPath: fakeSoffice }
    const outDir = join(dir, 'out-1')
    await mkdir(outDir, { recursive: true })
    const input = join(dir, 'letter.docx')
    await writeFile(input, 'docx-bytes')
    const output = await convertViaSoffice(tool, input, {
      filter: 'MS Word 97',
      extension: 'doc',
      outDir,
    })
    expect(output).toBe(join(outDir, 'letter.doc'))
    expect(await readFile(output, 'utf8')).toBe('converted-content')
    const args = JSON.parse((await readFile(argsDump, 'utf8')).trim()) as string[]
    expect(args).toContain('--headless')
    expect(args).toContain('--convert-to')
    expect(args[args.indexOf('--convert-to') + 1]).toBe('doc:MS Word 97')
    expect(args[args.indexOf('--outdir') + 1]).toBe(outDir)
    expect(args[args.length - 1]).toBe(input)
    const profile = args.find((arg) => arg.startsWith('-env:UserInstallation='))
    expect(profile).toMatch(/-env:UserInstallation=file:\/\/\/.*airy-soffice-profile-/)
  })

  it('fails with the exit code when no output is produced', async () => {
    process.env.FAKE_SOFFICE_FAIL = '1'
    const outDir = join(dir, 'out-2')
    await mkdir(outDir, { recursive: true })
    const input = join(dir, 'broken.docx')
    await writeFile(input, 'docx-bytes')
    await expect(
      convertViaSoffice({ binaryPath: fakeSoffice }, input, {
        filter: 'writer8',
        extension: 'odt',
        outDir,
      }),
    ).rejects.toThrow(/produced no output \(exit code 3\)/)
  })

  it('times out hung conversions and kills the process', async () => {
    process.env.FAKE_SOFFICE_HANG = '1'
    const outDir = join(dir, 'out-3')
    await mkdir(outDir, { recursive: true })
    const input = join(dir, 'slow.docx')
    await writeFile(input, 'docx-bytes')
    await expect(
      convertViaSoffice({ binaryPath: fakeSoffice }, input, {
        filter: 'MS Word 2007 XML',
        extension: 'docx',
        outDir,
        timeoutMs: 500,
      }),
    ).rejects.toThrow(/timed out after 500 ms/)
  }, 10_000)
})
