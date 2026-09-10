// Optional LibreOffice integration for legacy/ODF text formats (Phase 3b,
// ADR-9): .doc/.odt are converted to .docx in a temp dir for full editing,
// and format:'origin' saves export back through the same tool. LibreOffice is
// MPL-2.0 and invoked as an unmodified subprocess, which imposes no
// obligations on this code (SB9 section 5; same pattern as markitdown and
// libreoffice-convert).
//
// Every invocation gets its own -env:UserInstallation profile directory so
// concurrent conversions cannot fight over the per-user LibreOffice profile
// lock. The binary is looked up per call and the resolved path is cached.
import { spawn } from 'node:child_process'
import { access, constants, mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

export const SOFFICE_ENV = 'AIRY_SOFFICE'

const CONVERT_TIMEOUT_MS = 120_000

export interface SofficeTool {
  readonly binaryPath: string
}

/** Canonical LibreOffice export filters (SB9 convert-filter table). */
export const SOFFICE_FILTERS = {
  docx: 'MS Word 2007 XML',
  doc: 'MS Word 97',
  odt: 'writer8',
  ods: 'calc8',
} as const

let cachedTool: SofficeTool | null = null

async function isExecutable(path: string): Promise<boolean> {
  try {
    await access(path, constants.X_OK)
    return true
  } catch {
    return false
  }
}

/** PATH entries + platform install locations, in priority order. */
export async function sofficeCandidates(): Promise<string[]> {
  const candidates: string[] = []
  const fromEnv = process.env[SOFFICE_ENV]
  if (fromEnv && fromEnv.trim() !== '') candidates.push(fromEnv.trim())

  for (const dir of (process.env.PATH ?? '').split(':')) {
    if (dir.trim() === '') continue
    candidates.push(join(dir, 'soffice'), join(dir, 'soffice.com'))
  }
  if (process.platform === 'win32') {
    const programFiles = process.env.PROGRAMFILES ?? 'C:\\Program Files'
    const programFilesX86 = process.env['PROGRAMFILES(X86)'] ?? 'C:\\Program Files (x86)'
    for (const root of [programFiles, programFilesX86]) {
      candidates.push(
        join(root, 'LibreOffice', 'program', 'soffice.com'),
        join(root, 'LibreOffice', 'program', 'soffice.exe'),
      )
    }
  } else if (process.platform === 'darwin') {
    candidates.push('/Applications/LibreOffice.app/Contents/MacOS/soffice')
  } else {
    candidates.push('/usr/bin/soffice', '/usr/local/bin/soffice', '/snap/bin/soffice')
  }
  return candidates
}

/**
 * Find a usable soffice binary. Successful lookups are cached for the
 * process lifetime; failures are re-checked on the next call so installing
 * LibreOffice (or fixing AIRY_SOFFICE) takes effect without a restart.
 */
export async function findSoffice(): Promise<SofficeTool | null> {
  if (cachedTool) return cachedTool
  for (const candidate of await sofficeCandidates()) {
    if (await isExecutable(candidate)) {
      cachedTool = { binaryPath: candidate }
      return cachedTool
    }
  }
  return null
}

/** Drop the cached lookup (tests force fresh discovery through this). */
export function resetSofficeCache(): void {
  cachedTool = null
}

/** Agent-facing error when LibreOffice is required but absent. */
export function sofficeMissingError(reason: string): Error {
  return new Error(
    `${reason} LibreOffice is not installed (or not on PATH). Install LibreOffice ` +
      `(https://www.libreoffice.org) or point the ${SOFFICE_ENV} env var at the soffice binary, then retry.`,
  )
}

export interface SofficeConvertOptions {
  /** LibreOffice export filter name (SOFFICE_FILTERS). */
  readonly filter: string
  /** Output extension ("docx", "odt", ...) — also the --convert-to target. */
  readonly extension: string
  /** Existing directory that receives the converted file. */
  readonly outDir: string
  readonly timeoutMs?: number
}

function baseNameWithoutExtension(path: string): string {
  const segments = path.split(/[\\/]/)
  const name = segments.pop() ?? path
  const dot = name.lastIndexOf('.')
  return dot > 0 ? name.slice(0, dot) : name
}

/**
 * Convert one file through `soffice --headless --convert-to ext:filter`.
 * Returns the produced output path inside outDir. Fails with the collected
 * stderr when soffice exits non-zero or produces no output file (some builds
 * exit 0 even when a filter fails, so existence is checked too).
 */
export async function convertViaSoffice(
  tool: SofficeTool,
  inputPath: string,
  options: SofficeConvertOptions,
): Promise<string> {
  const profileDir = await mkdtemp(join(tmpdir(), 'airy-soffice-profile-'))
  try {
    const args = [
      '--headless',
      '--norecover',
      `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
      '--convert-to',
      `${options.extension}:${options.filter}`,
      '--outdir',
      options.outDir,
      inputPath,
    ]
    return await new Promise<string>((resolve, reject) => {
      const child = spawn(tool.binaryPath, args, {
        stdio: ['ignore', 'pipe', 'pipe'],
        windowsHide: true,
      })
      let stderr = ''
      let settled = false
      const timeout = setTimeout(() => {
        if (settled) return
        settled = true
        child.kill()
        reject(
          new Error(
            `LibreOffice conversion timed out after ${String(options.timeoutMs ?? CONVERT_TIMEOUT_MS)} ms.`,
          ),
        )
      }, options.timeoutMs ?? CONVERT_TIMEOUT_MS)
      const fail = (message: string) => {
        if (settled) return
        settled = true
        clearTimeout(timeout)
        reject(new Error(message))
      }
      child.stderr?.setEncoding('utf8')
      child.stderr?.on('data', (chunk: string) => {
        stderr = `${stderr}${chunk}`.slice(-2048)
      })
      // stdout of soffice goes to the convert log; drain it so the pipe
      // cannot fill up and block the child
      child.stdout?.resume()
      child.once('error', (error) => {
        fail(`LibreOffice failed to start (${tool.binaryPath}): ${error.message}`)
      })
      child.once('close', (code) => {
        if (settled) return
        const output = join(
          options.outDir,
          `${baseNameWithoutExtension(inputPath)}.${options.extension}`,
        )
        access(output, constants.F_OK)
          .then(() => {
            if (settled) return
            settled = true
            clearTimeout(timeout)
            resolve(output)
          })
          .catch(() => {
            if (settled) return
            settled = true
            clearTimeout(timeout)
            const detail = stderr.trim()
            reject(
              new Error(
                `LibreOffice conversion to .${options.extension} produced no output` +
                  `${code !== 0 ? ` (exit code ${String(code)})` : ''}${detail ? `: ${detail}` : '.'}`,
              ),
            )
          })
      })
    })
  } finally {
    await rm(profileDir, { recursive: true, force: true })
  }
}
