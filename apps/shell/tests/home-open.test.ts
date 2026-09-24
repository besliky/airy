import { beforeEach, afterAll, describe, expect, it, vi } from 'vitest'
import { mkdtempSync, mkdirSync, rmSync, writeFileSync, statSync, accessSync } from 'node:fs'
import type { PathLike } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { friendlyErrorKey } from '../src/shared/error-codes'
import { openFromHome } from '../src/main/home-open'
import { classifyOpenFailure, type OpenFailureDeps } from '../src/main/open-failure'

/**
 * openFromHome (src/main/home-open.ts): clicking a row in Home used to end
 * in total silence when the path could not be opened — routing returned
 * false, the IPC handler dropped the result, and no tab, dialog or toast
 * ever appeared (BUG-1677). These tests drive the real composition (the
 * #154 classify/report pair behind openFromHome) against the real
 * filesystem; the one deterministic fake is the accessSync probe so the
 * chmod-000 EACCES shape does not depend on the running user. "No tab" is
 * asserted via the return value: the wired openDocument mirrors the
 * routeDocumentPath guard (a tab is only attempted for a path that passes
 * the classify preflight) and its false is exactly the dropped result the
 * bug was about.
 */

const fsReal = await vi.importActual<typeof import('node:fs')>('node:fs')

vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return {
    ...actual,
    statSync: vi.fn(actual.statSync),
    accessSync: vi.fn(actual.accessSync),
  }
})

function errnoError(code: string, message: string): Error {
  return Object.assign(new Error(`${code}: ${message}`), { code })
}

let root: string

beforeEach(() => {
  // restore the real fs behavior; individual tests layer one-shot overrides
  vi.mocked(statSync).mockImplementation(((path: PathLike) =>
    fsReal.statSync(path)) as typeof statSync)
  vi.mocked(accessSync).mockImplementation(((path: PathLike) =>
    fsReal.accessSync(path)) as typeof accessSync)
  root = mkdtempSync(join(tmpdir(), 'bug1677-'))
})

afterAll(() => {
  rmSync(root, { recursive: true, force: true })
})

/** the production wiring under test: deps shaped like index.ts's */
function wiredDeps() {
  const showErrorDialog = vi.fn<(message: string, err: Error) => void>()
  const openFailedMessage = vi.fn<(name: string) => string>((name) => `Could not open “${name}”`)
  return {
    deps: {
      // a tab is only worth opening for a path that passes the preflight —
      // the exact guard routeDocumentPath applies before opening a tab
      openDocument: (path: string) => classifyOpenFailure(path) === null,
      openFailure: { showErrorDialog, openFailedMessage } satisfies OpenFailureDeps,
    },
    showErrorDialog,
    openFailedMessage,
  }
}

describe('openFromHome (BUG-1677)', () => {
  it('an unreadable file raises the dialog channel and opens no tab', () => {
    const file = join(root, 'eaccess.md')
    writeFileSync(file, 'x')
    // the guard and the final report each run the probe: keep the failure
    // active for the whole test, not just one call
    vi.mocked(accessSync).mockImplementation((path: PathLike) => {
      if (path === file) throw errnoError('EACCES', `permission denied, access '${file}'`)
      return fsReal.accessSync(path)
    })
    const { deps, showErrorDialog, openFailedMessage } = wiredDeps()
    expect(openFromHome(file, deps)).toBe(false) // no tab
    expect(showErrorDialog).toHaveBeenCalledTimes(1)
    expect(openFailedMessage).toHaveBeenCalledWith('eaccess.md')
    const [, err] = showErrorDialog.mock.calls[0]
    expect((err as { code?: string }).code).toBe('EACCES')
    expect(friendlyErrorKey(err)).toBe('eperm')
  })

  it('a valid file opens as before and no dialog is raised', () => {
    const file = join(root, 'valid.docx')
    writeFileSync(file, 'x')
    const { deps, showErrorDialog } = wiredDeps()
    expect(openFromHome(file, deps)).toBe(true) // opened
    expect(showErrorDialog).not.toHaveBeenCalled()
  })

  it('a directory wearing a document extension raises the folder-message dialog', () => {
    const dir = join(root, 'fake.docx')
    mkdirSync(dir)
    const { deps, showErrorDialog } = wiredDeps()
    expect(openFromHome(dir, deps)).toBe(false) // no tab
    expect(showErrorDialog).toHaveBeenCalledTimes(1)
    const [, err] = showErrorDialog.mock.calls[0]
    expect((err as { code?: string }).code).toBe('EISDIR')
    expect(friendlyErrorKey(err)).toBe('eisdir') // “This is a folder…” per #154
  })

  it('a vanished path raises the ENOENT dialog instead of staying silent', () => {
    const { deps, showErrorDialog } = wiredDeps()
    expect(openFromHome(join(root, 'gone.md'), deps)).toBe(false)
    expect(showErrorDialog).toHaveBeenCalledTimes(1)
    const [, err] = showErrorDialog.mock.calls[0]
    expect((err as { code?: string }).code).toBe('ENOENT')
  })

  it('ignores a non-string payload without touching either channel', () => {
    const { deps, showErrorDialog } = wiredDeps()
    expect(openFromHome(42, deps)).toBe(false)
    expect(openFromHome(undefined, deps)).toBe(false)
    expect(showErrorDialog).not.toHaveBeenCalled()
  })

  it('stays silent when routing failed for a reason the preflight cannot see', () => {
    // e.g. no focused window yet: openDocument returns false for a healthy
    // file — the caller keeps its Home fallback, no dialog (same contract
    // as the launch paths)
    const file = join(root, 'fine.md')
    writeFileSync(file, 'x')
    const showErrorDialog = vi.fn<(message: string, err: Error) => void>()
    expect(
      openFromHome(file, {
        openDocument: () => false,
        openFailure: { showErrorDialog, openFailedMessage: (n) => n },
      }),
    ).toBe(false)
    expect(showErrorDialog).not.toHaveBeenCalled()
  })
})
