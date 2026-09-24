import { beforeEach, afterAll, describe, expect, it, vi } from 'vitest'
import { mkdtempSync, mkdirSync, rmSync, writeFileSync, statSync, accessSync } from 'node:fs'
import type { PathLike } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { friendlyErrorKey } from '../src/shared/error-codes'
import {
  classifyOpenFailure,
  reportOpenFailure,
  resetOpenFailureDedupe,
} from '../src/main/open-failure'

/**
 * classifyOpenFailure / reportOpenFailure (src/main/open-failure.ts): an
 * intended open (CLI argument, double-click, open-file) of an unreadable path
 * used to end in silence — no dialog, no tab, the shell just fell back to the
 * home screen (BUG-1655). statSync/accessSync are wrapped in vi.fn so the two
 * EACCES shapes (chmod 000 file next to a readable directory, unreadable
 * parent directory) are simulated deterministically regardless of the running
 * user; every other case runs against the real filesystem.
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
  // the active-dialog dedupe is module state (BUG-1678): a leaked marker from
  // one test must never swallow the dialog of the next
  resetOpenFailureDedupe()
  root = mkdtempSync(join(tmpdir(), 'bug1655-'))
})

afterAll(() => {
  rmSync(root, { recursive: true, force: true })
})

describe('classifyOpenFailure', () => {
  it('a readable file passes the preflight: it opens as before', () => {
    const file = join(root, 'valid.docx')
    writeFileSync(file, 'x')
    expect(classifyOpenFailure(file)).toBeNull()
  })

  it('a vanished path classifies as ENOENT', () => {
    const err = classifyOpenFailure(join(root, 'gone.md'))
    expect(err).toBeInstanceOf(Error)
    expect((err as NodeJS.ErrnoException).code).toBe('ENOENT')
    expect(friendlyErrorKey(err)).toBe('enoent')
  })

  it('a directory wearing a document extension classifies as EISDIR', () => {
    const dir = join(root, 'deck.pptx')
    mkdirSync(dir)
    const err = classifyOpenFailure(dir)
    expect((err as { code?: string }).code).toBe('EISDIR')
    expect(friendlyErrorKey(err)).toBe('eisdir')
  })

  it('a chmod 000 file next to a readable directory classifies as EACCES via the read probe', () => {
    // stat() needs only +x on the parents, so it succeeds; the R_OK probe is
    // what catches the unreadable file — simulated deterministically here and
    // proven against a real chmod 000 file in the live check (BUG-1655)
    const file = join(root, 'secret.pdf')
    writeFileSync(file, 'x')
    vi.mocked(accessSync).mockImplementationOnce(() => {
      throw errnoError('EACCES', `permission denied, access '${file}'`)
    })
    const err = classifyOpenFailure(file)
    expect((err as { code?: string }).code).toBe('EACCES')
    expect(friendlyErrorKey(err)).toBe('eperm')
  })

  it('an unreadable parent directory classifies as EACCES from stat', () => {
    // existsSync already returned false for this shape (EACCES on the parent)
    vi.mocked(statSync).mockImplementationOnce(() => {
      throw errnoError('EACCES', `permission denied, stat '${root}/noaccess/f.md'`)
    })
    const err = classifyOpenFailure(join(root, 'noaccess', 'f.md'))
    expect((err as { code?: string }).code).toBe('EACCES')
    expect(friendlyErrorKey(err)).toBe('eperm')
  })

  it('an EPERM failure keeps its own code and maps to the same friendly key', () => {
    vi.mocked(statSync).mockImplementationOnce(() => {
      throw errnoError('EPERM', 'operation not permitted')
    })
    const err = classifyOpenFailure(join(root, 'locked.docx'))
    expect((err as { code?: string }).code).toBe('EPERM')
    expect(friendlyErrorKey(err)).toBe('eperm')
  })
})

describe('reportOpenFailure', () => {
  function deps() {
    return {
      showErrorDialog: vi.fn<(message: string, err: Error, onClosed?: () => void) => void>(),
      openFailedMessage: vi.fn<(name: string) => string>((name) => `Could not open “${name}”`),
    }
  }

  /** a vanished path: the cheapest deterministic failure for the dedupe tests */
  function missing(name: string): string {
    return join(root, name)
  }

  it.each(['EACCES', 'EPERM', 'ENOENT', 'EISDIR'])(
    'a %s failure raises the dialog channel, and no tab is created for it',
    (code) => {
      const target = join(root, `target-${code}.md`)
      if (code === 'EISDIR') {
        // a directory wearing a document extension
        mkdirSync(target)
      } else if (code !== 'ENOENT') {
        // ENOENT keeps the path absent; the others create a readable file and
        // inject the failure into the preflight
        writeFileSync(target, 'x')
        vi.mocked(statSync).mockImplementationOnce(() => {
          throw errnoError(code, `simulated ${code} for ${target}`)
        })
      }
      const d = deps()
      expect(reportOpenFailure(target, d)).toBe(true)
      expect(d.showErrorDialog).toHaveBeenCalledTimes(1)
      expect(d.openFailedMessage).toHaveBeenCalledWith(`target-${code}.md`)
      const [, err] = d.showErrorDialog.mock.calls[0]
      expect((err as { code?: string }).code).toBe(code)
      // every silent class resolves to a localized friendly text
      expect(friendlyErrorKey(err)).not.toBeNull()
    },
  )

  it('stays silent when the preflight sees nothing wrong (router failed elsewhere)', () => {
    const file = join(root, 'fine.md')
    writeFileSync(file, 'x')
    const d = deps()
    // e.g. no window yet: the caller keeps its home-tab fallback, no dialog
    expect(reportOpenFailure(file, d)).toBe(false)
    expect(d.showErrorDialog).not.toHaveBeenCalled()
  })

  describe('active-dialog dedupe (BUG-1678)', () => {
    it('N failures of the same path while its dialog is open raise exactly one dialog', () => {
      const target = missing('looped.docx')
      const d = deps()
      // the unpacked lock-retry loop re-broadcasts one forwarded open ~21x
      for (let i = 0; i < 21; i++) expect(reportOpenFailure(target, d)).toBe(true)
      expect(d.showErrorDialog).toHaveBeenCalledTimes(1)
      expect(d.openFailedMessage).toHaveBeenCalledTimes(1)
    })

    it('failures of different paths each raise their own dialog', () => {
      const d = deps()
      expect(reportOpenFailure(missing('a.docx'), d)).toBe(true)
      expect(reportOpenFailure(missing('b.pdf'), d)).toBe(true)
      expect(reportOpenFailure(missing('c.pptx'), d)).toBe(true)
      expect(d.showErrorDialog).toHaveBeenCalledTimes(3)
    })

    it('a dismissed dialog releases the path: the next failure shows again', () => {
      const target = missing('again.docx')
      const d = deps()
      let onClosed!: () => void
      d.showErrorDialog.mockImplementation((_m, _e, close) => {
        if (close) onClosed = close
      })
      expect(reportOpenFailure(target, d)).toBe(true)
      expect(d.showErrorDialog).toHaveBeenCalledTimes(1)
      // repeats while the dialog is up are swallowed…
      expect(reportOpenFailure(target, d)).toBe(true)
      expect(d.showErrorDialog).toHaveBeenCalledTimes(1)
      // …but the marker must not stick forever: dismissing the dialog lets a
      // later failure of the same path surface again
      onClosed()
      expect(reportOpenFailure(target, d)).toBe(true)
      expect(d.showErrorDialog).toHaveBeenCalledTimes(2)
    })

    it('dismissing one path never releases a still-open dialog of another', () => {
      const first = missing('first.docx')
      const second = missing('second.docx')
      const d = deps()
      const closers: Array<() => void> = []
      d.showErrorDialog.mockImplementation((_m, _e, close) => closers.push(close!))
      expect(reportOpenFailure(first, d)).toBe(true)
      expect(reportOpenFailure(second, d)).toBe(true)
      expect(d.showErrorDialog).toHaveBeenCalledTimes(2)
      closers[0]()
      expect(reportOpenFailure(first, d)).toBe(true)
      expect(d.showErrorDialog).toHaveBeenCalledTimes(3)
      expect(reportOpenFailure(second, d)).toBe(true)
      // second's dialog is still on screen — still deduped
      expect(d.showErrorDialog).toHaveBeenCalledTimes(3)
    })
  })
})
