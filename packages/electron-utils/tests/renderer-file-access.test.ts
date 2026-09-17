import { beforeEach, describe, expect, it } from 'vitest'
import {
  forgetRendererFileAccess,
  grantRendererDir,
  grantRendererFileAccess,
  grantedRendererDirs,
  isPathInsideDir,
  pathIsInsideAny,
  rendererMayReadPath,
  resetRendererFileGrants,
} from '../src/renderer-file-access'

/** two distinct renderer (webContents) ids, like two tabs of one shell */
const TAB_A = 11
const TAB_B = 22

beforeEach(() => {
  resetRendererFileGrants()
})

describe('isPathInsideDir', () => {
  it('accepts the dir itself and deeper paths, rejects siblings and prefixes', () => {
    expect(isPathInsideDir('/home/u/docs', '/home/u/docs')).toBe(true)
    expect(isPathInsideDir('/home/u/docs', '/home/u/docs/a/b.docx')).toBe(true)
    expect(isPathInsideDir('/home/u/docs', '/home/u/docs-other/b.docx')).toBe(false)
    expect(isPathInsideDir('/home/u/docs', '/home/u/docx')).toBe(false)
  })

  it('compares case-insensitively on win32 only', () => {
    expect(isPathInsideDir('C:\\Users\\u', 'c:\\users\\u\\a.docx', 'win32')).toBe(true)
    expect(isPathInsideDir('/home/u/docs', '/Home/U/Docs/a', 'linux')).toBe(false)
  })
})

describe('pathIsInsideAny', () => {
  it('matches any granted directory', () => {
    expect(pathIsInsideAny(['/a', '/b/c'], '/b/c/d.txt')).toBe(true)
    expect(pathIsInsideAny(['/a', '/b/c'], '/b/d.txt')).toBe(false)
    expect(pathIsInsideAny([], '/anything')).toBe(false)
  })
})

describe('renderer read allowlist (per sender)', () => {
  it('denies everything before a grant, then allows the granted subtree', () => {
    expect(rendererMayReadPath(TAB_A, '/home/u/docs/a.docx')).toBe(false)
    grantRendererFileAccess('/home/u/docs/a.docx', TAB_A)
    expect(rendererMayReadPath(TAB_A, '/home/u/docs/a.docx')).toBe(true)
    expect(rendererMayReadPath(TAB_A, '/home/u/docs/other.txt')).toBe(true)
    expect(rendererMayReadPath(TAB_A, '/home/u/docs/sub/x.png')).toBe(true)
    expect(rendererMayReadPath(TAB_A, '/home/u/secret.txt')).toBe(false)
    expect(rendererMayReadPath(TAB_A, '/etc/passwd')).toBe(false)
  })

  it('never grants the filesystem root or relative paths', () => {
    grantRendererDir('/', TAB_A)
    grantRendererDir('', TAB_A)
    expect(grantedRendererDirs(TAB_A)).toEqual([])
    expect(rendererMayReadPath(TAB_A, 'relative/file.txt')).toBe(false)
    expect(rendererMayReadPath(TAB_A, '')).toBe(false)
  })

  it('evicts the oldest grant once the bound is exceeded (per sender)', () => {
    for (let i = 0; i < 70; i++) grantRendererDir(`/tmp/grant-${i}`, TAB_A)
    const dirs = grantedRendererDirs(TAB_A)
    expect(dirs.length).toBeLessThanOrEqual(64)
    expect(dirs).not.toContain('/tmp/grant-0')
    expect(dirs).toContain('/tmp/grant-69')
  })

  it('isolates senders: a grant through tab A is unreadable by tab B', () => {
    grantRendererFileAccess('/home/u/docs/a.docx', TAB_A)
    expect(rendererMayReadPath(TAB_A, '/home/u/docs/a.docx')).toBe(true)
    expect(rendererMayReadPath(TAB_B, '/home/u/docs/a.docx')).toBe(false)
    expect(grantedRendererDirs(TAB_B)).toEqual([])
    // B granting its own dir widens nothing for A either
    grantRendererFileAccess('/home/u/pics/b.png', TAB_B)
    expect(rendererMayReadPath(TAB_A, '/home/u/pics/b.png')).toBe(false)
    expect(rendererMayReadPath(TAB_B, '/home/u/pics/b.png')).toBe(true)
  })

  it('forgets a sender entirely on teardown, without touching other senders', () => {
    grantRendererFileAccess('/home/u/docs/a.docx', TAB_A)
    grantRendererFileAccess('/home/u/docs/b.docx', TAB_B)
    forgetRendererFileAccess(TAB_A)
    expect(rendererMayReadPath(TAB_A, '/home/u/docs/a.docx')).toBe(false)
    expect(grantedRendererDirs(TAB_A)).toEqual([])
    expect(rendererMayReadPath(TAB_B, '/home/u/docs/b.docx')).toBe(true)
    // a fresh grant for the recycled webContents id starts empty
    grantRendererFileAccess('/tmp/new.txt', TAB_A)
    expect(grantedRendererDirs(TAB_A)).toEqual(['/tmp'])
  })

  it('bounds the number of tracked senders (stale senders evicted FIFO)', () => {
    for (let i = 0; i < 70; i++) {
      grantRendererFileAccess(`/home/u/f${i}/x.txt`, i + 1)
    }
    // the very first senders were evicted; the newest still holds its grant
    expect(rendererMayReadPath(1, '/home/u/f0/x.txt')).toBe(false)
    expect(rendererMayReadPath(70, '/home/u/f69/x.txt')).toBe(true)
  })
})
