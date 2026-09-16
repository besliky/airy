import { beforeEach, describe, expect, it } from 'vitest'
import {
  grantRendererDir,
  grantRendererFileAccess,
  grantedRendererDirs,
  isPathInsideDir,
  pathIsInsideAny,
  rendererMayReadPath,
  resetRendererFileGrants,
} from '../src/renderer-file-access'

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

describe('renderer read allowlist', () => {
  it('denies everything before a grant, then allows the granted subtree', () => {
    expect(rendererMayReadPath('/home/u/docs/a.docx')).toBe(false)
    grantRendererFileAccess('/home/u/docs/a.docx')
    expect(rendererMayReadPath('/home/u/docs/a.docx')).toBe(true)
    expect(rendererMayReadPath('/home/u/docs/other.txt')).toBe(true)
    expect(rendererMayReadPath('/home/u/docs/sub/x.png')).toBe(true)
    expect(rendererMayReadPath('/home/u/secret.txt')).toBe(false)
    expect(rendererMayReadPath('/etc/passwd')).toBe(false)
  })

  it('never grants the filesystem root or relative paths', () => {
    grantRendererDir('/')
    grantRendererDir('')
    expect(grantedRendererDirs()).toEqual([])
    expect(rendererMayReadPath('relative/file.txt')).toBe(false)
    expect(rendererMayReadPath('')).toBe(false)
  })

  it('evicts the oldest grant once the bound is exceeded', () => {
    for (let i = 0; i < 70; i++) grantRendererDir(`/tmp/grant-${i}`)
    const dirs = grantedRendererDirs()
    expect(dirs.length).toBeLessThanOrEqual(64)
    expect(dirs).not.toContain('/tmp/grant-0')
    expect(dirs).toContain('/tmp/grant-69')
  })
})
