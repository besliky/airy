import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { uniqueGeneratedAttachmentPath, uniqueGeneratedPdfPath } from '../src/main/generated-output'

describe('uniqueGeneratedPdfPath', () => {
  it('keeps generated PDFs inside the configured directory', () => {
    expect(uniqueGeneratedPdfPath('/save', '../report-merged.pdf', () => false)).toBe(
      join('/save', 'report-merged.pdf'),
    )
  })

  it('adds a PDF extension and skips existing names', () => {
    const occupied = new Set([join('/save', 'report.pdf'), join('/save', 'report-2.pdf')])
    expect(uniqueGeneratedPdfPath('/save', 'report', (path) => occupied.has(path))).toBe(
      join('/save', 'report-3.pdf'),
    )
  })

  it('sanitizes characters that are invalid in file names', () => {
    expect(uniqueGeneratedPdfPath('/save', 'a:b?.pdf', () => false)).toBe(join('/save', 'a_b_.pdf'))
  })
})

describe('uniqueGeneratedAttachmentPath (UX-1734)', () => {
  it('keeps a plausible extension and stays inside the save directory', () => {
    expect(uniqueGeneratedAttachmentPath('/save', '../../docs/report.pdf', () => false)).toBe(
      join('/save', 'report.pdf'),
    )
  })

  it('flattens invalid characters in the basename, keeping the extension', () => {
    expect(uniqueGeneratedAttachmentPath('/save', 'a/b:c?.txt', () => false)).toBe(
      join('/save', 'b_c_.txt'),
    )
  })

  it('skips existing names by suffixing before the extension', () => {
    const occupied = new Set([join('/save', 'doc.pdf'), join('/save', 'doc-2.pdf')])
    expect(uniqueGeneratedAttachmentPath('/save', 'doc.pdf', (path) => occupied.has(path))).toBe(
      join('/save', 'doc-3.pdf'),
    )
  })

  it('drops absurd extensions but keeps a recognizable name', () => {
    // no extension at all
    expect(uniqueGeneratedAttachmentPath('/save', 'README', () => false)).toBe(
      join('/save', 'README'),
    )
    // an overlong "extension" stays part of the name as-is
    expect(uniqueGeneratedAttachmentPath('/save', 'payload.notanextension', () => false)).toBe(
      join('/save', 'payload.notanextension'),
    )
    // a dotfile has no extension to preserve
    expect(uniqueGeneratedAttachmentPath('/save', '.hidden', () => false)).toBe(
      join('/save', '.hidden'),
    )
  })

  it('falls back to a safe name when nothing usable remains', () => {
    expect(uniqueGeneratedAttachmentPath('/save', '', () => false)).toBe(
      join('/save', 'attachment'),
    )
    expect(uniqueGeneratedAttachmentPath('/save', '???', () => false)).toBe(join('/save', '___'))
  })
})
