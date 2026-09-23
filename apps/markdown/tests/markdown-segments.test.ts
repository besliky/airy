import { describe, expect, it } from 'vitest'
import { splitBodyForHydration } from '../src/renderer/markdown/segments'

const paras = (n: number, from = 0): string =>
  Array.from({ length: n }, (_, i) => `Paragraph ${from + i} text`).join('\n\n') + '\n'

describe('splitBodyForHydration', () => {
  it('keeps small bodies in a single segment', () => {
    const body = paras(10)
    expect(splitBodyForHydration(body)).toEqual([body])
  })

  it('concatenating the segments reproduces the body byte for byte', () => {
    const body =
      paras(30) +
      '# Heading\n\n' +
      paras(30, 100) +
      '```\ncode with\n\nblank lines\n```\n\n' +
      paras(30, 200) +
      '- a\n\n- b\n\n- c\n\n' +
      paras(30, 300)
    const segments = splitBodyForHydration(body, { maxLines: 10, maxChars: 256 })
    expect(segments.length).toBeGreaterThan(3)
    expect(segments.join('')).toBe(body)
  })

  it('never cuts inside a fenced code block that contains blank lines', () => {
    const fence = '```\nline one\n\nline two\n\nline three\n```\n\n'
    const body = fence.repeat(10)
    // every candidate boundary falls inside the fence — one segment is the
    // only safe answer even though the budget is tiny
    expect(splitBodyForHydration(body, { maxLines: 2, maxChars: 32 })).toEqual([body])
  })

  it('never cuts inside a loose list (blank lines between items)', () => {
    const list = Array.from({ length: 10 }, (_, i) => `- item ${i}`).join('\n\n') + '\n\n'
    const body = list + paras(10)
    const segments = splitBodyForHydration(body, { maxLines: 4, maxChars: 64 })
    // the loose list may not be split even though it exceeds the budget:
    // the first segment holds the whole list (minus the separator blank run,
    // which leads the next segment)
    expect(segments[0]).toBe(list.slice(0, -1))
    expect(segments.join('')).toBe(body)
  })

  it('does not treat an indented continuation after a blank line as a boundary', () => {
    const item = '- item with continuation\n\n  continued text\n\n'
    const body = item.repeat(10)
    expect(splitBodyForHydration(body, { maxLines: 4, maxChars: 64 })).toEqual([body])
  })

  it('never cuts inside raw HTML blocks or comments containing blank lines', () => {
    const html = '<pre>\nraw\n\nlines\n</pre>\n\n<!-- comment\n\nspanning -->\n\n'
    const body = html.repeat(10)
    expect(splitBodyForHydration(body, { maxLines: 2, maxChars: 32 })).toEqual([body])
  })

  it('keeps bodies with link reference definitions whole', () => {
    const body = '[ref]: https://example.com\n\n' + paras(50) + 'See [ref][1] everywhere.\n'
    const def = '[1]: https://example.com\n\n'
    // any reference definition anywhere disables segmentation
    expect(splitBodyForHydration(body, { maxLines: 4 })).toEqual([body])
    expect(splitBodyForHydration(def + paras(50), { maxLines: 4 })).toEqual([def + paras(50)])
  })

  it('caps segment length at maxLines at safe boundaries', () => {
    const body = paras(200)
    const segments = splitBodyForHydration(body, { maxLines: 20, maxChars: 1 << 20 })
    expect(segments.length).toBeGreaterThanOrEqual(10)
    for (const segment of segments.slice(0, -1)) {
      // every non-last segment ends at a block boundary: one newline after the
      // block, never a blank run (that leads the next segment)
      expect(segment.endsWith('\n\n')).toBe(false)
      expect(segment.endsWith('\n')).toBe(true)
    }
    expect(segments.join('')).toBe(body)
  })

  it('handles an empty body and a body without trailing newline', () => {
    expect(splitBodyForHydration('')).toEqual([''])
    const noTrailing = paras(10).trimEnd()
    const segments = splitBodyForHydration(noTrailing, { maxLines: 2 })
    expect(segments.join('')).toBe(noTrailing)
  })
})
