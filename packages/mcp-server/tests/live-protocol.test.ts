// Unit tests for the live bridge wire codec (src/live/protocol.ts): the
// NdjsonFramer must frame on raw bytes and decode only complete lines, so a
// socket chunk boundary splitting a multi-byte UTF-8 sequence never turns into
// U+FFFD (BUG-401). Pure Node — mirrors apps/shell/tests/bridge-protocol.test.ts
// for the client-side twin of the framer.
import { describe, expect, it, vi } from 'vitest'

import { NdjsonFramer } from '../src/live/protocol.js'

describe('NdjsonFramer (live protocol)', () => {
  it('accumulates chunks split mid-line', () => {
    const framer = new NdjsonFramer()
    expect(framer.push('{"meth')).toEqual({ lines: [], overflow: false })
    expect(framer.push('od":"pi')).toEqual({ lines: [], overflow: false })
    const result = framer.push('ng"}\n')
    expect(result.lines).toEqual(['{"method":"ping"}'])
    expect(result.overflow).toBe(false)
  })

  it('returns several completed lines per chunk and keeps the tail buffered', () => {
    const framer = new NdjsonFramer()
    const result = framer.push('a\nb\r\nc\nd')
    expect(result.lines).toEqual(['a', 'b', 'c'])
    expect(framer.push('\n').lines).toEqual(['d'])
  })

  it('accepts buffers', () => {
    const framer = new NdjsonFramer()
    expect(framer.push(Buffer.from('x\n')).lines).toEqual(['x'])
  })

  it('flags a line over the cap as overflow', () => {
    const framer = new NdjsonFramer(8)
    const result = framer.push('0123456789abcdef')
    expect(result.overflow).toBe(true)
    expect(result.lines).toEqual([])
  })

  it('counts the cap in bytes even for multibyte tails', () => {
    const framer = new NdjsonFramer(2)
    // U+4E2D is 3 bytes on the wire; a 2-byte cap must already trip
    expect(framer.push(Buffer.from('中', 'utf8'))).toEqual({ lines: [], overflow: true })
  })

  it('reassembles 中 (E4 B8 AD) split after any byte boundary', () => {
    const bytes = Buffer.from('中\n', 'utf8')
    expect([...bytes]).toEqual([0xe4, 0xb8, 0xad, 0x0a]) // guard the encoding assumption
    for (let cut = 1; cut < bytes.length; cut += 1) {
      const framer = new NdjsonFramer()
      expect(framer.push(bytes.subarray(0, cut))).toEqual({ lines: [], overflow: false })
      const result = framer.push(bytes.subarray(cut))
      expect(result.lines).toEqual(['中'])
      expect(result.overflow).toBe(false)
    }
  })

  it('reassembles a 4-byte emoji (surrogate pair) split after any byte boundary', () => {
    const bytes = Buffer.from('😀\n', 'utf8')
    expect([...bytes]).toEqual([0xf0, 0x9f, 0x98, 0x80, 0x0a])
    for (let cut = 1; cut < bytes.length; cut += 1) {
      const framer = new NdjsonFramer()
      expect(framer.push(bytes.subarray(0, cut))).toEqual({ lines: [], overflow: false })
      const result = framer.push(bytes.subarray(cut))
      expect(result.lines).toEqual(['😀'])
    }
  })

  it('reassembles a split character across a string chunk and a buffer chunk', () => {
    const framer = new NdjsonFramer()
    framer.push('prefix-')
    expect(framer.push(Buffer.from('中\n', 'utf8')).lines).toEqual(['prefix-中'])
  })

  it('keeps a CJK payload intact inside a full response line split mid-character', () => {
    const line = '{"ok":true,"result":{"context":{"selection":"<sel>中文</sel>"}}}'
    const bytes = Buffer.from(`${line}\n`, 'utf8')
    const lead = bytes.indexOf(0xe4)
    expect(lead).toBeGreaterThan(0)
    // offsets 1 and 2 cut inside U+4E2D; 0 (right before the lead byte) is the
    // character-aligned control case
    for (const offset of [0, 1, 2]) {
      const framer = new NdjsonFramer()
      framer.push(bytes.subarray(0, lead + offset))
      const { lines } = framer.push(bytes.subarray(lead + offset))
      expect(lines).toEqual([line])
      expect(JSON.parse(lines[0] ?? '').result.context.selection).toBe('<sel>中文</sel>')
    }
  })

  it('frames several multibyte lines in one chunk and keeps the tail buffered', () => {
    const framer = new NdjsonFramer()
    const result = framer.push(Buffer.from('中\n文\nx', 'utf8'))
    expect(result.lines).toEqual(['中', '文'])
    expect(framer.push(Buffer.from('\n', 'utf8')).lines).toEqual(['x'])
  })

  it('emits a line when the chunk ends exactly on the newline', () => {
    const framer = new NdjsonFramer()
    expect(framer.push(Buffer.from('中\n', 'utf8')).lines).toEqual(['中'])
    // the next chunk starts a fresh line — no residue from the previous one
    expect(framer.push(Buffer.from('文\n', 'utf8')).lines).toEqual(['文'])
  })

  it('emits empty lines as empty strings', () => {
    const framer = new NdjsonFramer()
    expect(framer.push(Buffer.from('\n\n中\n', 'utf8')).lines).toEqual(['', '', '中'])
  })

  it('strips a CR before the LF after multibyte content', () => {
    const framer = new NdjsonFramer()
    expect(framer.push(Buffer.from('中\r\n', 'utf8')).lines).toEqual(['中'])
  })

  it('strips a CR that lands in the previous chunk before a chunk-leading LF', () => {
    const framer = new NdjsonFramer()
    expect(framer.push(Buffer.from('x\r', 'utf8'))).toEqual({ lines: [], overflow: false })
    expect(framer.push(Buffer.from('\ny', 'utf8'))).toEqual({ lines: ['x'], overflow: false })
    expect(framer.push(Buffer.from('\n', 'utf8')).lines).toEqual(['y'])
  })

  it('keeps a dribbled unterminated tail without recopying it per push', () => {
    // The old implementation rebuilt the pending buffer on every push
    // (Buffer.concat of the whole tail), making a slow-dripped line O(n²) in
    // copied bytes. The tail must accumulate as chunks; only a completed
    // line may pay for a concat.
    const concat = vi.spyOn(Buffer, 'concat')
    const framer = new NdjsonFramer()
    for (let i = 0; i < 2_000; i += 1) framer.push('x')
    expect(concat).not.toHaveBeenCalled()
    // completing the line concatenates exactly its own chunks, once
    const result = framer.push('\n')
    expect(result.lines).toEqual(['x'.repeat(2_000)])
    expect(result.overflow).toBe(false)
    expect(concat).toHaveBeenCalledTimes(1)
    // a line that fits inside one chunk decodes without any concat
    framer.push('done\n')
    expect(concat).toHaveBeenCalledTimes(1)
    concat.mockRestore()
  })

  it('coalesces a 1-byte dribble into a bounded number of live chunks', () => {
    // BUG-903: the chunk list kept every socket chunk as its own Buffer, so
    // a dribbled line amplified memory ~100x — an 8MB line sent one byte at
    // a time stayed within the byte cap but amounted to millions of Buffer
    // objects. The dribble must coalesce into a bounded tail structure and
    // still decode exactly.
    const framer = new NdjsonFramer()
    const pushes = 100_000
    for (let i = 0; i < pushes; i += 1) framer.push('x')
    const liveChunks = (framer as unknown as { chunks: Buffer[] }).chunks
    expect(liveChunks.length).toBeLessThanOrEqual(2)
    const result = framer.push('\n')
    expect(result.lines).toEqual(['x'.repeat(pushes)])
    expect(result.overflow).toBe(false)
  })

  it('frames a line that mixes dribbled and by-reference chunks', () => {
    // the scratch tail, a chunk stored by reference (over the coalesce
    // limit), and a fresh scratch after it must concatenate as one line
    const framer = new NdjsonFramer()
    for (let i = 0; i < 5; i += 1) framer.push('ab')
    framer.push(Buffer.alloc(70 * 1024, 0x62))
    for (let i = 0; i < 5; i += 1) framer.push('cd')
    const result = framer.push('\n')
    expect(result.lines.length).toBe(1)
    const line = result.lines[0] ?? ''
    expect(line.length).toBe(10 + 70 * 1024 + 10)
    expect(line.startsWith('ababababab')).toBe(true)
    expect(line.endsWith('cdcdcdcdcd')).toBe(true)
  })

  it('flags a completed line over the cap even within a single chunk', () => {
    // Net chunks stay small, but push() itself may be handed one huge
    // buffer: a whole 10-byte line inside one chunk must trip an 8-byte cap
    // instead of slipping past the pending-tail-only check.
    const framer = new NdjsonFramer(8)
    const result = framer.push('0123456789\n')
    expect(result.overflow).toBe(true)
    expect(result.lines).toEqual([])
  })
})
