// Unit tests for the live bridge wire codec (src/live/protocol.ts): the
// NdjsonFramer must frame on raw bytes and decode only complete lines, so a
// socket chunk boundary splitting a multi-byte UTF-8 sequence never turns into
// U+FFFD (BUG-401). Pure Node — mirrors apps/shell/tests/bridge-protocol.test.ts
// for the client-side twin of the framer.
import { describe, expect, it } from 'vitest'

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
    // 中 is 3 bytes on the wire; a 2-byte cap must already trip
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
    // offsets 1 and 2 cut inside 中; 0 (right before the lead byte) is the
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
})
