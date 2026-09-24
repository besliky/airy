/**
 * Test media builder: a valid PNG whose payload is intentionally incompressible
 * (pseudo-random noise), so stored media parts are multi-megabyte — the shape of
 * real photo/video decks — without committing binary fixtures.
 */
import { deflateSync } from 'node:zlib'

export function noisePng(width: number, height: number, seed: number): Buffer {
  const raw = Buffer.alloc(height * (1 + width * 3))
  let v = seed >>> 0
  for (let row = 0; row < height; row++) {
    const at = row * (1 + width * 3)
    raw[at] = 0 // filter: none
    for (let x = 0; x < width * 3; x++) {
      v = (Math.imul(v, 1103515245) + 12345) >>> 0
      raw[at + 1 + x] = v & 0xff
    }
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = 8 // bit depth
  ihdr[9] = 2 // color type RGB
  const chunk = (type: string, data: Buffer): Buffer => {
    const head = Buffer.alloc(8)
    head.writeUInt32BE(data.length, 0)
    head.write(type, 4, 'ascii')
    return Buffer.concat([head, data])
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw, { level: 6 })),
    chunk('IEND', Buffer.alloc(0)),
  ])
}
