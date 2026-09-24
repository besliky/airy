/**
 * PERF-1663 fix-round guard for the dual-backend zip writer: the Node fast
 * path (hand-rolled container + node:zlib) and the sandboxed-renderer
 * fallback (JSZip) must both emit a standard zip that a reader accepts with
 * CRC verification on, and zipFiles must keep picking the fast path where
 * node:zlib exists (the 500k-row win is Node-only by design).
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import {
  zipFiles,
  zipFilesWithJsZip,
  zipFilesWithNodeZlib,
  type ZipEntry,
} from '../src/gateway/xlsx-minizip'

const ENTRIES: readonly ZipEntry[] = [
  { name: '[Content_Types].xml', data: new TextEncoder().encode('<Types/>') },
  {
    name: 'xl/workbook.xml',
    data: new TextEncoder().encode('<workbook>строка «ёлки» 😀</workbook>'),
  },
  // the huge worksheet stand-in: repetitive bytes across the deflate window
  { name: 'xl/worksheets/sheet1.xml', data: new TextEncoder().encode('<row>'.repeat(10_000)) },
  { name: 'xl/empty-part.xml', data: new Uint8Array(0) },
]

const ENTRY_NAMES = ENTRIES.map((entry) => entry.name)

/** loads with CRC verification on, so a mis-wired checksum fails the test */
async function readEntries(bytes: Uint8Array): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true })
  const names = Object.keys(zip.files)
    .filter((name) => !zip.files[name]!.dir)
    .sort()
  expect(names).toEqual([...ENTRY_NAMES].sort())
  const contents = new Map<string, string>()
  for (const name of names) {
    contents.set(name, await zip.file(name)!.async('string'))
  }
  return contents
}

/** offset of the first occurrence of a name (the local header, not the CD) */
function localHeaderNameOffset(bytes: Uint8Array, name: string): number {
  const needle = new TextEncoder().encode(name)
  for (let index = 0; index + needle.length < bytes.length; index += 1) {
    if (needle.every((byte, position) => bytes[index + position] === byte)) return index
  }
  throw new Error(`${name} not found in container`)
}

describe('zipFilesWithNodeZlib (main-process fast path)', () => {
  it('emits a valid CRC-checked zip with every entry intact', async () => {
    const contents = await readEntries(await zipFilesWithNodeZlib(ENTRIES))
    expect(contents.get('xl/workbook.xml')).toContain('строка «ёлки» 😀')
    expect(contents.get('xl/worksheets/sheet1.xml')).toHaveLength(10_000 * 5)
    expect(contents.get('xl/empty-part.xml')).toBe('')
  })

  it('rejects a container with a flipped compressed byte (checksums wired)', async () => {
    const bytes = await zipFilesWithNodeZlib(ENTRIES)
    // 30-byte fixed local header sits between the name and the payload
    const payloadByte = localHeaderNameOffset(bytes, 'xl/worksheets/sheet1.xml') + 30 + 3
    const corrupt = new Uint8Array(bytes)
    corrupt[payloadByte] = corrupt[payloadByte]! ^ 0xff
    await expect(JSZip.loadAsync(corrupt, { checkCRC32: true })).rejects.toThrow()
  })
})

describe('zipFilesWithJsZip (sandboxed-renderer fallback)', () => {
  it('emits a valid CRC-checked zip with every entry intact', async () => {
    const contents = await readEntries(await zipFilesWithJsZip(ENTRIES))
    expect(contents.get('xl/workbook.xml')).toContain('строка «ёлки» 😀')
    expect(contents.get('xl/worksheets/sheet1.xml')).toHaveLength(10_000 * 5)
    expect(contents.get('xl/empty-part.xml')).toBe('')
  })
})

describe('zipFiles (environment selector)', () => {
  it('picks the node:zlib fast path under Node (byte-identical container)', async () => {
    const [selected, fast] = await Promise.all([zipFiles(ENTRIES), zipFilesWithNodeZlib(ENTRIES)])
    expect(Buffer.compare(Buffer.from(selected), Buffer.from(fast))).toBe(0)
  })
})
