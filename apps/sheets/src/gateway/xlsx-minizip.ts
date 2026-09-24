/// Minimal OPC-package (zip) writer for the CSV import path (PERF-1663).
/// The converted workbook is a handful of tiny XML parts plus one huge
/// worksheet part, and the profile showed JSZip's pure-JS DEFLATE dominating
/// the 500k-row import (6.5-8.7s of a ~16s pipeline). node:zlib deflates the
/// same bytes an order of magnitude faster, so the container is assembled by
/// hand: local headers + deflated parts + central directory + EOCD. The
/// output is a standard zip every OPC reader accepts (sidecar, JSZip, Excel);
/// the part count is single-digit, so no zip64 machinery is needed.

import { crc32, deflateRawSync } from 'node:zlib'

export interface ZipEntry {
  readonly name: string
  readonly data: Buffer
}

const LOCAL_HEADER_SIGNATURE = 0x04034b50
const CENTRAL_HEADER_SIGNATURE = 0x02014b50
const EOCD_SIGNATURE = 0x06054b50
/// ZIP 2.0 (deflate); the directory agrees so no zip64 markers are needed
const VERSION_NEEDED = 20
const METHOD_DEFLATE = 8
/// DOS timestamp 1980-01-01 00:00:00 (readers ignore it for workbooks)
const DOS_DATE = 0x0021

/**
 * Builds a zip container from the given entries, deflate-compressed.
 * Deflate level is pinned to 1: the worksheet part is written once and read
 * back once, so throughput matters far more than squeezing the temp file.
 */
export function zipFiles(entries: readonly ZipEntry[]): Buffer {
  const localChunks: Buffer[] = []
  const centralChunks: Buffer[] = []
  let offset = 0
  for (const entry of entries) {
    const nameBuffer = Buffer.from(entry.name, 'utf8')
    const compressed = deflateRawSync(entry.data, { level: 1 })
    const checksum = crc32(entry.data)

    const local = Buffer.alloc(30 + nameBuffer.length)
    local.writeUInt32LE(LOCAL_HEADER_SIGNATURE, 0)
    local.writeUInt16LE(VERSION_NEEDED, 4)
    local.writeUInt16LE(0, 6) // flags
    local.writeUInt16LE(METHOD_DEFLATE, 8)
    local.writeUInt16LE(0, 10) // DOS time
    local.writeUInt16LE(DOS_DATE, 12)
    local.writeUInt32LE(checksum, 14)
    local.writeUInt32LE(compressed.length, 18)
    local.writeUInt32LE(entry.data.length, 22)
    local.writeUInt16LE(nameBuffer.length, 26)
    local.writeUInt16LE(0, 28) // extra field length
    nameBuffer.copy(local, 30)
    localChunks.push(local, compressed)

    const central = Buffer.alloc(46 + nameBuffer.length)
    central.writeUInt32LE(CENTRAL_HEADER_SIGNATURE, 0)
    central.writeUInt16LE(VERSION_NEEDED, 4)
    central.writeUInt16LE(VERSION_NEEDED, 6) // version made by
    central.writeUInt16LE(0, 8) // flags
    central.writeUInt16LE(METHOD_DEFLATE, 10)
    central.writeUInt16LE(0, 12) // DOS time
    central.writeUInt16LE(DOS_DATE, 14)
    central.writeUInt32LE(checksum, 16)
    central.writeUInt32LE(compressed.length, 20)
    central.writeUInt32LE(entry.data.length, 24)
    central.writeUInt16LE(nameBuffer.length, 28)
    central.writeUInt32LE(offset, 42) // local header offset
    nameBuffer.copy(central, 46)
    centralChunks.push(central)

    offset += local.length + compressed.length
  }

  const centralSize = centralChunks.reduce((sum, chunk) => sum + chunk.length, 0)
  const eocd = Buffer.alloc(22)
  eocd.writeUInt32LE(EOCD_SIGNATURE, 0)
  eocd.writeUInt16LE(entries.length, 8)
  eocd.writeUInt16LE(entries.length, 10)
  eocd.writeUInt32LE(centralSize, 12)
  eocd.writeUInt32LE(offset, 16) // central directory offset

  return Buffer.concat([...localChunks, ...centralChunks, eocd])
}
