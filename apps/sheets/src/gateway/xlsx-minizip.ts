/// Minimal OPC-package (zip) writer for the CSV import path (PERF-1663).
/// The converted workbook is a handful of tiny XML parts plus one huge
/// worksheet part, and the profile showed JSZip's pure-JS DEFLATE dominating
/// the 500k-row import (6.5-8.7s of a ~16s pipeline). Where node:zlib exists
/// it deflates the same bytes an order of magnitude faster, so the container
/// is assembled by hand: local headers + deflated parts + central directory +
/// EOCD. The output is a standard zip every OPC reader accepts (sidecar,
/// JSZip, Excel); the part count is single-digit, so no zip64 machinery is
/// needed.
///
/// The module is also part of the renderer bundle graph (the lazy csv-import
/// chunk behind Data -> From Text/CSV), and the sandboxed renderer can never
/// load node builtins — so node:zlib is only ever a dynamic import, and
/// zipFiles falls back to the existing JSZip dependency (the pre-PERF-1663
/// writer) when it is unavailable. zipFilesWithNodeZlib / zipFilesWithJsZip
/// pin each backend so tests can cover both paths.

export interface ZipEntry {
  readonly name: string
  readonly data: Uint8Array
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
 * node:zlib under Node (main process, tooling, tests); null in the sandboxed
 * renderer. Cached because the renderer build rewrites the specifier to
 * vite's browser-external stub — a namespace with no named exports — and the
 * property probe below must stay cheap and stable after the first call.
 */
let nodeZlibPromise: Promise<typeof import('node:zlib') | null> | undefined
function loadNodeZlib(): Promise<typeof import('node:zlib') | null> {
  nodeZlibPromise ??= (async () => {
    try {
      const zlib = await import('node:zlib')
      return typeof zlib.deflateRawSync === 'function' ? zlib : null
    } catch {
      return null
    }
  })()
  return nodeZlibPromise
}

/**
 * Builds a zip container from the given entries, deflate-compressed, on the
 * fastest backend the current environment offers. Deflate level is pinned to
 * 1: the worksheet part is written once and read back once, so throughput
 * matters far more than squeezing the temp file.
 */
export async function zipFiles(entries: readonly ZipEntry[]): Promise<Uint8Array> {
  const zlib = await loadNodeZlib()
  if (zlib) return zipContainerWithNodeZlib(entries, zlib)
  return zipFilesWithJsZip(entries)
}

/**
 * The Node fast path (hand-rolled container + node:zlib). Throws outside
 * Node — zipFiles is the environment-agnostic entry point; this export
 * exists so tests can pin the shipped backend explicitly.
 */
export async function zipFilesWithNodeZlib(entries: readonly ZipEntry[]): Promise<Uint8Array> {
  const zlib = await loadNodeZlib()
  if (!zlib) throw new Error('node:zlib is unavailable in this environment')
  return zipContainerWithNodeZlib(entries, zlib)
}

/**
 * The browser fallback: the same JSZip pipeline the CSV path shipped before
 * PERF-1663 (its public API does not expose a standalone deflate, so JSZip
 * assembles the whole container). Slower by an order of magnitude, correct
 * everywhere; a lazily loaded chunk so the sandboxed renderer never pulls it
 * unless this path actually runs.
 */
export async function zipFilesWithJsZip(entries: readonly ZipEntry[]): Promise<Uint8Array> {
  const { default: JSZip } = await import('jszip')
  const zip = new JSZip()
  for (const entry of entries) zip.file(entry.name, entry.data)
  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 1 },
  })
}

/// Hand-rolled container: local headers + deflated parts + central directory
/// + EOCD, with deflate and CRC32 supplied by node:zlib. Runs under Node only.
function zipContainerWithNodeZlib(
  entries: readonly ZipEntry[],
  zlib: typeof import('node:zlib'),
): Buffer {
  const localChunks: Buffer[] = []
  const centralChunks: Buffer[] = []
  let offset = 0
  for (const entry of entries) {
    const nameBuffer = Buffer.from(entry.name, 'utf8')
    const compressed = zlib.deflateRawSync(entry.data, { level: 1 })
    const checksum = zlib.crc32(entry.data)

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
