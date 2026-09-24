/**
 * Minimal structural reader for the zip produced by savePptxToFile, used to pin the
 * LibreOffice-compat contract (BUG-1671): local file headers must carry the real
 * crc/size values and must NOT use general-purpose bit 3 / trailing data descriptors,
 * and the central directory must agree with every local header.
 *
 * Deliberately not built on JSZip: JSZip tolerates data descriptors, so it cannot
 * observe the very defect this guards against. We walk the raw bytes instead.
 */

export const LOCAL_HEADER_SIG = 0x04034b50
export const CENTRAL_HEADER_SIG = 0x02014b50
export const DATA_DESCRIPTOR_SIG = 0x08074b50

export interface CentralDirectoryEntry {
  name: string
  /** general purpose bit flag from the central directory */
  flags: number
  crc32: number
  compressedSize: number
  uncompressedSize: number
  /** absolute offset of the matching local file header */
  localHeaderOffset: number
}

export interface ZipStructureAudit {
  entries: CentralDirectoryEntry[]
  /** entries whose central-directory record sets bit 3 (data descriptors) */
  descriptorFlagged: string[]
  /** entries whose local header disagrees with the central directory or uses bit 3 */
  headerProblems: string[]
  /** entries whose data is followed by a data-descriptor signature instead of the next record */
  trailingDescriptors: string[]
}

export function auditZipStructure(buf: Buffer): ZipStructureAudit {
  const eocd = buf.lastIndexOf('PK\x05\x06')
  if (eocd < 0) throw new Error('not a zip: no end-of-central-directory record')
  const count = buf.readUInt16LE(eocd + 10)
  let p = buf.readUInt32LE(eocd + 16)

  const entries: CentralDirectoryEntry[] = []
  const descriptorFlagged: string[] = []
  const headerProblems: string[] = []
  const trailingDescriptors: string[] = []

  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(p) !== CENTRAL_HEADER_SIG) throw new Error(`bad central header at ${p}`)
    const flags = buf.readUInt16LE(p + 8)
    const crc32 = buf.readUInt32LE(p + 16)
    const compressedSize = buf.readUInt32LE(p + 20)
    const uncompressedSize = buf.readUInt32LE(p + 24)
    const nameLen = buf.readUInt16LE(p + 28)
    const extraLen = buf.readUInt16LE(p + 30)
    const commentLen = buf.readUInt16LE(p + 32)
    const localHeaderOffset = buf.readUInt32LE(p + 42)
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen)
    entries.push({ name, flags, crc32, compressedSize, uncompressedSize, localHeaderOffset })
    if (flags & 0x8) descriptorFlagged.push(name)
    p += 46 + nameLen + extraLen + commentLen
  }

  for (const entry of entries) {
    const o = entry.localHeaderOffset
    if (buf.readUInt32LE(o) !== LOCAL_HEADER_SIG) {
      headerProblems.push(`${entry.name}: bad local header signature`)
      continue
    }
    const localFlags = buf.readUInt16LE(o + 6)
    const localCrc = buf.readUInt32LE(o + 14)
    const localCsize = buf.readUInt32LE(o + 18)
    const localUsize = buf.readUInt32LE(o + 22)
    // Bit 3 means "sizes/crc live in a trailing data descriptor, local header is zeroed" —
    // exactly the structure LibreOffice Impress refuses.
    if (localFlags & 0x8) headerProblems.push(`${entry.name}: local header sets bit 3 (0x8)`)
    if (
      localCrc !== entry.crc32 ||
      localCsize !== entry.compressedSize ||
      localUsize !== entry.uncompressedSize
    ) {
      headerProblems.push(
        `${entry.name}: local header (${localCrc}/${localCsize}/${localUsize}) != ` +
          `central directory (${entry.crc32}/${entry.compressedSize}/${entry.uncompressedSize})`,
      )
    }
    const localNameLen = buf.readUInt16LE(o + 26)
    const localExtraLen = buf.readUInt16LE(o + 28)
    const dataEnd = o + 30 + localNameLen + localExtraLen + entry.compressedSize
    const nextSig = buf.readUInt32LE(dataEnd)
    if (nextSig === DATA_DESCRIPTOR_SIG) trailingDescriptors.push(entry.name)
  }

  return { entries, descriptorFlagged, headerProblems, trailingDescriptors }
}
