/**
 * Open-time size fences for the binary document sessions (SEC-1102): the
 * text sessions (markdown/html, and pdf via file-parse) always opened with a
 * byte cap, but .docx/.pptx/.xlsx read whatever is on disk. A hostile file
 * pulled into an agent's workspace could then balloon the headless server's
 * memory either directly (a multi-gigabyte raw file) or as a zip bomb (a few
 * KiB whose entries declare gigabytes of uncompressed bytes).
 *
 * Two fences, cheapest first:
 *
 * 1. stat-first raw cap — refuse before reading a byte into memory.
 * 2. declared-uncompressed budget for the zip-based formats — walk the
 *    central-directory metadata (JSZip keeps the declared sizes lazily, no
 *    inflation happens) and refuse when one part or their total exceeds the
 *    budget. The docx engine already runs this fence inside parseDocx
 *    (docx-engine zip-load); the pptx engine does not, so the slides session
 *    applies the same numbers here. Workbook bytes never transit Node (the
 *    Rust sidecar reads the file itself), so xlsx gets only the raw cap.
 */
import JSZip from 'jszip'

/** raw on-disk size an open accepts (any binary document session) */
export const MAX_OPEN_BYTES = 512 * 1024 * 1024

/** the raw-cap refusal (typed so open paths can rethrow it past their
 *  generic "Cannot read" wrap without string matching) */
export class OpenSizeError extends Error {}

/** refuse an oversized raw file before reading it into memory */
export function assertWithinOpenCap(path: string, size: number, kind: string): void {
  if (size <= MAX_OPEN_BYTES) return
  throw new OpenSizeError(
    `"${path}" is ${String(size)} bytes; ${kind} sessions cap open size at ` +
      `${String(MAX_OPEN_BYTES)} bytes (512 MiB).`,
  )
}

const MAX_ZIP_PARTS = 10_000
const MAX_PART_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
const MAX_TOTAL_UNCOMPRESSED_BYTES = 1.5 * 1024 * 1024 * 1024

/**
 * Reject a zip-shaped package whose central directory declares more
 * uncompressed bytes than the budget allows, before any entry is inflated.
 * Mirrors the docx-engine fence (assertZipWithinLimits) with the same
 * numbers, so .pptx and .docx opens refuse the same bombs.
 */
export async function assertPptxWithinZipBudget(bytes: Uint8Array): Promise<void> {
  const zip = await JSZip.loadAsync(bytes)
  const files = Object.values(zip.files).filter((f) => !f.dir)
  if (files.length > MAX_ZIP_PARTS) {
    throw new Error(
      `pptx rejected: ${String(files.length)} parts exceeds the ${MAX_ZIP_PARTS} limit`,
    )
  }
  let total = 0
  for (const file of files) {
    // JSZip keeps the declared size on the lazy compressed object without
    // inflating the entry — same seam the docx-engine fence reads
    const size =
      (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0
    if (size > MAX_PART_UNCOMPRESSED_BYTES) {
      throw new Error(
        `pptx rejected: part ${file.name} declares ${String(size)} uncompressed bytes ` +
          `(limit ${String(MAX_PART_UNCOMPRESSED_BYTES)})`,
      )
    }
    if (size > 0) total += size
  }
  if (total > MAX_TOTAL_UNCOMPRESSED_BYTES) {
    throw new Error(
      `pptx rejected: total uncompressed size ${String(total)} exceeds the ` +
        `${String(MAX_TOTAL_UNCOMPRESSED_BYTES)} limit`,
    )
  }
}
