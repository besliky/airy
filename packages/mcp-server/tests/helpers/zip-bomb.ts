// Zip-bomb fixture helper (SEC-1102 / SEC-1103): rewrites the declared
// uncompressed size of every central-directory record, keeping the file a few
// hundred bytes while its metadata declares gigabytes. The central directory
// is the seam every open fence reads (JSZip's lazy `_data.uncompressedSize`,
// the Rust sidecar's `by_index_raw` walk), so the forged file trips the
// budget checks before anything inflates it.
export function patchCentralSizes(bytes: Uint8Array, size: number): Uint8Array {
  const out = bytes.slice()
  const dv = new DataView(out.buffer, out.byteOffset, out.byteLength)
  for (let i = 0; i + 28 <= out.length; i++) {
    if (out[i] === 0x50 && out[i + 1] === 0x4b && out[i + 2] === 0x01 && out[i + 3] === 0x02) {
      dv.setUint32(i + 24, size, true)
    }
  }
  return out
}
