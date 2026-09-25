/// Text → bytes for the remembered per-file encodings (BUG-1741). The save
/// path used to write `Buffer.from(text, 'utf8')` unconditionally, silently
/// transcoding a file the user had pinned to windows-1251 (or any other
/// legacy charset) into UTF-8 while the encoding-memory pick kept claiming
/// otherwise — the next open re-decoded the new bytes as the old charset and
/// produced mojibake. Saving must respect the remembered charset.
///
/// Node/Electron only ships a UTF-8 TextEncoder, so every other selectable
/// charset is encoded through its decoder: a charset's decode side defines a
/// closed character set, and its inverse map is built once per charset by
/// decoding the whole byte space (256 single bytes, plus the lead × trail
/// grid for the double-byte CJK charsets) and remembering every char that
/// decodes back. The map is small (≤ ~24k chars), built in a few tens of
/// milliseconds, cached for the process, and verified to round-trip.

/** the charsets whose text decodes from two-byte lead/trail pairs */
const DOUBLE_BYTE_CHARSETS: ReadonlySet<string> = new Set([
  'gb18030',
  'shift_jis',
  'big5',
  'euc-kr',
])

const inverseMaps = new Map<string, Map<string, number[]>>()

/** charsets TextDecoder does not know — remembered once, so the throw happens a single time per bad label */
const unknownCharsets = new Set<string>()

/**
 * The inverse of the charset's decode table: char → the bytes that decode to
 * it. Invalid byte combinations (a lone lead byte, a forbidden trail 0x7F,
 * unassigned slots) decode to U+FFFD and are skipped — those byte sequences
 * are unreachable for the encoder, which is exactly what the round-trip
 * contract needs. Scanning a slightly wider grid than any single charset's
 * real ranges keeps this free of per-charset range tables: out-of-range pairs
 * simply decode to U+FFFD and drop out.
 */
function inverseMap(charset: string): Map<string, number[]> | null {
  if (unknownCharsets.has(charset)) return null
  let map = inverseMaps.get(charset)
  if (map) return map
  let decoder: TextDecoder
  try {
    decoder = new TextDecoder(charset)
  } catch {
    unknownCharsets.add(charset)
    return null
  }
  map = new Map()
  for (let byte = 0x00; byte <= 0xff; byte++) {
    const char = decoder.decode(new Uint8Array([byte]))
    if (char !== '\uFFFD' && !map.has(char)) map.set(char, [byte])
  }
  if (DOUBLE_BYTE_CHARSETS.has(charset)) {
    for (let lead = 0x81; lead <= 0xfe; lead++) {
      for (let trail = 0x30; trail <= 0xfe; trail++) {
        if (trail === 0x7f) continue
        const char = decoder.decode(new Uint8Array([lead, trail]))
        if (char !== '\uFFFD' && !map.has(char)) map.set(char, [lead, trail])
      }
    }
  }
  inverseMaps.set(charset, map)
  return map
}

function utf16beFromLe(le: Buffer): Buffer {
  for (let i = 0; i + 1 < le.length; i += 2) {
    const first = le[i]
    le[i] = le[i + 1]
    le[i + 1] = first
  }
  return le
}

/**
 * Encode `text` as `encoding` (a SELECTABLE_ENCODINGS member). Returns null
 * when the text cannot be represented — a character outside the charset
 * (CJK ideographs in a windows-1251 file, a lone surrogate, a gb18030 char
 * only reachable through the four-byte form this map does not model) or an
 * unknown charset. The caller falls back to a lossless UTF-8 write and drops
 * the now-false remembered pick, so the reopen auto-detects the file the way
 * it actually is instead of misreading it.
 */
export function encodeTextAsEncoding(text: string, encoding: string): Buffer | null {
  if (encoding === 'utf-8') return Buffer.from(text, 'utf8')
  if (encoding === 'utf-16le') return Buffer.from(text, 'utf16le')
  if (encoding === 'utf-16be') return utf16beFromLe(Buffer.from(text, 'utf16le'))
  const map = inverseMap(encoding)
  if (!map) return null
  const out: number[] = []
  // code-point iteration: the map's keys are decoded code points, so a
  // surrogate pair in the text must arrive as one lookup, not two
  for (const char of text) {
    const bytes = map.get(char)
    if (!bytes) return null
    out.push(...bytes)
  }
  return Buffer.from(out)
}
