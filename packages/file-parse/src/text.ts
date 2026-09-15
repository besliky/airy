/// Shared byte→text decoding for user-opened plain-text files (CSV, TXT,
/// Markdown, HTML…): byte order marks win, then strict UTF-8, then the legacy
/// single- and double-byte charsets that legacy spreadsheet and word
/// processors actually write. Candidate decodes are scored so the one whose
/// letters form a real alphabet beats the plausible-looking soup a wrong
/// charset produces.

/**
 * Legacy charsets a non-UTF-8 file may be in, in tie-break order. The four
 * CJK charsets are mutually ambiguous (the same bytes decode to different but
 * equally plausible ideographs), so among them only the UI-language hint or
 * the list order can decide — hence CJK first, hint-adjusted. windows-1252
 * leads the single-byte block as the safest read of mostly-ASCII files.
 */
const LEGACY_CHARSETS = [
  'gb18030',
  'shift_jis',
  'big5',
  'euc-kr',
  'windows-1252',
  'windows-1251',
  'koi8-r',
  'windows-1250',
  'windows-1253',
  'windows-1255',
  'windows-1256',
  'windows-874',
] as const

/** how far into the file to look when scoring candidates */
const SCORE_SAMPLE_BYTES = 64 * 1024

/**
 * A script only earns its bonus once its letters cover this share of the
 * file's bytes. Real non-Latin text is alphabet-dense (Russian, Greek, Hebrew,
 * CJK easily pass); the accented tail of a Western-European file — whose
 * bytes also decode to Cyrillic letters under windows-1251 — does not.
 */
const SCRIPT_DENSITY_MIN = 0.18
/** …and at least this many bytes, so tiny headers stay hint-ordered */
const SCRIPT_BYTES_MIN = 4
/** frequent letters must reach this share of a script's letters before the common-letter bonus applies — a real language hits ~85%, a misdecode only ~50% */
const COMMON_RATIO_MIN = 0.7
/** …and at least this many of them, so a handful of lucky hits in a tiny sample cannot fake a language */
const COMMON_COUNT_MIN = 8
/** a UI-language hint only decides ties the same way: a hint whose script barely appears in the file must not override the ordered default */
const HINT_DENSITY_MIN = 0.12

/**
 * Top-frequency letters of the alphabets the single-byte charsets carry —
 * Western plus Central European Latin, Cyrillic, Greek, Hebrew, Arabic,
 * Thai. Deliberately small (roughly the fifteen most frequent letters, not
 * the whole alphabet): running text in a real language hits ~85% of them,
 * while the same bytes read through a wrong charset hit only ~50% — that
 * gap is the discriminator, and a near-complete set would erase it.
 */
const COMMON_LETTERS: Readonly<Record<Script, ReadonlySet<string>>> = {
  cjk: new Set(),
  latin: new Set('àâäçéèêëîïôöûüùÿñóáíúãõąęśćżźńłřšťčďěň'),
  cyrillic: new Set('оеаинтсрвлкмдпуязибг'),
  greek: new Set('αεινροστηκμπτά'),
  hebrew: new Set('ויהלרמאשנע'),
  arabic: new Set('ايلمنعروتبس'),
  thai: new Set('กงนมรลวสหะาีูเอ'),
}

/**
 * which UI language most plausibly produced a non-UTF-8 file — Excel writes
 * CSV in the system's legacy charset, and the UI language is the best guess
 * at the system's locale. Hindi has no legacy single-byte charset (its files
 * are UTF-8), so it maps to nothing.
 */
const LANG_CHARSETS: Record<string, string> = {
  zh: 'gb18030',
  'zh-TW': 'big5',
  ja: 'shift_jis',
  ko: 'euc-kr',
  en: 'windows-1252',
  fr: 'windows-1252',
  de: 'windows-1252',
  es: 'windows-1252',
  pt: 'windows-1252',
  it: 'windows-1252',
  nl: 'windows-1252',
  id: 'windows-1252',
  ms: 'windows-1252',
  ru: 'windows-1251',
  pl: 'windows-1250',
  cs: 'windows-1250',
  he: 'windows-1255',
  ar: 'windows-1256',
  th: 'windows-874',
}

export function legacyCharsetForLang(lang: string | undefined): string | undefined {
  return lang === undefined ? undefined : LANG_CHARSETS[lang]
}

type Script = 'cjk' | 'latin' | 'cyrillic' | 'greek' | 'hebrew' | 'arabic' | 'thai'

interface Sample {
  score: number
  /** which script the decode's letters concentrate in, for hint tie-breaking */
  script: Script | 'none'
  /** bytes the script's letters plausibly cover (see scoreSample) */
  scriptBytes: number
}

function tryDecode(bytes: Uint8Array, charset: string): string | null {
  try {
    return new TextDecoder(charset).decode(bytes)
  } catch {
    return null
  }
}

function tryDecodeStrict(bytes: Uint8Array): string | null {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes)
  } catch {
    return null
  }
}

/** how plausible a decode is: replacement chars and control bytes are bad, real alphabets and ASCII are good */
function scoreSample(text: string, byteLength: number, highBytes: number): Sample {
  let ascii = 0
  let punctuation = 0
  let replacement = 0
  let control = 0
  let halfWidthKana = 0
  let highChars = 0
  const scripts = new Map<Script, { count: number; common: number }>()

  for (const character of text) {
    const code = character.codePointAt(0) ?? 0
    if (character === '�') replacement += 1
    else if (code < 0x20 && character !== '\n' && character !== '\r' && character !== '\t')
      control += 1
    else if (code < 0x7f) ascii += 1
    // half-width katakana is the signature of double-byte text misread as
    // Shift_JIS, and near-absent from real spreadsheets — never a good sign
    else if (code >= 0xff61 && code <= 0xff9f) halfWidthKana += 1
    else {
      highChars += 1
      let script: Script | undefined
      if (
        (code >= 0x3000 && code <= 0x9fff) || // CJK punctuation, kana, unified ideographs
        (code >= 0xac00 && code <= 0xd7af) || // hangul
        (code >= 0xff00 && code <= 0xffef) // full-width forms (kana handled above)
      )
        script = 'cjk'
      else if (code >= 0x0400 && code <= 0x04ff) script = 'cyrillic'
      else if (code >= 0x0370 && code <= 0x03ff) script = 'greek'
      else if (code >= 0x0590 && code <= 0x05ff) script = 'hebrew'
      else if (code >= 0x0600 && code <= 0x06ff) script = 'arabic'
      else if (code >= 0x0e00 && code <= 0x0e7f) script = 'thai'
      // Latin-1/Extended letters — but not the ×, «», ½… symbol row, which any
      // charset produces and which says nothing about the decode
      else if (code >= 0xc0 && code <= 0x24f && code !== 0xd7 && code !== 0xf7) script = 'latin'
      if (script !== undefined) {
        const bucket = scripts.get(script) ?? { count: 0, common: 0 }
        bucket.count += 1
        if (COMMON_LETTERS[script].has(character.toLowerCase())) bucket.common += 1
        scripts.set(script, bucket)
      } else punctuation += 1
    }
  }

  let score = ascii + punctuation - 20 * replacement - 10 * control - 2 * halfWidthKana

  let dominant: { script: Script; bucket: { count: number; common: number } } | undefined
  for (const entry of scripts) {
    if (!dominant || entry[1].count > dominant.bucket.count) dominant = { script: entry[0], bucket: entry[1] }
  }
  // Score scripts in bytes, not letters: a multibyte charset turns the same
  // bytes into half as many (twice as wide) characters, so per-letter scores
  // would systematically favor single-byte misreads of CJK text.
  const scriptBytes =
    dominant === undefined || highChars === 0
      ? 0
      : (dominant.bucket.count * highBytes) / highChars
  if (scriptBytes >= SCRIPT_BYTES_MIN && scriptBytes >= SCRIPT_DENSITY_MIN * byteLength) {
    score += 2 * scriptBytes
    if (
      dominant!.bucket.common >= COMMON_RATIO_MIN * dominant!.bucket.count &&
      dominant!.bucket.common >= COMMON_COUNT_MIN
    ) {
      score += dominant!.bucket.common
    }
  }

  return { score, script: dominant === undefined ? 'none' : dominant.script, scriptBytes }
}

/**
 * Decodes text bytes: BOM, then strict UTF-8, then the legacy charsets Excel
 * and friends write. `preferred` (from the UI language) breaks exact score
 * ties between candidates that land in the same script family — GBK and
 * Shift_JIS both decode the same bytes to plausible-looking but different
 * CJK, and windows-1251/KOI8-R to different Cyrillic.
 */
export function decodeTextBytes(bytes: Uint8Array, preferred?: string): string {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return tryDecode(bytes.subarray(3), 'utf-8') ?? ''
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return tryDecode(bytes.subarray(2), 'utf-16le') ?? ''
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return tryDecode(bytes.subarray(2), 'utf-16be') ?? ''

  const utf8 = tryDecodeStrict(bytes)
  if (utf8 !== null) return utf8

  // Score candidates on a sample, then decode the full buffer once with the
  // winner — a 32 MB CSV should not pay for a dozen full legacy decodes.
  const sample = bytes.subarray(0, SCORE_SAMPLE_BYTES)
  let highBytes = 0
  for (const b of sample) if (b >= 0x80) highBytes += 1
  const score = (text: string | null): Sample =>
    scoreSample(text ?? '', sample.length, highBytes)

  let bestCharset: string | null = null
  let best = score(tryDecode(sample, 'utf-8'))
  let preferredSample: Sample | undefined
  for (const charset of LEGACY_CHARSETS) {
    const candidate = tryDecode(sample, charset)
    if (candidate === null) continue
    const scored = score(candidate)
    if (charset === preferred) preferredSample = scored
    if (scored.score > best.score) {
      best = scored
      bestCharset = charset
    }
  }
  // The hint only settles ties its script plausibly owns: which CJK variant,
  // which Cyrillic charset — never a locale guess over a script the file
  // barely contains (a French file must survive a Russian UI).
  if (
    preferred !== undefined &&
    preferredSample !== undefined &&
    bestCharset !== preferred &&
    preferredSample.score === best.score &&
    (preferredSample.script === 'none' ||
      preferredSample.scriptBytes >= HINT_DENSITY_MIN * sample.length)
  ) {
    bestCharset = preferred
  }
  return tryDecode(bytes, bestCharset ?? 'utf-8') ?? ''
}

const META_CHARSET_RE = /<meta[^>]+charset\s*=\s*["']?\s*([a-z0-9_.:-]+)/i

/**
 * Decodes an HTML document: BOM and strict UTF-8 first, then the charset the
 * document itself declares (browsers trust `<meta charset>` / the
 * content-type header equivalent), then generic detection. A `utf-8` claim
 * on bytes that are not valid UTF-8 falls through to detection instead of
 * producing replacement characters.
 */
export function decodeHtmlText(bytes: Uint8Array, preferred?: string): string {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return tryDecode(bytes.subarray(3), 'utf-8') ?? ''
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return tryDecode(bytes.subarray(2), 'utf-16le') ?? ''
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return tryDecode(bytes.subarray(2), 'utf-16be') ?? ''

  const utf8 = tryDecodeStrict(bytes)
  if (utf8 !== null) return utf8

  // windows-1252 maps every byte to a character, so it shows the ASCII parts
  // of the header verbatim for the sniff.
  const head = tryDecode(bytes.subarray(0, 4096), 'windows-1252') ?? ''
  const declared = META_CHARSET_RE.exec(head)?.[1]
  if (declared !== undefined && !/^utf-?8$/i.test(declared)) {
    const decoded = tryDecode(bytes, declared)
    if (decoded !== null && !decoded.includes('�')) return decoded
  }
  return decodeTextBytes(bytes, preferred)
}
