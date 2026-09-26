/// The <meta charset> declaration must tell the truth about the bytes on disk
/// (BUG-1782). The GUI editor's save path used to leave the declaration
/// untouched while re-encoding the document: a cp1251 file whose pick pinned
/// windows-1251 was written back as cp1251 with its old claim intact (fine),
/// but a fallback UTF-8 write — or a plain edited save of a legacy-declared
/// file — produced UTF-8 bytes under a stale legacy claim, and browsers trust
/// the claim: they decoded the new bytes as the old charset and rendered
/// mojibake (the same damage the MCP server fixed for its own saves in
/// BUG-761).
///
/// Update, never remove — that is what the office suites do. Word's "Save As
/// Web Page" writes a declaration matching the encoding chosen in Web Options
/// (e.g. windows-1251), and LibreOffice's HTML export writes one matching its
/// HTML-compatibility encoding; neither strips the declaration when the bytes
/// are legacy-encoded. Removing it would hand browsers to encoding sniffing —
/// modern engines default unlabeled content to UTF-8, so genuinely
/// legacy-encoded bytes would mojibake everywhere — while a truthful
/// declaration keeps the file rendering correctly in every consumer. Only an
/// existing declaration is rewritten (no injection), matching the MCP
/// server's pinCharsetDeclaration precedent: a document that never declared a
/// charset gets byte-for-byte what the editor serialized.

/** the charset token of the first <meta charset>-style declaration; cannot span lines (charset names carry no whitespace) */
const META_CHARSET_RE = /<meta[^>]+charset\s*=\s*["']?\s*([a-z0-9_.:-]+)/i

/** TextDecoder's canonical name per probed label — labels are aliases ('cp1251' → 'windows-1251'), probed once each; null for labels no decoder accepts */
const canonicalNames = new Map<string, string | null>()

/**
 * The canonical name a WHATWG label resolves to, or null when no decoder
 * accepts it (a declaration no browser understands either — it is rewritten
 * rather than trusted).
 */
function canonicalCharset(label: string): string | null {
  const key = label.trim().toLowerCase()
  const known = canonicalNames.get(key)
  if (known !== undefined) return known
  let canonical: string | null
  try {
    canonical = new TextDecoder(key).encoding
  } catch {
    // a label no decoder accepts — cached so the throw happens once per label
    canonical = null
  }
  canonicalNames.set(key, canonical)
  return canonical
}

/** The charset the document declares, or null when there is no declaration. */
export function declaredHtmlCharset(text: string): string | null {
  return META_CHARSET_RE.exec(text)?.[1] ?? null
}

/**
 * Point the document's charset declaration at `encoding` — the charset the
 * save is about to write. A declaration that already names the same encoding
 * (compared canonically, so an alias like `cp1251` is not churned by a
 * windows-1251 save) is left byte-identical, which keeps untouched
 * save/save cycles idempotent. Returns the text to write and the replaced
 * label, or null when nothing needed rewriting.
 */
export function syncCharsetDeclaration(
  text: string,
  encoding: string,
): { text: string; replaced: string | null } {
  const match = META_CHARSET_RE.exec(text)
  const declared = match?.[1]
  if (!match || !declared) return { text, replaced: null }
  if (canonicalCharset(declared) === canonicalCharset(encoding)) {
    return { text, replaced: null }
  }
  // the token cannot span lines, so a plain slice splice preserves every
  // byte around it — quoting style, the rest of the tag, EOLs
  const start = match.index + match[0].length - declared.length
  return {
    text: `${text.slice(0, start)}${encoding}${text.slice(start + declared.length)}`,
    replaced: declared,
  }
}
