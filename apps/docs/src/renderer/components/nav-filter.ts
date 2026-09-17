/**
 * Pure search helpers for the navigation pane, extracted so the filter logic
 * is testable without an editor. Case-insensitive substring matching on the
 * heading text; an empty/blank query means "show everything".
 */
export interface NavHeadingLike {
  text: string
}

export function navFilterNeedle(query: string): string {
  return query.trim().toLowerCase()
}

export function filterNavHeadings<T extends NavHeadingLike>(headings: T[], query: string): T[] {
  const needle = navFilterNeedle(query)
  if (!needle) return headings
  return headings.filter((h) => h.text.toLowerCase().includes(needle))
}

/** One piece of a heading label: matched substrings are flagged for emphasis. */
export interface NavLabelPart {
  text: string
  hit: boolean
}

/**
 * Split a heading label around every case-insensitive occurrence of the
 * query. An empty needle returns the whole text as a single non-hit part.
 * Adjacent/overlapping matches are not merged beyond their natural spans
 * (indexOf advances past each hit).
 */
export function splitNavLabel(text: string, query: string): NavLabelPart[] {
  const needle = navFilterNeedle(query)
  if (!needle) return [{ text, hit: false }]
  const hay = text.toLowerCase()
  const parts: NavLabelPart[] = []
  let at = hay.indexOf(needle)
  let from = 0
  while (at !== -1) {
    if (at > from) parts.push({ text: text.slice(from, at), hit: false })
    parts.push({ text: text.slice(at, at + needle.length), hit: true })
    from = at + needle.length
    at = hay.indexOf(needle, from)
  }
  if (from < text.length) parts.push({ text: text.slice(from), hit: false })
  return parts
}
