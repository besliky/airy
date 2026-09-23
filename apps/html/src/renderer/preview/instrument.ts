import type { ParseMap } from '../document/parse-map'

export const SID_ATTR = 'data-sid'

/**
 * The preview copy of the document: every element with a source location gets
 * a `data-sid` so the inspector can map clicks straight back to the parse map,
 * and the inspector script is appended. The saved text never sees either.
 */
export function instrumentForPreview(text: string, map: ParseMap, inspectorSource: string): string {
  // one ascending pass with a parts join: rebuilding the string per insert is
  // O(inserts × document length) and turns a 3MB document into tens of seconds
  const inserts = map.elements
    .map((e) => {
      const tag = text.slice(e.startTag[0], e.startTag[1])
      const close = /\s*\/?>$/.exec(tag)
      const at = e.startTag[0] + (close ? close.index : tag.length)
      return { at, text: ` ${SID_ATTR}="${e.sid}"` }
    })
    .sort((a, b) => a.at - b.at)
  const parts: string[] = []
  let pos = 0
  for (const ins of inserts) {
    parts.push(text.slice(pos, ins.at), ins.text)
    pos = ins.at
  }
  parts.push(text.slice(pos))
  const out = parts.join('')
  // the frame reports this version with every message so the app can drop input from a stale reload
  const script = `<script data-gx-inspector>${inspectorSource.replace('__GX_VERSION__', String(map.version))}</script>`
  const bodyClose = out.search(/<\/body\s*>/i)
  return bodyClose >= 0 ? out.slice(0, bodyClose) + script + out.slice(bodyClose) : out + script
}
