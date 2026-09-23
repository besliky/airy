import { describe, expect, it } from 'vitest'
import { buildParseMap } from '../src/renderer/document/parse-map'
import { instrumentForPreview, SID_ATTR } from '../src/renderer/preview/instrument'

/** the pre-optimization algorithm: splice from the back, one full string copy per insert */
function instrumentNaive(
  text: string,
  map: ReturnType<typeof buildParseMap>,
  inspector: string,
): string {
  const inserts = map.elements
    .map((e) => {
      const tag = text.slice(e.startTag[0], e.startTag[1])
      const close = /\s*\/?>$/.exec(tag)
      const at = e.startTag[0] + (close ? close.index : tag.length)
      return { at, text: ` ${SID_ATTR}="${e.sid}"` }
    })
    .sort((a, b) => b.at - a.at)
  let out = text
  for (const ins of inserts) out = out.slice(0, ins.at) + ins.text + out.slice(ins.at)
  const script = `<script data-gx-inspector>${inspector.replace('__GX_VERSION__', String(map.version))}</script>`
  const bodyClose = out.search(/<\/body\s*>/i)
  return bodyClose >= 0 ? out.slice(0, bodyClose) + script + out.slice(bodyClose) : out + script
}

function inspectorScript(version: number): string {
  return `<script data-gx-inspector>${version}</script>`
}

describe('instrumentForPreview', () => {
  const inspector = '__GX_VERSION__'

  it('is byte-identical to the naive back-to-front splice on a structured document', () => {
    const text =
      '<!doctype html><html><head><title>T</title></head><body>' +
      '<h1 class="a">Hi</h1><p>one</p><div><p>two</p><img src="x.png"><br></div>' +
      '<table><tbody><tr><td>c</td></tr></tbody></table></body></html>'
    const map = buildParseMap(text, 3, null)
    const out = instrumentForPreview(text, map, inspector)
    expect(out).toBe(instrumentNaive(text, map, inspector))
    expect(out).toContain(inspectorScript(3))
    // stripping the injections restores the source byte-for-byte
    expect(
      out.replace(new RegExp(` ${SID_ATTR}="\\d+"`, 'g'), '').replace(inspectorScript(3), ''),
    ).toBe(text)
    expect(map.elements.length).toBeGreaterThan(8)
  })

  it('lands the sid before the end of the start tag, keeping self-closing slashes intact', () => {
    const text = '<body><p>plain</p><img src="a.png"/><em>x</em></body>'
    const map = buildParseMap(text, 1, null)
    const out = instrumentForPreview(text, map, inspector)
    expect(out).toBe(instrumentNaive(text, map, inspector))
    expect(out).toContain('<p data-sid=')
    expect(out).toContain('<img src="a.png" data-sid=')
    expect(out).toContain('/><em data-sid=')
  })

  it('injects the inspector before </body>, or appends when there is none', () => {
    const withBody = '<html><body><p>a</p></body></html>'
    const map1 = buildParseMap(withBody, 7, null)
    const out1 = instrumentForPreview(withBody, map1, inspector)
    expect(out1.endsWith('</body></html>')).toBe(true)
    expect(out1).toContain(`${inspectorScript(7)}</body>`)

    const noBody = '<html><head></head><p>a</p>'
    const map2 = buildParseMap(noBody, 7, null)
    expect(instrumentForPreview(noBody, map2, inspector).endsWith(inspectorScript(7))).toBe(true)
  })

  it('stays linear: a wide document instruments within the budget', () => {
    // regression canary for the removed O(inserts x length) splice: 8k paragraphs / ~0.9MB
    const p = '<p class="row">Lorem ipsum dolor sit amet, consectetur adipiscing elit.</p>'
    const text = `<html><body>${p.repeat(8000)}</body></html>`
    const map = buildParseMap(text, 1, null)
    const t0 = performance.now()
    const out = instrumentForPreview(text, map, inspector)
    const ms = performance.now() - t0
    expect(out.length).toBeGreaterThan(text.length)
    expect(ms).toBeLessThan(5000)
  })
})
