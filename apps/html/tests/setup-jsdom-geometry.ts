// jsdom has no layout engine: `Range` carries no geometry API at all, so the
// first CodeMirror measure pass that runs inside a test that builds a real
// `EditorView` throws `getClientRects is not a function` from a
// requestAnimationFrame callback. Vitest reports it as an unhandled error and
// exits non-zero even when every test passes, which turns the whole suite red
// in CI. Give `Range` the same inert geometry that jsdom `Element`s already
// expose (zero-size client rects), so measuring becomes a harmless no-op.
// Found on UX-1704: the first html tests to construct a live EditorView.
type ZeroRect = {
  x: number
  y: number
  top: number
  left: number
  right: number
  bottom: number
  width: number
  height: number
  toJSON: () => Record<string, number>
}

function zeroRect(): ZeroRect {
  return {
    x: 0,
    y: 0,
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    width: 0,
    height: 0,
    toJSON: () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
  }
}

const rangePrototype = Range.prototype as unknown as Record<string, unknown>
if (typeof rangePrototype.getClientRects !== 'function') {
  rangePrototype.getClientRects = () => Object.assign([] as DOMRect[], { item: () => null })
}
if (typeof rangePrototype.getBoundingClientRect !== 'function') {
  rangePrototype.getBoundingClientRect = zeroRect
}
