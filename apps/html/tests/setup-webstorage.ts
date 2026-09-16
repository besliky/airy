import { JSDOM } from 'jsdom'

// Node >= 25 enables webstorage by default, so `localStorage`/`sessionStorage`
// already exist on the Node global by the time vitest copies the jsdom window
// onto it, and vitest never overrides pre-existing globals. Without a
// `--localstorage-file` path Node's accessors resolve to inert objects, which
// surfaces as `localStorage.getItem is not a function` in tests. Swap the
// inert accessors for a real Storage backed by a private JSDOM window so the
// suites keep the empty, per-file storage contract they were written against.
const storageWindow = new JSDOM('', { url: 'http://localhost/' }).window

for (const name of ['localStorage', 'sessionStorage'] as const) {
  const present = (globalThis as Record<string, Storage | undefined>)[name]
  if (present && typeof present.getItem === 'function') continue
  Object.defineProperty(globalThis, name, {
    value: name === 'localStorage' ? storageWindow.localStorage : storageWindow.sessionStorage,
    writable: true,
    configurable: true,
    enumerable: true,
  })
}
