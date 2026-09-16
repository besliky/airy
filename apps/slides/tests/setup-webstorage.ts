// Node >= 25 enables webstorage by default, so `localStorage`/`sessionStorage`
// already exist on the Node global by the time vitest copies the jsdom window
// onto it, and vitest never overrides pre-existing globals. Without a
// `--localstorage-file` path Node's accessors resolve to inert objects, which
// surfaces as `localStorage.getItem is not a function` in tests. Swap the
// inert accessors for a small in-memory Storage so the suites keep the empty,
// per-file storage contract they were written against.
function makeStorage(): Storage {
  const map = new Map<string, string>()
  return {
    get length() {
      return map.size
    },
    clear() {
      map.clear()
    },
    getItem(key: string) {
      return map.has(key) ? (map.get(key) as string) : null
    },
    key(index: number) {
      return [...map.keys()][index] ?? null
    },
    removeItem(key: string) {
      map.delete(key)
    },
    setItem(key: string, value: string) {
      map.set(key, String(value))
    },
  }
}

const globalRecord = globalThis as unknown as Record<string, Storage | undefined>
for (const name of ['localStorage', 'sessionStorage'] as const) {
  const present = globalRecord[name]
  if (present && typeof present.getItem === 'function') continue
  globalRecord[name] = makeStorage()
}
