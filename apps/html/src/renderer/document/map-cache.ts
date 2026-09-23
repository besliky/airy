/**
 * Cache for the renderer-wide parse map that keeps full-document rebuilds off
 * the typing path. A full parse5 rebuild costs O(document) — on giant single-line
 * files that is seconds — so `get()` serves the previous map (marked stale) and
 * debounces the rebuild, while correctness-critical readers (op compilation, AI
 * access, anything that must not see shifted offsets) call `now()`.
 */
export class MapCache<T> {
  private map: T | null = null
  /** document text the cached map was built from */
  private source: string | null = null
  private version = -1
  private timer: number | null = null
  private lastText = ''
  private lastVersion = -1
  private builds = 0
  private staleServes = 0

  constructor(
    private readonly build: (text: string, version: number, previous: T | null) => T,
    /** called from the debounced rebuild after the map actually changed */
    private readonly onRebuild: () => void,
    /** debounce delay for the off-path rebuild */
    private readonly delay = 300,
    /** documents above this size never auto-rebuild (the sync build would hitch the
     * main thread); their map is refreshed on demand through now() */
    private readonly autoRebuildLimit = Infinity,
  ) {}

  private valid(text: string, version: number): boolean {
    return this.map !== null && this.version === version && this.source === text
  }

  private buildNow(text: string, version: number): T {
    this.clearTimer()
    this.map = this.build(text, version, this.map)
    this.source = text
    this.version = version
    this.builds++
    return this.map
  }

  /**
   * Best-effort map for render/cursor paths. Fresh when cached; otherwise the
   * previous map plus a scheduled rebuild; a synchronous build only when there
   * is nothing to serve yet (first load).
   */
  get(text: string, version: number): { map: T; stale: boolean } {
    this.lastText = text
    this.lastVersion = version
    if (this.valid(text, version)) return { map: this.map!, stale: false }
    if (this.map === null) return { map: this.buildNow(text, version), stale: false }
    this.staleServes++
    if (text.length <= this.autoRebuildLimit) this.scheduleRebuild()
    return { map: this.map, stale: true }
  }

  /** Synchronous fresh build — for readers that cannot tolerate shifted offsets. */
  now(text: string, version: number): T {
    this.lastText = text
    this.lastVersion = version
    if (this.valid(text, version)) return this.map!
    return this.buildNow(text, version)
  }

  private clearTimer(): void {
    if (this.timer !== null) {
      window.clearTimeout(this.timer)
      this.timer = null
    }
  }

  private scheduleRebuild(): void {
    if (this.timer !== null) return
    this.timer = window.setTimeout(() => {
      this.timer = null
      if (this.valid(this.lastText, this.lastVersion)) return
      this.buildNow(this.lastText, this.lastVersion)
      this.onRebuild()
    }, this.delay)
  }

  /** test/teardown hook: drops the pending rebuild */
  dispose(): void {
    this.clearTimer()
  }

  /** counters for the debounce/incrementality regression tests */
  stats(): { builds: number; staleServes: number; pending: boolean } {
    return { builds: this.builds, staleServes: this.staleServes, pending: this.timer !== null }
  }
}
