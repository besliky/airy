import { defineConfig } from 'vitest/config'

// Every test file launches its own Chromium in beforeAll (see
// tests/support/convert.ts), so vitest's file-level worker pool equals that
// many concurrent browsers. Locally that parallelism is the point: cases are
// bound by ~1.4s of fixed converter waits, not CPU, and a beefy dev box runs
// 6 Chrome instances (~1-1.5 GB) without noticing. CI runners are small
// (ubuntu-latest: 4 vCPU / 16 GB) and the root `npm test` already runs TWO
// workspace suites side by side (tools/run-workspaces.mjs), so six Chrome on
// top of a neighbor suite pushes peak memory into OOM/swap territory
// (PERF-701). Clamp the worker count only when CI=true — local runs keep
// full speed. 2 workers cap this suite at ~2 Chrome while the wall-clock
// cost stays modest (cases are waits, and testTimeout has 30-60x headroom).
const CI_CHROME_BUDGET = 2

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    // every case renders in a real Chromium
    testTimeout: 90000,
    hookTimeout: 60000,
    ...(process.env.CI ? { maxWorkers: CI_CHROME_BUDGET } : {}),
  },
})
