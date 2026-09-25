import { defineConfig } from 'vitest/config'

// PERF-1642: every suite here imports the full Univer stack (~1.4s of module
// graph per file, >340s summed), which dominated the serial wall. The bulk
// project reuses workers across files (one module graph per worker instead of
// one per file); the suites tolerate process-level state sharing (per-runtime
// unitIds, explicit dispose). Files that register `vi.mock` module mocks (or
// fs-level spies) keep per-file isolation: a mock left in the shared module
// graph poisons whichever unrelated file runs next in that worker.
const NEED_ISOLATION = [
  // Files that permanently mutate shared Univer modules (FontCache context /
  // statics, SpreadsheetSkeleton or DocumentSkeleton prototypes): the patch
  // would change measurement behavior for every unrelated file that runs
  // later in the same worker.
  'tests/autofit-line-pitch.test.ts',
  'tests/autofit-wrap-budget.test.ts',
  'tests/rtl-rich-text.test.ts',
  // Files that register `vi.mock` module mocks: a mock left in the shared
  // module graph poisons whichever unrelated file runs next in that worker.
  'tests/csv-save-back-atomic.test.ts',
  'tests/column-width-mdw.test.ts',
  'tests/column-width-narrow.test.ts',
  'tests/csv-import-lazy.test.ts',
  'tests/error-checking.test.ts',
  'tests/formula-audit.test.ts',
  'tests/lazy-find.test.ts',
  // asserts column budgets derived from the module-default MDW (7); a synced
  // workbook in an earlier file leaves a different width behind
  'tests/numfmt-fix.test.ts',
  'tests/print-cf-visuals.test.ts',
  'tests/print-dialog.test.ts',
  'tests/print-header-footer-position.test.ts',
  'tests/print-hidden.test.ts',
  'tests/print-scope-order.test.ts',
  'tests/promote-file-atomically.test.ts',
  'tests/workbook-search.test.ts',
  'tests/xlsx-atomic-sync.test.ts',
  'tests/xlsx-sidecar-cancel.test.ts',
  // fs spies only, kept isolated for the same reason
  'tests/delete-ref-rewrite.test.ts',
  'tests/delete-ref-rewrite.univer.test.ts',
  // PERF-1642 FIX ROUND 2: fail only under CI sharding (--shard=n/2), green
  // without it. When one of these files runs after certain bulk neighbors in
  // the same isolate:false worker, the global Univer DI graph is left broken
  // (`[redi]: Expect 1 dependency item(s) for id
  // "engine-render.render-manager.service" but get 0` at `new FUniver`) —
  // which neighbors matter depends on the shard's file split, so a local
  // full run cannot reproduce it. Quarantined per the documented protocol.
  'tests/dv-source-gate.univer.test.ts',
  'tests/grid-grow.univer.test.ts',
  'tests/function-catalog.registry.test.ts',
  // Boots the real Univer DI graph (same FUniver breakage class as above).
  'tests/sheet-protection.univer.test.ts',
  // Boots the real Univer DI graph (filter model + commands) for the by-color
  // criteria round-trip.
  'tests/color-filter.univer.test.ts',
  // TEST-1726: boots the real Univer DI graph for its two wrapper tests and
  // flaked in a full serial run with the exact FUniver breakage above (2
  // failed of 3138; alone it is green 5/5, so the cause is the shared
  // isolate:false worker, not the test).
  'tests/cf-segment-merge.test.ts',
]

export default defineConfig({
  test: {
    environment: 'node',
    projects: [
      {
        name: 'sheets-bulk',
        test: {
          environment: 'node',
          include: ['tests/**/*.test.ts'],
          exclude: ['**/node_modules/**', ...NEED_ISOLATION],
          isolate: false,
        },
      },
      {
        name: 'sheets-mocked',
        test: {
          environment: 'node',
          include: NEED_ISOLATION,
        },
      },
    ],
  },
})
