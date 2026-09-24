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
