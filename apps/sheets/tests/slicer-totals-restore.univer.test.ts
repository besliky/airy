/**
 * BUG-1755 regression at the model level: restoring the file's own table
 * slicers after a reopen must collect members over the table's data rows
 * only — the totals band sits outside the filter (Excel semantics), the same
 * way handleCreateTableSlicer collects at creation time (PAR-203 #246).
 * Before the fix the restore read the full table area, so every reopen grew
 * the panel by a phantom "Total" member no data row can ever match. Boots the
 * real Univer DI graph (model plugin only — the filter panel needs a canvas),
 * like slicer-filter-model.univer.test.ts.
 */
import { IUniverInstanceService, LocaleType, UniverInstanceType } from '@univerjs/core'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/lib/facade'
import { UniverSheetsFilterPlugin } from '@univerjs/sheets-filter'
import '@univerjs/sheets-filter/lib/facade'
import { describe, expect, it } from 'vitest'

import { restoreImportedTableSlicers, type PivotActionContext } from '../src/renderer/pivot-actions'
import { createUniver } from '../src/renderer/create-univer'
import type { TableSlicerUiState } from '../src/renderer/SlicerPanel'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

function boot() {
  const rt = createUniver({
    locale: LocaleType.EN_US,
    locales: {},
    presets: [{ plugins: [UniverSheetsPlugin, UniverSheetsFilterPlugin] }],
  })
  // A 2-column table with a 1-row totals band: header row 0, data rows 1-4,
  // totals row 5 (what table/@totalsRowCount="1" means on disk).
  rt.univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id: 'wb',
    name: 'wb',
    styles: {},
    sheets: {
      main: {
        id: 'main',
        name: 'Main',
        rowCount: 10,
        columnCount: 5,
        cellData: {
          0: { 0: { v: 'Product' }, 1: { v: 'Qty' } },
          1: { 0: { v: 'alpha' }, 1: { v: 1 } },
          2: { 0: { v: 'beta' }, 1: { v: 2 } },
          3: { 0: { v: 'gamma' }, 1: { v: 3 } },
          4: { 0: { v: 'delta' }, 1: { v: 4 } },
          5: { 0: { v: 'Total' }, 1: { v: 10 } },
        },
      },
    },
  })
  rt.univer.__getInjector().get(IUniverInstanceService).focusUnit('wb')
  const workbook = rt.univerAPI.getActiveWorkbook()!
  const worksheet = workbook.getSheetBySheetId('main')!
  // The table-owned file filter spans the FULL table area (header + data +
  // totals), the way applySheetFilter installs it at load.
  worksheet.getRange(0, 0, 6, 2).createFilter()!
  return rt
}

function makeCtx(
  rt: ReturnType<typeof createUniver>,
  totalsRowCount: number | undefined,
): {
  ctx: PivotActionContext
  panels: TableSlicerUiState[]
} {
  // Only the fields restoreImportedTableSlicers reads: the file meta of the
  // saved workbook (tables + imported slicers) and the preload gate.
  const state = {
    file: {
      sheets: [
        {
          id: 'main',
          tables: [
            {
              name: 'Sales',
              range: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 1 },
              headerRowCount: 1,
              showRowStripes: true,
              showColumnStripes: false,
              ...(totalsRowCount === undefined ? {} : { totalsRowCount }),
              columns: ['Product', 'Qty'],
            },
          ],
          slicers: [
            {
              name: 'Slicer_Product',
              cacheName: 'Slicer_Product',
              tableName: 'Sales',
              column: 1,
            },
          ],
        },
      ],
    },
    flags: { preloadComplete: true, preloadRunning: false },
  } as unknown as LazyWorkbookState
  const panels: TableSlicerUiState[] = []
  const ctx = {
    univerRef: { current: rt },
    lazyWorkbookRef: { current: state },
    tableSlicers: [],
    setTableSlicers: (
      update: (current: readonly TableSlicerUiState[]) => readonly TableSlicerUiState[],
    ) => {
      panels.length = 0
      panels.push(...update(panels))
    },
  } as unknown as PivotActionContext
  return { ctx, panels }
}

describe('imported table slicer restore skips the totals band (BUG-1755)', () => {
  it('a reopen restores 4 members from the data rows, not 5 with "Total"', () => {
    const rt = boot()
    const { ctx, panels } = makeCtx(rt, 1)
    restoreImportedTableSlicers(ctx)
    expect(panels).toHaveLength(1)
    const labels = panels[0]!.members.map((member) => member.label)
    expect(labels).toEqual(['alpha', 'beta', 'gamma', 'delta'])
    expect(labels).not.toContain('Total')
    // No file criteria: the panel starts fully selected over the real members.
    expect(panels[0]!.selected).toEqual([0, 1, 2, 3])
  })

  it('a table without a declared totals band keeps the last row as a member', () => {
    const rt = boot()
    const { ctx, panels } = makeCtx(rt, undefined)
    restoreImportedTableSlicers(ctx)
    const labels = panels[0]!.members.map((member) => member.label)
    // No totalsRowCount in the file meta: the same grid reads as 5 data rows.
    expect(labels).toEqual(['alpha', 'beta', 'gamma', 'delta', 'Total'])
  })
})
