/**
 * BUG-1752 regression at the model level: a table-slicer toggle must leave
 * its criteria in the filter model (the save snapshot's source of truth) AND
 * hide/unhide the sheet's rows — Univer's filter packages only compute the
 * filtered row set, the visibility write is ours. Uses the model plugin only
 * (the filter panel needs a canvas), like color-filter.univer.test.ts.
 */
import { IUniverInstanceService, LocaleType, UniverInstanceType } from '@univerjs/core'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/lib/facade'
import { UniverSheetsFilterPlugin } from '@univerjs/sheets-filter'
import '@univerjs/sheets-filter/lib/facade'
import { describe, expect, it } from 'vitest'

import { applySlicerCriteria } from '../src/renderer/univer-sync'
import { createUniver } from '../src/renderer/create-univer'

/// The sheet's row-manager hidden flags — the state set-row-hidden/visible
/// mutations drive and the grid renders.
function rowHidden(worksheet: { getSheet(): unknown }, row: number): number {
  const manager = (
    worksheet.getSheet() as {
      getRowManager(): { getRow(row: number): { hd?: number } | null | undefined }
    }
  ).getRowManager()
  return manager.getRow(row)?.hd ?? 0
}

function boot() {
  const rt = createUniver({
    locale: LocaleType.EN_US,
    locales: {},
    presets: [{ plugins: [UniverSheetsPlugin, UniverSheetsFilterPlugin] }],
  })
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
          3: { 0: { v: 'alpha' }, 1: { v: 3 } },
        },
      },
    },
  })
  rt.univer.__getInjector().get(IUniverInstanceService).focusUnit('wb')
  const workbook = rt.univerAPI.getActiveWorkbook()!
  const worksheet = workbook.getSheetBySheetId('main')!
  // The filter over the table area, the way applySheetFilter installs a
  // table-owned file filter at load.
  const filter = worksheet.getRange(0, 0, 4, 2).createFilter()!
  return { rt, worksheet, filter }
}

describe('table slicer criteria reach the filter model and rows (BUG-1752)', () => {
  it('hides the rows outside the selection and reports the applied criteria', () => {
    const { rt, worksheet, filter } = boot()
    const applied = applySlicerCriteria(rt, worksheet, 0, { values: ['beta'] })
    // The model holds the criteria the save snapshot will collect.
    expect(applied?.filters?.filters).toEqual(['beta'])
    expect(applied?.filters?.blank).toBeUndefined()
    expect(filter.getColumnFilterCriteria(0)?.filters?.filters).toEqual(['beta'])
    // The model's filtered set (hiddenRows in the save snapshot) and the
    // sheet's actual row visibility agree.
    expect(filter.getFilteredOutRows()).toEqual([1, 3])
    expect(rowHidden(worksheet, 1)).toBe(1)
    expect(rowHidden(worksheet, 3)).toBe(1)
    expect(rowHidden(worksheet, 2)).toBe(0)
  })

  it('unhides the filter rows again when the selection clears', () => {
    const { rt, worksheet, filter } = boot()
    expect(applySlicerCriteria(rt, worksheet, 0, { values: ['alpha'] })).not.toBeNull()
    expect(filter.getFilteredOutRows()).toEqual([2])
    const cleared = applySlicerCriteria(rt, worksheet, 0, null)
    expect(cleared).toBeNull()
    // The facade reports a cleared column as null/undefined.
    expect(filter.getColumnFilterCriteria(0) ?? null).toBeNull()
    expect(filter.getFilteredOutRows()).toEqual([])
    expect(rowHidden(worksheet, 1)).toBe(0)
    expect(rowHidden(worksheet, 2)).toBe(0)
    expect(rowHidden(worksheet, 3)).toBe(0)
  })

  it('carries the blank slot and throws on a column outside the filter', () => {
    const { rt, worksheet } = boot()
    const applied = applySlicerCriteria(rt, worksheet, 1, { values: ['1'], blank: true })
    expect(applied?.filters?.filters).toEqual(['1'])
    expect(applied?.filters?.blank).toBe(true)
    expect(() => applySlicerCriteria(rt, worksheet, 4, null)).toThrow(
      'outside the auto-filter range',
    )
  })
})
