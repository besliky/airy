/**
 * Color auto-filter criteria end to end at the model level: the wire shape
 * the import restores and the save snapshot collects drives Univer's own
 * by-color filter evaluation against real cell fills, and resets cleanly.
 * Uses the model plugin only — the filter panel UI needs a canvas.
 */
import { LocaleType, IUniverInstanceService, UniverInstanceType } from '@univerjs/core'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/lib/facade'
import { UniverSheetsFilterPlugin } from '@univerjs/sheets-filter'
import '@univerjs/sheets-filter/lib/facade'
import { describe, expect, it } from 'vitest'

import { createUniver } from '../src/renderer/create-univer'
import { toWireColorFilter } from '../src/renderer/univer-sync'

function bootFilter() {
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
          // Header row 0; data rows paint direct cell fills (the shape a
          // file's colorFilter dxf resolves to and the panel lists).
          0: { 0: { v: 'n' }, 1: { v: 'v' } },
          1: {
            0: { v: 1, s: { bg: { rgb: '#9CC2E6' } } },
            1: { v: 'a' },
          },
          2: { 0: { v: 2 }, 1: { v: 'b' } },
          3: {
            0: { v: 3, s: { bg: { rgb: '#FFC000' } } },
            1: { v: 'c' },
          },
        },
      },
    },
  })
  rt.univer.__getInjector().get(IUniverInstanceService).focusUnit('wb')
  const workbook = rt.univerAPI.getActiveWorkbook()!
  const worksheet = workbook.getSheetBySheetId('main')!
  const filter = worksheet.getRange(0, 0, 4, 2).createFilter()!
  return { filter }
}

describe('filter by cell color (model)', () => {
  it('keeps only rows whose fill matches the criterion and resets', () => {
    const { filter } = bootFilter()
    // The exact wire shape restoreFilterCriteria installs after resolving a
    // file's <colorFilter dxfId> — the hex normalized through ColorKit, the
    // same comparison form Univer's panel applies.
    filter.setColumnFilterCriteria(0, {
      colId: 0,
      colorFilters: { cellFillColors: ['rgb(156,194,230)'] },
    })
    expect(filter.getFilteredOutRows()).toEqual([2, 3])
    // Unfiltered (untinted) rows stay visible only when the criterion is
    // cleared: the panel's reset hands criteria=null.
    filter.setColumnFilterCriteria(0, null)
    expect(filter.getFilteredOutRows()).toEqual([])
  })

  it('filters by font color the same way', () => {
    const { filter } = bootFilter()
    filter.setColumnFilterCriteria(1, {
      colId: 1,
      colorFilters: { cellTextColors: ['rgb(0,0,0)'] },
    })
    // Column 1 has no explicit font colors: every cell resolves to Univer's
    // default black text criterion, so nothing is filtered out.
    expect(filter.getFilteredOutRows()).toEqual([])
    filter.setColumnFilterCriteria(1, null)
    expect(filter.getFilteredOutRows()).toEqual([])
  })
})

describe('toWireColorFilter', () => {
  it('maps a single fill criterion to kind + #RRGGBB', () => {
    expect(toWireColorFilter({ cellFillColors: ['rgb(156,194,230)'] })).toEqual({
      kind: 'fill',
      color: '#9CC2E6',
    })
    expect(toWireColorFilter({ cellTextColors: ['#00B050'] })).toEqual({
      kind: 'font',
      color: '#00B050',
    })
  })

  it('prefers fills over fonts, mirroring Univer evaluation', () => {
    expect(toWireColorFilter({ cellFillColors: ['#FFC000'], cellTextColors: ['#FF0000'] })).toEqual(
      { kind: 'fill', color: '#FFC000' },
    )
  })

  it('fails closed on multi-color or "no fill" criteria', () => {
    expect(() => toWireColorFilter({ cellFillColors: ['#FFC000', '#9CC2E6'] })).toThrow()
    expect(() => toWireColorFilter({ cellFillColors: [null] })).toThrow()
    expect(() => toWireColorFilter({})).toThrow()
  })
})
