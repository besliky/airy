/**
 * BUG-1715 wiring: the protection gate's worksheet reads must see the
 * `custom.unlocked` install flag through the REAL Univer Worksheet model.
 * The live refusal came from the composed getCell path; the gate now reads
 * the raw cell matrix (getCellRaw), which always carries the flag the chunk
 * installer bakes in for xfs with <protection locked="0">.
 */
import { LocaleType, IUniverInstanceService, UniverInstanceType } from '@univerjs/core'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/lib/facade'
import { describe, expect, it } from 'vitest'

import { createUniver } from '../src/renderer/create-univer'
import {
  cellIsUnlocked,
  protectionRefusal,
  type ProtectionWorksheet,
} from '../src/renderer/sheet-protection'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

function bootWorksheet(): { sheet: ProtectionWorksheet; state: LazyWorkbookState } {
  const rt = createUniver({
    locale: LocaleType.EN_US,
    locales: {},
    presets: [{ plugins: [UniverSheetsPlugin] }],
  })
  rt.univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id: 'wb',
    name: 'wb',
    styles: {},
    sheets: {
      main: {
        id: 'main',
        name: 'Main',
        rowCount: 20,
        columnCount: 10,
        cellData: {
          // B2: the file xf says unlocked; A1 stays at the locked default.
          1: { 1: { v: 'input', custom: { unlocked: true } } },
          0: { 0: { v: 'fixed' } },
        },
      },
    },
  })
  rt.univer.__getInjector().get(IUniverInstanceService).focusUnit('wb')
  const workbook = rt.univerAPI.getActiveWorkbook()!
  const fWorksheet = workbook.getSheetBySheetId('main')!
  const sheet = fWorksheet.getSheet() as unknown as ProtectionWorksheet
  const state = {
    file: { styles: [], sheets: [{ id: 'main', columnWidths: [] }] },
    sheetProtections: new Map([['main', { protected: true, hasPassword: false }]]),
    editJournal: { cells: new Map(), structuralOps: new Map() },
  } as unknown as LazyWorkbookState
  return { sheet, state }
}

describe('protection gate against the real Univer worksheet model', () => {
  it('reads the unlocked install flag through the raw cell matrix (BUG-1715)', () => {
    const { sheet, state } = bootWorksheet()
    expect(sheet.getCellRaw).toBeTypeOf('function')
    expect(cellIsUnlocked(state, 'main', 1, 1, sheet)).toBe(true)
    expect(cellIsUnlocked(state, 'main', 0, 0, sheet)).toBe(false)
  })

  it('lets edits through on the unlocked cell and refuses the locked one', () => {
    const { sheet, state } = bootWorksheet()
    const protection = { protected: true, hasPassword: false }
    expect(
      protectionRefusal(
        state,
        'main',
        'sheet.command.set-range-values',
        { value: { 1: { 1: { v: 'typed' } } } },
        sheet,
        protection,
      ),
    ).toBeNull()
    expect(
      protectionRefusal(
        state,
        'main',
        'sheet.command.set-range-values',
        { value: { 0: { 0: { v: 'typed' } } } },
        sheet,
        protection,
      ),
    ).toBe('appSheetCellProtected')
  })
})
