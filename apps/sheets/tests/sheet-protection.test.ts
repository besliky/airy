import { describe, expect, it } from 'vitest'

import {
  recordSheetProtection,
  createEditJournal,
  type EditJournal,
} from '../src/renderer/edit-journal'
import {
  actionPrevented,
  cellIsUnlocked,
  effectiveSheetProtection,
  inspectProtectionPatch,
  protectionRefusal,
  rangeHasLockedCell,
  unprotectPasswordStatus,
} from '../src/renderer/sheet-protection'
import { excelLegacyPasswordHash } from '../src/shared/legacy-password'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

type FileProtection = { protected: boolean; hasPassword: boolean; passwordHash?: string }

function makeState(
  file: Record<string, FileProtection | undefined>,
  overrides: Partial<Record<string, unknown>> = {},
): LazyWorkbookState {
  return {
    file: {
      styles: [{}, { locked: false }],
      sheets: [
        {
          id: 's1',
          columnWidths: [{ startColumn: 1, endColumn: 3, hidden: false, styleIndex: 1 }],
        },
      ],
    },
    sheetProtections: new Map(Object.entries(file).filter(([, v]) => v !== undefined)) as Map<
      string,
      FileProtection
    >,
    editJournal: createEditJournal(),
    ...overrides,
  } as unknown as LazyWorkbookState
}

describe('effectiveSheetProtection', () => {
  it('returns the file state untouched without a delta', () => {
    const state = makeState({ s1: { protected: true, hasPassword: true, passwordHash: '83AF' } })
    expect(effectiveSheetProtection(state, 's1')).toMatchObject({
      protected: true,
      hasPassword: true,
      passwordHash: '83AF',
    })
  })

  it('the journal delta wins over the file state', () => {
    const state = makeState({ s1: { protected: true, hasPassword: true, passwordHash: '83AF' } })
    recordSheetProtection(
      state.editJournal,
      's1',
      { protected: false, passwordHash: null },
      { protected: true, passwordHash: '83AF' },
    )
    const protection = effectiveSheetProtection(state, 's1')
    expect(protection?.protected).toBe(false)
    expect(protection?.hasPassword).toBe(false)
  })

  it('null for unknown sheets and added sheets without a record', () => {
    const state = makeState({})
    expect(effectiveSheetProtection(state, 'missing')).toBeNull()
  })
})

describe('recordSheetProtection', () => {
  it('deletes the delta when the desired state matches the original', () => {
    const journal: EditJournal = createEditJournal()
    recordSheetProtection(journal, 's1', { protected: true }, { protected: true })
    expect(journal.sheetProtection.has('s1')).toBe(false)
  })

  it('keeps the password in the delta when protecting with one', () => {
    const journal: EditJournal = createEditJournal()
    const hash = excelLegacyPasswordHash('secret')
    recordSheetProtection(
      journal,
      's1',
      { protected: true, passwordHash: hash },
      { protected: false },
    )
    expect(journal.sheetProtection.get('s1')).toEqual({
      protected: true,
      passwordHash: hash,
    })
  })

  it('round-trips protect → unprotect back to a clean journal', () => {
    const journal: EditJournal = createEditJournal()
    const original = { protected: true, passwordHash: '83AF' }
    recordSheetProtection(journal, 's1', { protected: false, passwordHash: null }, original)
    expect(journal.sheetProtection.get('s1')?.protected).toBe(false)
    recordSheetProtection(journal, 's1', original, original)
    expect(journal.sheetProtection.has('s1')).toBe(false)
  })
})

describe('unprotectPasswordStatus', () => {
  it('accepts the correct legacy password and refuses a wrong one', () => {
    const file = { hasPassword: true, passwordHash: excelLegacyPasswordHash('secret') }
    expect(unprotectPasswordStatus(file, 'secret')).toBe('ok')
    expect(unprotectPasswordStatus(file, 'Secret')).toBe('wrong')
    expect(unprotectPasswordStatus(file, '')).toBe('wrong')
  })

  it('fails closed on the modern hash form and passes passwordless sheets', () => {
    expect(unprotectPasswordStatus({ hasPassword: true }, 'anything')).toBe('unsupported')
    expect(unprotectPasswordStatus({ hasPassword: false }, '')).toBe('ok')
  })
})

describe('cellIsUnlocked', () => {
  it('defaults to locked for plain cells', () => {
    const state = makeState({})
    expect(cellIsUnlocked(state, 's1', 0, 0)).toBe(false)
  })

  it('honors the journal Lock Cell patch over everything', () => {
    const journal = createEditJournal()
    journal.cells.set(
      's1',
      new Map([
        [
          '0:0',
          { row: 0, column: 0, hasValue: false, value: null, style: { protectionLocked: false } },
        ],
      ]),
    )
    const state = makeState({}, { editJournal: journal })
    expect(cellIsUnlocked(state, 's1', 0, 0)).toBe(true)
  })

  it('reads the custom install flag from the worksheet', () => {
    const state = makeState({})
    const worksheet = {
      getCell: (row: number, column: number) =>
        row === 1 && column === 1 ? { custom: { unlocked: true } } : { custom: null },
    }
    expect(cellIsUnlocked(state, 's1', 1, 1, worksheet)).toBe(true)
    expect(cellIsUnlocked(state, 's1', 0, 0, worksheet)).toBe(false)
  })

  it('resolves column-default styles from the file metadata', () => {
    // sheet1's columnWidths unlock columns 1..3 (styleIndex 1 → locked=false).
    const state = makeState({})
    expect(cellIsUnlocked(state, 's1', 5, 2)).toBe(true)
    expect(cellIsUnlocked(state, 's1', 5, 4)).toBe(false)
  })
})

describe('inspectProtectionPatch', () => {
  it('finds the locked cell in a value patch', () => {
    const state = makeState({})
    const patch = { 0: { 0: { v: 'x' } }, 1: { 1: { v: 'y' } } }
    const inspection = inspectProtectionPatch(state, 's1', patch, false, {
      getCell: (row: number, column: number) =>
        row === 1 && column === 1 ? { custom: { unlocked: true } } : null,
    })
    expect(inspection.locked).toEqual({ row: 0, column: 0 })
    expect(inspection.allStyleOnly).toBe(false)
  })

  it('skips style-only patches when formatting is allowed', () => {
    const state = makeState({})
    const patch = { 0: { 0: { s: { bg: { rgb: '#FF0000' } } } } }
    expect(inspectProtectionPatch(state, 's1', patch, true).locked).toBeNull()
    // Same patch with formatting disallowed still hits the locked cell.
    expect(inspectProtectionPatch(state, 's1', patch, false).locked).toEqual({ row: 0, column: 0 })
  })

  it('allows patches on unlocked cells', () => {
    const state = makeState({})
    const patch = { 1: { 1: { v: 'typed' } } }
    const inspection = inspectProtectionPatch(state, 's1', patch, false, {
      getCell: (row: number, column: number) =>
        row === 1 && column === 1 ? { custom: { unlocked: true } } : null,
    })
    expect(inspection.locked).toBeNull()
  })
})

describe('actionPrevented', () => {
  const protection = { protected: true, hasPassword: false }
  it('defaults the prevented-by-default actions', () => {
    for (const action of [
      'formatCells',
      'insertRows',
      'deleteColumns',
      'sort',
      'autoFilter',
    ] as const) {
      expect(actionPrevented(protection as never, action)).toBe(true)
    }
  })

  it('respects explicit raw attributes', () => {
    expect(
      actionPrevented(
        { protected: true, hasPassword: false, insertRows: false } as never,
        'insertRows',
      ),
    ).toBe(false)
  })
})

describe('rangeHasLockedCell / protectionRefusal', () => {
  const unlockedAt = (rows: readonly number[], columns: readonly number[]) => ({
    getCell: (row: number, column: number) =>
      rows.includes(row) && columns.includes(column) ? { custom: { unlocked: true } } : null,
  })

  it('scans a range for locked cells', () => {
    const state = makeState({})
    // Only (0,0) is unlocked: a range confined to it is clean...
    expect(
      rangeHasLockedCell(
        state,
        's1',
        { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        unlockedAt([0], [0]),
      ),
    ).toBe(false)
    // ...a range reaching the locked neighbor is not, and a column-default
    // locked cell (column 0 has no <col style>) counts as locked too.
    expect(
      rangeHasLockedCell(
        state,
        's1',
        { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
        unlockedAt([1], [0]),
      ),
    ).toBe(true)
  })

  it('refuses set-range-values touching locked cells (Excel dialog message)', () => {
    const state = makeState({ s1: { protected: true, hasPassword: false } })
    const worksheet = unlockedAt([1], [1])
    expect(
      protectionRefusal(
        state,
        's1',
        'sheet.command.set-range-values',
        { value: { 0: { 0: { v: 1 } } } },
        worksheet,
        { protected: true, hasPassword: false },
      ),
    ).toBe('appSheetCellProtected')
    // Unlocked target: allowed.
    expect(
      protectionRefusal(
        state,
        's1',
        'sheet.command.set-range-values',
        { value: { 1: { 1: { v: 1 } } } },
        worksheet,
        { protected: true, hasPassword: false },
      ),
    ).toBeNull()
  })

  it('refuses structural actions per attributes and allows them when the file allows', () => {
    const state = makeState({ s1: { protected: true, hasPassword: false } })
    const protection = { protected: true, hasPassword: false }
    expect(
      protectionRefusal(
        state,
        's1',
        'sheet.command.insert-row',
        { range: {} },
        undefined,
        protection,
      ),
    ).toBe('appSheetActionProtected')
    expect(
      protectionRefusal(
        state,
        's1',
        'sheet.command.remove-col',
        { range: {} },
        undefined,
        protection,
      ),
    ).toBe('appSheetActionProtected')
    expect(
      protectionRefusal(
        state,
        's1',
        'sheet.command.sort-range',
        { range: {} },
        undefined,
        protection,
      ),
    ).toBe('appSheetActionProtected')
    expect(
      protectionRefusal(
        state,
        's1',
        'sheet.command.set-filter-criteria',
        {},
        undefined,
        protection,
      ),
    ).toBe('appSheetActionProtected')
    // formatCells="0" in the file: styling commands ride the value path and
    // must not be refused for a style-only patch.
    const allowed = {
      protected: true,
      hasPassword: false,
      formatCells: false,
      insertRows: false,
      sort: false,
      autoFilter: false,
    }
    expect(
      protectionRefusal(state, 's1', 'sheet.command.insert-row', { range: {} }, undefined, allowed),
    ).toBeNull()
    expect(
      protectionRefusal(state, 's1', 'sheet.command.sort-range', { range: {} }, undefined, allowed),
    ).toBeNull()
    expect(
      protectionRefusal(state, 's1', 'sheet.command.set-filter-criteria', {}, undefined, allowed),
    ).toBeNull()
  })

  it('refuses row/column formatting per formatRows/formatColumns', () => {
    const state = makeState({ s1: { protected: true, hasPassword: false } })
    const protection = { protected: true, hasPassword: false }
    expect(
      protectionRefusal(state, 's1', 'sheet.command.set-row-height', {}, undefined, protection),
    ).toBe('appSheetActionProtected')
    expect(
      protectionRefusal(state, 's1', 'sheet.command.delta-column-width', {}, undefined, protection),
    ).toBe('appSheetActionProtected')
    expect(
      protectionRefusal(state, 's1', 'sheet.command.set-row-height', {}, undefined, {
        protected: true,
        hasPassword: false,
        formatRows: false,
      }),
    ).toBeNull()
  })

  it('ignores non-gated commands', () => {
    const state = makeState({ s1: { protected: true, hasPassword: false } })
    expect(
      protectionRefusal(state, 's1', 'sheet.command.scroll-to-cell', {}, undefined, {
        protected: true,
        hasPassword: false,
      }),
    ).toBeNull()
  })
})
