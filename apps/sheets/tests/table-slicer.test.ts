import { describe, expect, it } from 'vitest'

import {
  TABLE_SLICER_MAX_MEMBERS,
  slicerCriteriaApplied,
  tableSlicerMembers,
  tableSlicerSelection,
} from '../src/renderer/pivot-actions'
import {
  createEditJournal,
  isTableConverted,
  recordSlicerAdd,
  recordSheetRemove,
  recordTableAdd,
  recordTableEdit,
  removeSlicerAddsForTable,
  tableNameAfterEdits,
  toSaveSlicerAdds,
} from '../src/renderer/edit-journal'

describe('tableSlicerMembers', () => {
  it('collects distinct values in first-appearance order and collapses blanks', () => {
    const { members } = tableSlicerMembers(
      ['de', 'fr', 'de', null, '', 10, true, 'fr', false],
      '(blank)',
    )
    expect(members.map((member) => member.label)).toEqual([
      'de',
      'fr',
      '(blank)',
      '10',
      'TRUE',
      'FALSE',
    ])
    expect(members.map((member) => member.member)).toEqual([0, 1, 2, 3, 4, 5])
  })

  it('caps the member list and counts the distinct values beyond it (UX-1762)', () => {
    const values = Array.from({ length: TABLE_SLICER_MAX_MEMBERS + 50 }, (_v, i) => `v${i}`)
    const result = tableSlicerMembers(values, '(blank)')
    expect(result.members).toHaveLength(TABLE_SLICER_MAX_MEMBERS)
    expect(result.moreCount).toBe(50)
  })

  it('reports zero overflow when the column fits under the cap', () => {
    expect(tableSlicerMembers(['de', 'fr', 'de'], '(blank)')).toEqual({
      members: [
        { member: 0, label: 'de' },
        { member: 1, label: 'fr' },
      ],
      moreCount: 0,
    })
  })

  it('keeps deduplicating and blank-collapsing past the cap', () => {
    const values: (string | null)[] = [
      // Under the cap: one blank plus 199 distinct strings.
      null,
      ...Array.from({ length: TABLE_SLICER_MAX_MEMBERS - 1 }, (_v, i) => `v${i}`),
      // Over the cap: a repeat of a shown value, a repeat of the shown
      // blank, and two genuinely new distinct values.
      'v0',
      null,
      'new-1',
      'new-2',
      'new-1',
    ]
    const result = tableSlicerMembers(values, '(blank)')
    expect(result.members).toHaveLength(TABLE_SLICER_MAX_MEMBERS)
    expect(result.moreCount).toBe(2)
  })
})

describe('tableSlicerSelection', () => {
  const { members } = tableSlicerMembers(['de', 'fr', 'it'], '(blank)')

  it('selects everything without criteria', () => {
    expect(tableSlicerSelection(members, null)).toEqual([0, 1, 2])
    expect(tableSlicerSelection(members, undefined)).toEqual([0, 1, 2])
    expect(tableSlicerSelection(members, {})).toEqual([0, 1, 2])
  })

  it('selects exactly the values the filter keeps', () => {
    // OOXML <filters> lists the KEPT values.
    expect(tableSlicerSelection(members, { values: ['de', 'it'] })).toEqual([0, 2])
    expect(tableSlicerSelection(members, { values: [], blank: true })).toEqual([])
  })

  it('maps a kept blank criterion onto the blank member', () => {
    const { members: withBlank } = tableSlicerMembers(['de', '', 'it'], '(blank)')
    // The blank member's label is the locale's blank text, matched by the
    // blank flag, not by the label.
    expect(tableSlicerSelection(withBlank, { values: ['de'], blank: true })).toEqual([0, 1])
    expect(tableSlicerSelection(withBlank, { values: ['de'] })).toEqual([0])
  })

  it('starts fully selected on custom-criteria filters (unrepresentable)', () => {
    // A criteria object without values/blank carries no member mapping.
    expect(tableSlicerSelection(members, {})).toEqual([0, 1, 2])
  })
})

describe('slicer adds across table rename/convert (BUG-1751)', () => {
  it('re-addresses a pending rename: the save carries the post-rename table name', () => {
    const journal = createEditJournal()
    recordSlicerAdd(journal, { sheetId: 's1', tableName: 'Sales', colId: 0 })
    recordTableEdit(journal, { sheetId: 's1', tableName: 'Sales', rename: 'Sales2026' })
    // The gateway renames the table part first and then finds the slicer's
    // table by name — the entry must arrive under the new name.
    expect(toSaveSlicerAdds(journal)).toEqual([{ sheetId: 's1', tableName: 'Sales2026', colId: 0 }])
    // A slicer created after the rename records the file's name and resolves
    // to the same target (successive renames merge into one entry).
    recordSlicerAdd(journal, { sheetId: 's1', tableName: 'Sales', colId: 2 })
    expect(toSaveSlicerAdds(journal)).toEqual([
      { sheetId: 's1', tableName: 'Sales2026', colId: 0 },
      { sheetId: 's1', tableName: 'Sales2026', colId: 2 },
    ])
  })

  it('drops slicers of a converted table and keeps the other tables slicers', () => {
    const journal = createEditJournal()
    recordSlicerAdd(journal, { sheetId: 's1', tableName: 'Sales', colId: 0 })
    recordSlicerAdd(journal, { sheetId: 's1', tableName: 'Other', colId: 1 })
    recordTableEdit(journal, { sheetId: 's1', tableName: 'Sales', convertToRange: true })
    expect(isTableConverted(journal, 's1', 'Sales')).toBe(true)
    expect(toSaveSlicerAdds(journal)).toEqual([{ sheetId: 's1', tableName: 'Other', colId: 1 }])
  })

  it('removeSlicerAddsForTable drops the matching entries at apply time', () => {
    const journal = createEditJournal()
    recordTableAdd(journal, {
      sheetId: 's1',
      name: 'Session1',
      area: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 },
      columnNames: ['A', 'B'],
      bandedRows: true,
    })
    recordSlicerAdd(journal, { sheetId: 's1', tableName: 'Session1', colId: 0 })
    recordSlicerAdd(journal, { sheetId: 's1', tableName: 'Session1', colId: 1 })
    recordSlicerAdd(journal, { sheetId: 's2', tableName: 'Session1', colId: 0 })
    expect(removeSlicerAddsForTable(journal, 's1', 'Session1')).toBe(2)
    expect(journal.slicerAdds).toEqual([{ sheetId: 's2', tableName: 'Session1', colId: 0 }])
    expect(removeSlicerAddsForTable(journal, 's1', 'Session1')).toBe(0)
  })

  it('renames do not strand removed-sheet filtering and convert survives sheet removal', () => {
    const journal = createEditJournal()
    recordSlicerAdd(journal, { sheetId: 's1', tableName: 'Sales', colId: 0 })
    recordSheetRemove(journal, 's1')
    recordTableEdit(journal, { sheetId: 's1', tableName: 'Sales', rename: 'Sales2026' })
    expect(toSaveSlicerAdds(journal)).toEqual([])
    expect(tableNameAfterEdits(journal, 's1', 'Sales')).toBe('Sales2026')
  })
})

describe('slicerCriteriaApplied (honest applied status, BUG-1752)', () => {
  it('accepts only an exact model snapshot of the request', () => {
    expect(slicerCriteriaApplied(null, null)).toBe(true)
    expect(slicerCriteriaApplied(null, undefined)).toBe(true)
    expect(slicerCriteriaApplied(null, { colId: 0, filters: { filters: ['de'] } })).toBe(false)
    const applied = { colId: 0, filters: { filters: ['de', 'it'] } }
    expect(slicerCriteriaApplied({ values: ['it', 'de'] }, applied)).toBe(true)
    expect(slicerCriteriaApplied({ values: ['de'] }, applied)).toBe(false)
    expect(slicerCriteriaApplied({ values: ['de', 'it', 'fr'] }, applied)).toBe(false)
    // The blank slot is part of the contract in both directions.
    expect(
      slicerCriteriaApplied(
        { values: ['de'], blank: true },
        {
          colId: 0,
          filters: { filters: ['de'], blank: true },
        },
      ),
    ).toBe(true)
    expect(slicerCriteriaApplied({ values: ['de'], blank: true }, applied)).toBe(false)
    expect(
      slicerCriteriaApplied(
        { values: ['de'] },
        {
          colId: 0,
          filters: { filters: ['de'], blank: true },
        },
      ),
    ).toBe(false)
    // Custom-criteria columns (no filters block) never count as applied.
    expect(
      slicerCriteriaApplied(
        { values: ['de'] },
        {
          colId: 0,
          customFilters: { customFilters: [{ val: 1 }] },
        },
      ),
    ).toBe(false)
  })
})
