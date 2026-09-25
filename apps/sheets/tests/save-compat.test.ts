import { describe, expect, it } from 'vitest'

import {
  createEditJournal,
  recordSheetInsert,
  recordStructuralOp,
  type EditJournal,
  type StructuralJournalOp,
} from '../src/renderer/edit-journal'
import { zh } from '../src/renderer/i18n/app/zh'
import {
  SAVE_COMPAT_FEATURES,
  collectSaveCompatFindings,
  saveCompatDismissKey,
  saveCompatLabelKey,
  saveCompatVisible,
  type SaveCompatFinding,
} from '../src/renderer/save-compat'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

/// PAR-206: the save-compat registry is the single inventory of fail-closed
/// constructs; these tests pin each construct → feature mapping, the
/// detectors' verdicts against the gateway's fail-closed rules, and the
/// banner's appear/disappear/dismiss policy.

interface FixtureOptions {
  tables?: {
    name?: string
    range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
    headerRowCount: number
    totalsRowCount?: number
  }[]
  pivotRanges?: { startRow: number; endRow: number; startColumn: number; endColumn: number }[]
  visuals?: { sheetId: string }[]
  definedNames?: { name: string; formula: string; sheetIndex?: number }[]
  csvPath?: string
  filterOrigin?: {
    origin: 'worksheet' | 'table'
    range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
  }
}

function fixtureState(options: FixtureOptions = {}): {
  state: LazyWorkbookState
  journal: EditJournal
} {
  const journal = createEditJournal()
  const state = {
    file: {
      sessionId: 'session-1',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Alpha',
          tables: options.tables ?? [],
          pivotRanges: options.pivotRanges ?? [],
        },
        {
          id: 'sheet-2',
          name: 'Beta',
          tables: [],
          pivotRanges: [],
        },
      ],
      visuals: options.visuals ?? [],
      definedNames: options.definedNames ?? [],
      ...(options.csvPath === undefined ? {} : { csvPath: options.csvPath }),
    },
    editJournal: journal,
    filterOrigins: new Map(
      options.filterOrigin === undefined ? [] : [['sheet-1', options.filterOrigin]],
    ),
  } as unknown as LazyWorkbookState
  return { state, journal }
}

function runtimeWithDvRules(
  rulesBySheet: Record<string, { rule: { type?: string } }[]>,
): UniverRuntime {
  return {
    univerAPI: {
      getActiveWorkbook: () => ({
        getSheetBySheetId: (id: string) =>
          rulesBySheet[id] === undefined ? null : { getDataValidations: () => rulesBySheet[id] },
      }),
    },
  } as unknown as UniverRuntime
}

const cellArea = (
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
): { startRow: number; endRow: number; startColumn: number; endColumn: number } => ({
  startRow,
  endRow,
  startColumn,
  endColumn,
})

/// A 4-column table with a header row on top: rows 0-9, columns 0-3.
const DEFAULT_TABLE = {
  name: 'Table1',
  range: cellArea(0, 9, 0, 3),
  headerRowCount: 1,
}

function findings(state: LazyWorkbookState, runtime: UniverRuntime | null = null): string[] {
  return collectSaveCompatFindings({ state, runtime }).map((finding) => finding.id)
}

function firstDetail(state: LazyWorkbookState): string | undefined {
  const found = collectSaveCompatFindings({ state, runtime: null })
  return found[0]?.detail
}

function record(journal: EditJournal, op: StructuralJournalOp): void {
  recordStructuralOp(journal, 'sheet-1', op, 'Alpha')
}

describe('save-compat registry', () => {
  it('has unique feature ids', () => {
    const ids = SAVE_COMPAT_FEATURES.map((feature) => feature.id)
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('maps every feature to a zh-defined label key (i18n key set)', () => {
    const strings = zh as Record<string, string>
    for (const feature of SAVE_COMPAT_FEATURES) {
      expect(feature.labelKey in zh, feature.labelKey).toBe(true)
      expect(strings[feature.labelKey]?.length, feature.labelKey).toBeGreaterThan(0)
    }
  })

  it('anchors every feature to at least one fail-closed refusal site', () => {
    for (const feature of SAVE_COMPAT_FEATURES) {
      expect(feature.refs.length, feature.id).toBeGreaterThan(0)
      for (const ref of feature.refs) {
        expect(ref.startsWith('apps/sheets/'), ref).toBe(true)
      }
    }
  })

  it('marks renderer-invisible constructs as save-side-only (null detector)', () => {
    for (const feature of SAVE_COMPAT_FEATURES) {
      const rendererInvisible = feature.id === 'x14-cf-edit' || feature.id === 'x14-dv-edit'
      expect(feature.detect === null, feature.id).toBe(rendererInvisible)
    }
  })

  it('resolves label keys through the registry and rejects unknown ids', () => {
    const feature = SAVE_COMPAT_FEATURES[0]
    if (feature === undefined) throw new Error('registry must not be empty')
    expect(saveCompatLabelKey(feature.id)).toBe(feature.labelKey)
    expect(() => saveCompatLabelKey('no-such-feature' as never)).toThrow()
  })
})

describe('csv-flatten detector', () => {
  it('warns for CSV sessions and stays quiet for xlsx', () => {
    const csv = fixtureState({ csvPath: '/tmp/book.csv' })
    expect(findings(csv.state)).toContain('csv-flatten')
    const xlsx = fixtureState()
    expect(findings(xlsx.state)).not.toContain('csv-flatten')
  })
})

describe('multi-select-dv detector', () => {
  it('fires for a listMultiple rule on a dv-dirty sheet', () => {
    const { state, journal } = fixtureState()
    journal.dvDirty.add('sheet-1')
    const runtime = runtimeWithDvRules({ 'sheet-1': [{ rule: { type: 'listMultiple' } }] })
    const found = collectSaveCompatFindings({ state, runtime })
    expect(found.map((f) => f.id)).toContain('multi-select-dv')
    expect(found.find((f) => f.id === 'multi-select-dv')?.detail).toBe('Alpha')
  })

  it('stays quiet for saveable rule types and clean sheets', () => {
    const { state, journal } = fixtureState()
    journal.dvDirty.add('sheet-1')
    const runtime = runtimeWithDvRules({ 'sheet-1': [{ rule: { type: 'list' } }] })
    expect(findings(state, runtime)).not.toContain('multi-select-dv')
    expect(findings(fixtureState().state, runtimeWithDvRules({}))).not.toContain('multi-select-dv')
  })
})

describe('stranded-addition detector', () => {
  const pivotOn = (sheetId: string, sourceSheetId: string) => ({
    sheetId,
    sourceSheetId,
    sourceArea: cellArea(0, 5, 0, 1),
    location: cellArea(0, 3, 3, 4),
    name: 'Pivot1',
    fieldNames: ['A', 'B'],
    rowFieldIndices: [0],
    rowItems: ['x'],
  })

  it('fires for a pivot added on a session-added sheet with pending shifts', () => {
    const { state, journal } = fixtureState()
    recordSheetInsert(journal, 'sheet-new', 'New')
    journal.pivotAdds.push(pivotOn('sheet-new', 'sheet-1') as never)
    record(journal, { kind: 'insert-rows', index: 0, count: 1 })
    expect(findings(state)).toContain('stranded-addition')
  })

  it('fires for a table added on a session-added sheet with pending row/col ops', () => {
    const { state, journal } = fixtureState()
    recordSheetInsert(journal, 'sheet-new', 'New')
    journal.tableAdds.push({
      sheetId: 'sheet-new',
      area: cellArea(0, 3, 0, 1),
      name: 'Local',
      columnNames: ['A'],
      bandedRows: true,
    } as never)
    record(journal, { kind: 'remove-rows', index: 0, count: 1 })
    expect(findings(state)).toContain('stranded-addition')
  })

  it('fires when only sheet-management ops are pending (pivot hold)', () => {
    const { state, journal } = fixtureState()
    recordSheetInsert(journal, 'sheet-new', 'New')
    journal.pivotAdds.push(pivotOn('sheet-new', 'sheet-1') as never)
    journal.sheets.renamed.set('sheet-2', 'Beta2')
    expect(findings(state)).toContain('stranded-addition')
  })

  it('stays quiet when the additions ride original sheets (two-phase save resolves)', () => {
    const { state, journal } = fixtureState()
    journal.pivotAdds.push(pivotOn('sheet-1', 'sheet-2') as never)
    record(journal, { kind: 'insert-rows', index: 0, count: 1 })
    expect(findings(state)).not.toContain('stranded-addition')
  })

  it('stays quiet when there is nothing to hold back', () => {
    const { state, journal } = fixtureState()
    recordSheetInsert(journal, 'sheet-new', 'New')
    record(journal, { kind: 'insert-rows', index: 0, count: 1 })
    expect(findings(state)).not.toContain('stranded-addition')
  })
})

describe('structure-conflict detector', () => {
  it('refuses a whole-row move that relocates the table header row', () => {
    const { state, journal } = fixtureState({ tables: [DEFAULT_TABLE] })
    // Swap rows [0] with [1..4]: the header row lands inside data.
    record(journal, { kind: 'move-rows', index: 0, count: 1, before: 5 })
    expect(findings(state)).toContain('structure-conflict')
    expect(firstDetail(state)).toBe('Alpha')
  })

  it('allows relocating the whole table inside one swapped block', () => {
    const { state, journal } = fixtureState({ tables: [DEFAULT_TABLE] })
    record(journal, { kind: 'move-rows', index: 0, count: 10, before: 15 })
    expect(findings(state)).not.toContain('structure-conflict')
  })

  it('allows an interior data reorder under a table containing the swap', () => {
    const { state, journal } = fixtureState({ tables: [DEFAULT_TABLE] })
    // Swap [1,2] with [3,4]: fully inside the table, header/totals untouched.
    record(journal, { kind: 'move-rows', index: 1, count: 2, before: 5 })
    expect(findings(state)).not.toContain('structure-conflict')
  })

  it('refuses deleting the header row and allows deleting data rows', () => {
    const { state, journal } = fixtureState({ tables: [DEFAULT_TABLE] })
    record(journal, { kind: 'remove-rows', index: 0, count: 1 })
    expect(findings(state)).toContain('structure-conflict')

    const ok = fixtureState({ tables: [DEFAULT_TABLE] })
    record(ok.journal, { kind: 'remove-rows', index: 3, count: 2 })
    expect(findings(ok.state)).not.toContain('structure-conflict')
  })

  it('refuses a deletion that leaves the table without data rows', () => {
    const { state, journal } = fixtureState({ tables: [DEFAULT_TABLE] })
    record(journal, { kind: 'remove-rows', index: 1, count: 9 })
    expect(findings(state)).toContain('structure-conflict')
  })

  it('respects totals rows mirrored from the file metadata', () => {
    const table = { ...DEFAULT_TABLE, totalsRowCount: 1 }
    const { state, journal } = fixtureState({ tables: [table] })
    // Deleting the last row takes the totals row (table end row 9).
    record(journal, { kind: 'remove-rows', index: 9, count: 1 })
    expect(findings(state)).toContain('structure-conflict')
  })

  it('refuses a column move that would reorder table columns, allows wholesale', () => {
    const torn = fixtureState({ tables: [DEFAULT_TABLE] })
    record(torn.journal, { kind: 'move-cols', index: 1, count: 1, before: 3 })
    expect(findings(torn.state)).toContain('structure-conflict')

    const whole = fixtureState({ tables: [DEFAULT_TABLE] })
    record(whole.journal, { kind: 'move-cols', index: 0, count: 4, before: 6 })
    expect(findings(whole.state)).not.toContain('structure-conflict')
  })

  it('refuses any range-move contact with a table', () => {
    const hit = fixtureState({ tables: [DEFAULT_TABLE] })
    record(hit.journal, {
      kind: 'move-range',
      from: cellArea(5, 6, 0, 1),
      to: cellArea(2, 3, 0, 1),
    })
    expect(findings(hit.state)).toContain('structure-conflict')

    const miss = fixtureState({ tables: [DEFAULT_TABLE] })
    record(miss.journal, {
      kind: 'move-range',
      from: cellArea(20, 21, 5, 6),
      to: cellArea(25, 26, 5, 6),
    })
    expect(findings(miss.state)).not.toContain('structure-conflict')
  })

  it('refuses merging cells over a table', () => {
    const { state, journal } = fixtureState({ tables: [DEFAULT_TABLE] })
    record(journal, { kind: 'merge-cells', range: cellArea(4, 5, 1, 2) })
    expect(findings(state)).toContain('structure-conflict')
  })

  it('replays ops before judging later ones (no false alarm after a shift)', () => {
    const { state, journal } = fixtureState({ tables: [DEFAULT_TABLE] })
    // Five rows inserted above push the table to rows 5-14; a range move of
    // the now-empty rows 0-4 must NOT be judged against the stale rows 0-9.
    record(journal, { kind: 'insert-rows', index: 0, count: 5 })
    record(journal, { kind: 'move-range', from: cellArea(0, 4, 0, 3), to: cellArea(20, 24, 0, 3) })
    expect(findings(state)).not.toContain('structure-conflict')
  })

  it('still sees a conflict that only exists after replaying earlier ops', () => {
    const { state, journal } = fixtureState({ tables: [DEFAULT_TABLE] })
    // After the insert the table occupies rows 5-14; a move of rows 8-9 (data
    // rows in post-insert coordinates) overlaps the shifted table.
    record(journal, { kind: 'insert-rows', index: 0, count: 5 })
    record(journal, { kind: 'move-range', from: cellArea(8, 9, 0, 1), to: cellArea(30, 31, 0, 1) })
    expect(findings(state)).toContain('structure-conflict')
  })

  it('refuses a range move over the sheet auto-filter', () => {
    const { state, journal } = fixtureState({
      filterOrigin: { origin: 'worksheet', range: cellArea(2, 8, 0, 4) },
    })
    record(journal, { kind: 'move-range', from: cellArea(5, 6, 0, 1), to: cellArea(20, 21, 0, 1) })
    expect(findings(state)).toContain('structure-conflict')

    const miss = fixtureState({
      filterOrigin: { origin: 'worksheet', range: cellArea(2, 8, 0, 4) },
    })
    record(miss.journal, {
      kind: 'move-range',
      from: cellArea(30, 31, 0, 1),
      to: cellArea(35, 36, 0, 1),
    })
    expect(findings(miss.state)).not.toContain('structure-conflict')
  })

  it('ignores session-added and removed sheets', () => {
    const { state, journal } = fixtureState({ tables: [DEFAULT_TABLE] })
    journal.sheets.removed.add('sheet-1')
    record(journal, { kind: 'remove-rows', index: 0, count: 1 })
    expect(findings(state)).not.toContain('structure-conflict')
  })
})

describe('duplicate-carries-parts detector', () => {
  function duplicateOf(state: LazyWorkbookState, journal: EditJournal, sourceId: string): void {
    journal.sheets.added.set('sheet-copy', { name: 'Alpha copy', sourceSheetId: sourceId })
    void state
  }

  it('fires when the source carries a table', () => {
    const { state, journal } = fixtureState({ tables: [DEFAULT_TABLE] })
    duplicateOf(state, journal, 'sheet-1')
    const found = collectSaveCompatFindings({ state, runtime: null })
    expect(found.map((f) => f.id)).toContain('duplicate-carries-parts')
    expect(found.find((f) => f.id === 'duplicate-carries-parts')?.detail).toBe('Alpha copy')
  })

  it('fires for sources with pivots, visuals, or sheet-scoped names', () => {
    const pivot = fixtureState({ pivotRanges: [cellArea(0, 3, 5, 7)] })
    duplicateOf(pivot.state, pivot.journal, 'sheet-1')
    expect(findings(pivot.state)).toContain('duplicate-carries-parts')

    const visual = fixtureState({ visuals: [{ sheetId: 'sheet-1' }] })
    duplicateOf(visual.state, visual.journal, 'sheet-1')
    expect(findings(visual.state)).toContain('duplicate-carries-parts')

    const scoped = fixtureState({
      definedNames: [{ name: 'Local', formula: 'Alpha!$A$1', sheetIndex: 0 }],
    })
    duplicateOf(scoped.state, scoped.journal, 'sheet-1')
    expect(findings(scoped.state)).toContain('duplicate-carries-parts')
  })

  it('stays quiet for a plain source and unknown sources', () => {
    const { state, journal } = fixtureState()
    duplicateOf(state, journal, 'sheet-1')
    expect(findings(state)).not.toContain('duplicate-carries-parts')

    const missing = fixtureState()
    missing.journal.sheets.added.set('sheet-copy', { name: 'Ghost', sourceSheetId: 'sheet-x' })
    expect(findings(missing.state)).not.toContain('duplicate-carries-parts')
  })
})

describe('banner appear/disappear/dismiss policy', () => {
  const a: SaveCompatFinding = { id: 'csv-flatten' }
  const b: SaveCompatFinding = { id: 'multi-select-dv' }

  it('shows findings while nothing was dismissed', () => {
    expect(saveCompatVisible([a], null)).toEqual([a])
  })

  it('hides when the construct resolves (empty findings are always invisible)', () => {
    expect(saveCompatVisible([], null)).toEqual([])
    expect(saveCompatVisible([], saveCompatDismissKey([a]))).toEqual([])
  })

  it('hides exactly the dismissed set for the session', () => {
    expect(saveCompatVisible([a, b], saveCompatDismissKey([a, b]))).toEqual([])
  })

  it('re-shows when a different construct arrives after dismissing', () => {
    expect(saveCompatVisible([a, b], saveCompatDismissKey([a]))).toEqual([a, b])
  })

  it('keys are order-insensitive', () => {
    expect(saveCompatDismissKey([a, b])).toBe(saveCompatDismissKey([b, a]))
  })
})
