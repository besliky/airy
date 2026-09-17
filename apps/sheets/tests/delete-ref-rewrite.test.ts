/**
 * Excel-compatible #REF! handling for row/column deletion: the deletion must
 * succeed and orphaned references must rewrite (#REF! / clipped) instead of
 * being refused. Univer owns same-sheet rewriting in the live model; these
 * tests cover the renderer's cross-sheet pass — the collect/rewrite/apply
 * split the BeforeCommandExecute gate drives — plus the shared token
 * rewriter's semantics. The gateway's own emission is covered in
 * xlsx-structure.test.ts.
 */
import { describe, expect, it, vi } from 'vitest'

import {
  applyCrossSheetRewrites,
  collectCrossSheetDependentRewrites,
  deletionLanded,
  deleteSpanSpec,
  finishCrossSheetRewrites,
  rewriteFormulaForDeletedSpan,
  rewriteHarvestedFormulaTexts,
  type RewriteWorkbook,
} from '../src/renderer/delete-ref-rewrite'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

function state(overrides: Record<string, unknown> = {}): LazyWorkbookState {
  return {
    file: {
      sessionId: 'session-1',
      sheets: [
        { id: 'sh1', name: 'Data', rowCount: 100, columnCount: 26, pivotRanges: [] },
        { id: 'sh2', name: 'Other', rowCount: 100, columnCount: 26, pivotRanges: [] },
      ],
      visuals: [],
    },
    editJournal: {
      cells: new Map(),
      structuralOps: new Map(),
      sheets: { added: new Set(), removed: new Set() },
    },
    formulaText: new Map(),
    ...overrides,
  } as unknown as LazyWorkbookState
}

function specFor(operation: Parameters<typeof deleteSpanSpec>[2]) {
  return deleteSpanSpec(state(), (id) => (id === 'sh1' ? 'Data' : 'Other'), operation)
}

/// Deleting Data!D:E (1-based D, two columns) — 0-based columns 3..4.
const delDE = specFor({ op: 'delete_cols', sheetId: 'sh1', column: 'D', count: 2 })

/// Deleting Data rows 7-8 (1-based) — 0-based rows 6..7.
const delRows78 = specFor({ op: 'delete_rows', sheetId: 'sh1', row: 7, count: 2 })

describe('rewriteFormulaForDeletedSpan', () => {
  it('turns a wholly-inside range into #REF!', () => {
    expect(rewriteFormulaForDeletedSpan('=SUM($D$7:$D$9)', delDE, false)).toBe('=SUM(#REF!)')
    expect(rewriteFormulaForDeletedSpan('=SUM(D3:E5)', delDE, false)).toBe('=SUM(#REF!)')
  })

  it('clips partially overlapping ranges', () => {
    // column deletion clips the column span only — the row span is untouched;
    // the surviving old-F column lands at D after D:E collapse
    expect(rewriteFormulaForDeletedSpan('=SUM(C7:E9)', delDE, false)).toBe('=SUM(C7:C9)')
    expect(rewriteFormulaForDeletedSpan('=SUM(D7:F9)', delDE, false)).toBe('=SUM(D7:D9)')
    // a row deletion clips the row span; ranges above stay put
    expect(rewriteFormulaForDeletedSpan('=SUM(A5:A8)', delRows78, false)).toBe('=SUM(A5:A6)')
    expect(rewriteFormulaForDeletedSpan('=SUM($A$7:$C$7)', delRows78, false)).toBe('=SUM(#REF!)')
    expect(rewriteFormulaForDeletedSpan('=SUM($A$1:$C$7)', delRows78, false)).toBe(
      '=SUM($A$1:$C$6)',
    )
  })

  it('shifts references beyond the span', () => {
    expect(rewriteFormulaForDeletedSpan('=SUM(F7:F9)', delDE, false)).toBe('=SUM(D7:D9)')
    expect(rewriteFormulaForDeletedSpan('=A9+A10', delRows78, false)).toBe('=A7+A8')
  })

  it('rewrites cross-sheet references to a bare #REF! token', () => {
    expect(rewriteFormulaForDeletedSpan('=SUM(Data!$D$7:$D$9)', delDE, true)).toBe('=SUM(#REF!)')
    expect(rewriteFormulaForDeletedSpan('=Data!D7+Data!G7', delDE, true)).toBe('=#REF!+Data!E7')
    expect(rewriteFormulaForDeletedSpan('=SUM(Data!C7:E9)', delDE, true)).toBe('=SUM(Data!C7:C9)')
  })

  it('leaves other sheets and string literals alone in qualified mode', () => {
    expect(rewriteFormulaForDeletedSpan('=SUM(Other!D7:E9)', delDE, true)).toBeNull()
    // D7/E7 inside quotes are literals; only a real reference rewrites
    expect(rewriteFormulaForDeletedSpan('="D7"-"E7"', delDE, false)).toBeNull()
    expect(rewriteFormulaForDeletedSpan('="D7"&E7', delDE, false)).toBe('="D7"&#REF!')
  })

  it('returns null when nothing changes', () => {
    expect(rewriteFormulaForDeletedSpan('=SUM(A1:A2)', delDE, false)).toBeNull()
  })
})

function scanWorkbook(
  formulasBySheet: Record<string, string[][]>,
  dims: Record<string, { rows?: number; columns?: number }> = {},
): RewriteWorkbook {
  return {
    getSheets: () =>
      Object.entries(formulasBySheet).map(([sheetId, rows]) => ({
        getSheetId: () => sheetId,
        getSheetName: () => (sheetId === 'sh1' ? 'Data' : 'Other'),
        getMaxRows: () => dims[sheetId]?.rows ?? 20,
        getMaxColumns: () => dims[sheetId]?.columns ?? 8,
        getRange: () => ({ getFormulas: () => rows }),
      })),
  }
}

describe('collectCrossSheetDependentRewrites', () => {
  it('collects qualified references on other sheets only', () => {
    const rewrites = collectCrossSheetDependentRewrites(
      state(),
      scanWorkbook({
        sh1: [['=SUM(D7:E9)']],
        sh2: [
          ['=Data!D7+1', '=SUM(Data!C7:E9)'],
          ['', '=Other!D7'],
        ],
      }),
      delDE,
    )
    expect(rewrites).toEqual([
      { sheetId: 'sh2', row: 0, column: 0, formula: '=#REF!+1' },
      { sheetId: 'sh2', row: 0, column: 1, formula: '=SUM(Data!C7:C9)' },
    ])
  })

  it('clips partial overlaps against the deleted span', () => {
    const rewrites = collectCrossSheetDependentRewrites(
      state(),
      scanWorkbook({ sh2: [['=SUM(Data!D7:F9)']] }),
      delDE,
    )
    // deleting Data columns D:E leaves old F, which lands at D
    expect(rewrites).toEqual([{ sheetId: 'sh2', row: 0, column: 0, formula: '=SUM(Data!D7:D9)' }])
  })

  it('skips the deleted sheet itself (Univer rewrites it) and removed sheets', () => {
    const removed = state({
      editJournal: {
        cells: new Map(),
        structuralOps: new Map(),
        sheets: { added: new Set(), removed: new Set(['sh2']) },
      },
    })
    const rewrites = collectCrossSheetDependentRewrites(
      removed,
      scanWorkbook({ sh1: [['=D7']], sh2: [['=Data!D7']] }),
      delDE,
    )
    expect(rewrites).toEqual([])
  })
})

describe('applyCrossSheetRewrites', () => {
  it('writes each rewrite through a synchronous set-range-values command (journaled, undoable)', () => {
    const commands: Array<{ id: string; params: Record<string, unknown> }> = []
    const runtime = {
      univer: {
        __getInjector: () => ({
          get: () => ({
            syncExecuteCommand: (id: string, params?: Record<string, unknown>) => {
              commands.push({ id, params: params ?? {} })
            },
          }),
        }),
      },
      univerAPI: { getActiveWorkbook: () => ({ getId: () => 'file-sha' }) },
    }
    applyCrossSheetRewrites(runtime, [
      { sheetId: 'sh2', row: 4, column: 0, formula: '=#REF!' },
      { sheetId: 'sh2', row: 4, column: 1, formula: '=SUM(Data!C7)' },
    ])
    expect(commands).toEqual([
      {
        id: 'sheet.command.set-range-values',
        params: {
          unitId: 'file-sha',
          subUnitId: 'sh2',
          range: { startRow: 4, endRow: 4, startColumn: 0, endColumn: 0 },
          value: { 4: { 0: { f: '=#REF!' } } },
        },
      },
      {
        id: 'sheet.command.set-range-values',
        params: {
          unitId: 'file-sha',
          subUnitId: 'sh2',
          range: { startRow: 4, endRow: 4, startColumn: 1, endColumn: 1 },
          value: { 4: { 1: { f: '=SUM(Data!C7)' } } },
        },
      },
    ])
  })

  it('does nothing without an active workbook', () => {
    const runtime = {
      univer: {
        __getInjector: () => ({
          get: () => ({
            syncExecuteCommand: () => {
              throw new Error('must not run')
            },
          }),
        }),
      },
      univerAPI: { getActiveWorkbook: () => undefined },
    }
    expect(() =>
      applyCrossSheetRewrites(runtime, [{ sheetId: 'sh2', row: 0, column: 0, formula: '=#REF!' }]),
    ).not.toThrow()
  })
})

describe('deletionLanded', () => {
  it('is true when the sheet shrank by exactly the deleted span', () => {
    expect(deletionLanded(scanWorkbook({ sh1: [] }, { sh1: { rows: 18 } }), delRows78, 20, 2)).toBe(
      true,
    )
    expect(deletionLanded(scanWorkbook({ sh1: [] }, { sh1: { columns: 6 } }), delDE, 8, 2)).toBe(
      true,
    )
  })

  it('is false when the model is unchanged (declined command) or moved otherwise', () => {
    // Univer emits CommandExecuted even when the handler declines
    // (protected sheet, invalid range) — the model never moved.
    expect(deletionLanded(scanWorkbook({ sh1: [] }), delRows78, 20, 2)).toBe(false)
    expect(deletionLanded(scanWorkbook({ sh1: [] }, { sh1: { rows: 19 } }), delRows78, 20, 2)).toBe(
      false,
    )
    expect(deletionLanded(scanWorkbook({ sh1: [] }, { sh1: { rows: 20 } }), delRows78, 20, 2)).toBe(
      false,
    )
  })

  it('is false without the sheet or a real span', () => {
    expect(deletionLanded(scanWorkbook({ sh2: [] }), delRows78, 20, 2)).toBe(false)
    expect(deletionLanded(scanWorkbook({ sh1: [] }, { sh1: { rows: 20 } }), delRows78, 20, 0)).toBe(
      false,
    )
  })
})

describe('finishCrossSheetRewrites', () => {
  function spyRuntime(): {
    runtime: Parameters<typeof applyCrossSheetRewrites>[0]
    commands: Array<{ id: string; params: Record<string, unknown> }>
  } {
    const commands: Array<{ id: string; params: Record<string, unknown> }> = []
    return {
      commands,
      runtime: {
        univer: {
          __getInjector: () => ({
            get: () => ({
              syncExecuteCommand: (id: string, params?: Record<string, unknown>) => {
                commands.push({ id, params: params ?? {} })
              },
            }),
          }),
        },
        univerAPI: { getActiveWorkbook: () => ({ getId: () => 'file-sha' }) },
      },
    }
  }

  it('drops the rewrite when the deletion never landed (declined command)', () => {
    const debug = vi.spyOn(console, 'debug').mockImplementation(() => {})
    try {
      const { runtime, commands } = spyRuntime()
      const harvested = new Map([['sh2', new Map([['6:0', '=Data!D7']])]])
      const st = state({ formulaText: harvested })
      // the model still has its original 20 rows — the command was declined
      const applied = finishCrossSheetRewrites({
        runtime: runtime as never,
        state: st,
        workbook: scanWorkbook({ sh2: [['=Data!D7']] }),
        spec: delRows78,
        countBefore: 20,
        spanCount: 2,
      })
      expect(applied).toBe(false)
      expect(commands).toEqual([])
      expect(harvested.get('sh2')?.get('6:0')).toBe('=Data!D7')
    } finally {
      debug.mockRestore()
    }
  })

  it('applies journaled rewrites and harvested-text updates once landed', () => {
    const { runtime, commands } = spyRuntime()
    const harvested = new Map([['sh2', new Map([['6:0', '=Data!D7']])]])
    const st = state({ formulaText: harvested })
    const applied = finishCrossSheetRewrites({
      runtime: runtime as never,
      state: st,
      // rows shrank 20 → 18 as requested
      workbook: scanWorkbook({ sh1: [], sh2: [['=Data!D7']] }, { sh1: { rows: 18 } }),
      spec: delRows78,
      countBefore: 20,
      spanCount: 2,
    })
    expect(applied).toBe(true)
    expect(commands.map((c) => c.id)).toEqual(['sheet.command.set-range-values'])
    expect(harvested.get('sh2')?.get('6:0')).toBe('=#REF!')
  })
})

describe('rewriteHarvestedFormulaTexts', () => {
  it('rewrites harvested texts on other sheets in place', () => {
    const harvested = new Map([
      ['sh1', new Map([['6:0', '=SUM(D7:E9)']])],
      [
        'sh2',
        new Map([
          ['6:0', '=Data!D7'],
          ['0:0', '=SUM(A1:A2)'],
        ]),
      ],
    ])
    const st = state({ formulaText: harvested })
    rewriteHarvestedFormulaTexts(st, delDE)
    expect(harvested.get('sh2')?.get('6:0')).toBe('=#REF!')
    expect(harvested.get('sh2')?.get('0:0')).toBe('=SUM(A1:A2)')
    // the deleted sheet's index goes dark after its structural op, not rewritten here
    expect(harvested.get('sh1')?.get('6:0')).toBe('=SUM(D7:E9)')
  })

  it('leaves keys superseded by a journal entry alone', () => {
    const harvested = new Map([['sh2', new Map([['6:0', '=Data!D7']])]])
    const st = state({
      formulaText: harvested,
      editJournal: {
        cells: new Map([['sh2', new Map([['6:0', { row: 6, column: 0, hasValue: false }]])]]),
        structuralOps: new Map(),
        sheets: { added: new Set(), removed: new Set() },
      },
    })
    rewriteHarvestedFormulaTexts(st, delDE)
    expect(harvested.get('sh2')?.get('6:0')).toBe('=Data!D7')
  })
})

describe('deleteSpanSpec', () => {
  it('builds the 0-based deleted span for rows and columns', () => {
    expect(delRows78.shift as { boundary: number; delta: number }).toEqual({
      boundary: 6,
      delta: -2,
      deleted: { start: 6, end: 7 },
    })
    expect(delRows78.axis).toBe('row')
    expect(delRows78.deletedSheetName).toBe('Data')
    expect(delDE.axis).toBe('column')
    expect((delDE.shift as { deleted: { start: number; end: number } }).deleted).toEqual({
      start: 3,
      end: 4,
    })
  })

  it('falls back to the live sheet name for session-added sheets', () => {
    const spec = deleteSpanSpec(state(), () => 'LiveName', {
      op: 'delete_rows',
      sheetId: 'shX',
      row: 2,
      count: 1,
    })
    expect(spec.deletedSheetName).toBe('LiveName')
  })
})
