/**
 * The main process parses every save through workbookSaveRequestSchema, so its
 * "at least one edit" refine must exempt explicit Save As — a clean workbook
 * saved to a new path is a valid request with nothing to apply.
 */
import { describe, expect, it } from 'vitest'
import { workbookSaveRequestSchema } from '../src/shared/desktop-api'

function emptyRequest(mode: 'save' | 'save-as') {
  return {
    sessionId: '5d4f6f7a-1c2b-4e3d-9a8f-0b1c2d3e4f5a',
    mode,
    edits: [],
    structuralOps: [],
    chartEdits: [],
    visualEdits: [],
    visualAdditions: [],
    tableAdditions: [],
    pivotAdditions: [],
    sheetOps: [],
    sheetOrder: [],
    filterStates: [],
    hyperlinkEdits: [],
    cfStates: [],
    dvStates: [],
    pageSetupStates: [],
    noteStates: [],
    formulaValues: [],
    pivotCacheRefreshPaths: [],
    pivotRefreshUpdates: [],
    sheetProtections: [],
    sparklineAdditions: [],
    definedNamesState: null,
  }
}

describe('workbookSaveRequestSchema', () => {
  it('accepts an empty save-as request', () => {
    expect(() => workbookSaveRequestSchema.parse(emptyRequest('save-as'))).not.toThrow()
  })

  it('still rejects an empty ordinary save', () => {
    expect(() => workbookSaveRequestSchema.parse(emptyRequest('save'))).toThrow(/at least one edit/)
  })

  it('accepts a save whose edits arrive via a chunked transfer', () => {
    const request = {
      ...emptyRequest('save'),
      editsTransferId: '0f9e8d7c-6b5a-4c3d-8e2f-1a0b9c8d7e6f',
    }
    expect(() => workbookSaveRequestSchema.parse(request)).not.toThrow()
  })

  it('accepts a same-size range move and rejects a resized or off-sheet one', () => {
    const area = (startRow: number, startColumn: number, endRow: number, endColumn: number) => ({
      startRow,
      startColumn,
      endRow,
      endColumn,
    })
    const request = {
      ...emptyRequest('save'),
      structuralOps: [
        { sheetId: 'sh1', kind: 'move-range', from: area(0, 0, 1, 2), to: area(4, 5, 5, 7) },
      ],
    }
    expect(() => workbookSaveRequestSchema.parse(request)).not.toThrow()
    expect(() =>
      workbookSaveRequestSchema.parse({
        ...request,
        structuralOps: [
          { sheetId: 'sh1', kind: 'move-range', from: area(0, 0, 1, 2), to: area(4, 5, 5, 6) },
        ],
      }),
    ).toThrow(/equally sized/)
    expect(() =>
      workbookSaveRequestSchema.parse({
        ...request,
        structuralOps: [
          {
            sheetId: 'sh1',
            kind: 'move-range',
            from: area(0, 0, 1, 2),
            to: area(1_048_575, 5, 1_048_576, 7),
          },
        ],
      }),
    ).toThrow()
  })

  it('rejects moves whose block runs past the sheet edge (index+count)', () => {
    // index and count are each in bounds, but the block they describe must
    // still fit inside the sheet — a renderer-side op with the sum over the
    // edge would swap blocks past the last row/column.
    const request = (op: Record<string, unknown>) => ({
      ...emptyRequest('save'),
      structuralOps: [{ sheetId: 'sh1', ...op }],
    })
    expect(() =>
      workbookSaveRequestSchema.parse(
        request({ kind: 'move-rows', index: 1_048_575, count: 2, before: 0 }),
      ),
    ).toThrow(/fit inside the sheet/)
    expect(() =>
      workbookSaveRequestSchema.parse(request({ kind: 'move-cols', index: 16_383, count: 2, before: 0 })),
    ).toThrow(/fit inside the sheet/)
    // The largest block that exactly reaches the edge stays valid.
    expect(() =>
      workbookSaveRequestSchema.parse(
        request({ kind: 'move-rows', index: 1_048_575, count: 1, before: 0 }),
      ),
    ).not.toThrow()
    expect(() =>
      workbookSaveRequestSchema.parse(
        request({ kind: 'move-cols', index: 16_383, count: 1, before: 0 }),
      ),
    ).not.toThrow()
  })
})
