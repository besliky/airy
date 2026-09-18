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
})
