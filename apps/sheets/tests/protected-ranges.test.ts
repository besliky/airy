import { describe, expect, it } from 'vitest'

import { mapProtectedRanges } from '../src/renderer/protected-ranges'

const range = (name: string, sqref: string) => ({ name, sqref, hasPassword: false })

describe('mapProtectedRanges', () => {
  it('shifts areas through row inserts and column removals', () => {
    const mapped = mapProtectedRanges(
      [range('Data', 'B3:D6'), range('Cell', 'B2')],
      [
        { kind: 'insert-rows', index: 1, count: 2 },
        { kind: 'remove-cols', index: 0, count: 1 },
      ],
    )
    expect(mapped).toEqual([range('Data', 'A5:C8'), range('Cell', 'A4')])
  })

  it('shrinks partially deleted areas and drops fully deleted ones', () => {
    const mapped = mapProtectedRanges(
      [range('Shrinks', 'A2:A5'), range('Gone', 'A3:A4'), range('Multi', 'A1 A3:A4')],
      [{ kind: 'remove-rows', index: 2, count: 2 }],
    )
    expect(mapped).toEqual([range('Shrinks', 'A2:A3'), range('Multi', 'A1')])
  })

  it('splits a partially moved range into exact runs instead of the envelope', () => {
    // Rows 2-3 (file) move to the tail: file rows 1,2,3 land on screen 4,5,1.
    const mapped = mapProtectedRanges(
      [range('Data', 'A2:A4')],
      [{ kind: 'move-rows', index: 1, count: 2, before: 6 }],
    )
    expect(mapped).toEqual([range('Data', 'A2 A5:A6')])
  })

  it('splits a partially moved column range into exact column runs', () => {
    // Columns 2-3 (file) move to the tail: file columns 1,2,3 land on screen
    // 4,5,1 — the envelope (B:E) would whitelist the unrelated screen C-D too.
    const mapped = mapProtectedRanges(
      [range('Data', 'B2:D2')],
      [{ kind: 'move-cols', index: 1, count: 2, before: 6 }],
    )
    expect(mapped).toEqual([range('Data', 'B2 E2:F2')])
  })

  it('crosses exact row runs with exact column runs under both move kinds', () => {
    // Row 1 and column A each relocate before line 5: file rows 2-4 land on
    // screens 1-3 (row 1 sits at screen 4), and file columns A-C land on
    // screens A-B plus E. The image is the cross product of the runs.
    const mapped = mapProtectedRanges(
      [range('Data', 'A2:C4')],
      [
        { kind: 'move-rows', index: 0, count: 1, before: 5 },
        { kind: 'move-cols', index: 0, count: 1, before: 5 },
      ],
    )
    expect(mapped).toEqual([range('Data', 'A1:B3 E1:E3')])
  })

  it('returns the input unchanged without ops and keeps unparseable parts', () => {
    const untouched = [range('Data', 'B3:D6')]
    expect(mapProtectedRanges(untouched, [])).toEqual(untouched)
    expect(
      mapProtectedRanges([range('Odd', 'NotARef')], [{ kind: 'insert-rows', index: 0, count: 1 }]),
    ).toEqual([range('Odd', 'NotARef')])
  })
})
