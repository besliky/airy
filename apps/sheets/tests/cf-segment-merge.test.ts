/**
 * BUG-1717 regression: a row inserted inside an expression conditional-
 * formatting range must keep ONE rule whose range extends, not a pair of
 * anchor-drifted duplicates. Covers the segment merge itself (including the
 * audited B1:B2 + B3:B11 shape) and the wrapper end-to-end: a real Univer
 * runtime's FormulaRefRangeService hands back one merged segment when the
 * insert-row command lands inside a registered formula range.
 */
import {
  ICommandService,
  IUniverInstanceService,
  LocaleType,
  UniverInstanceType,
} from '@univerjs/core'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { FormulaRefRangeService, UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import '@univerjs/sheets/lib/facade'
import { describe, expect, it } from 'vitest'

import {
  installCfRefSegmentMerge,
  mergeCfRuleSegments,
  type CfSegment,
} from '../src/renderer/cf-segment-merge'
import { createUniver } from '../src/renderer/create-univer'

const range = (startRow: number, endRow: number, startColumn = 1, endColumn = startColumn) => ({
  startRow,
  endRow,
  startColumn,
  endColumn,
})

describe('mergeCfRuleSegments', () => {
  it('merges the audited split back into one anchor-true rule', () => {
    // Live shape from the audit: insert a row inside B1:B10 (=B1>50) split
    // the rule into B1:B2 (formula unchanged) and B3:B11 re-anchored to
    // B4>50 — one row off Excel's semantics for every covered cell.
    const segments: CfSegment[] = [
      { ranges: [range(0, 1)], formulas: ['=B1>50'] },
      { ranges: [range(2, 10)], formulas: ['=B4>50'] },
    ]
    expect(mergeCfRuleSegments(segments, { row: 0, column: 1 })).toEqual([
      { ranges: [range(0, 10)], formulas: ['=B1>50'] },
    ])
  })

  it('merges three bands with verbatim formulas into the box', () => {
    // Headless shape for the same insert: the service hands back three
    // non-overlapping bands (anchor row, inserted band, shifted tail) with
    // the formula text kept verbatim.
    const segments: CfSegment[] = [
      { ranges: [range(0, 0)], formulas: ['B1>50'] },
      { ranges: [range(3, 10)], formulas: ['B1>50'] },
      { ranges: [range(1, 2)], formulas: ['B1>50'] },
    ]
    expect(mergeCfRuleSegments(segments, { row: 0, column: 1 })).toEqual([
      { ranges: [range(0, 10)], formulas: ['B1>50'] },
    ])
  })

  it('prefers the segment still covering the original anchor', () => {
    const segments: CfSegment[] = [
      { ranges: [range(4, 9, 3, 3)], formulas: ['=D5<>1'] },
      { ranges: [range(0, 3, 3, 3)], formulas: ['=D1<>1'] },
    ]
    expect(mergeCfRuleSegments(segments, { row: 0, column: 3 })).toEqual([
      { ranges: [range(0, 9, 3, 3)], formulas: ['=D1<>1'] },
    ])
  })

  it('keeps a torn range split (gap between segments)', () => {
    const segments: CfSegment[] = [
      { ranges: [range(0, 1)], formulas: ['=B1>50'] },
      { ranges: [range(4, 9)], formulas: ['=B4>50'] },
    ]
    expect(mergeCfRuleSegments(segments, { row: 0, column: 1 })).toEqual(segments)
  })

  it('keeps segments that overlap or leave a hole untouched', () => {
    const overlapping: CfSegment[] = [
      { ranges: [range(0, 3)], formulas: ['f'] },
      { ranges: [range(2, 5)], formulas: ['f'] },
      { ranges: [range(6, 9)], formulas: ['f'] },
    ]
    expect(mergeCfRuleSegments(overlapping, { row: 0, column: 1 })).toEqual(overlapping)
    const hole: CfSegment[] = [
      { ranges: [range(0, 3, 1, 2)], formulas: ['f'] },
      { ranges: [range(0, 9, 3, 3)], formulas: ['f'] },
    ]
    expect(mergeCfRuleSegments(hole, { row: 0, column: 1 })).toEqual(hole)
  })

  it('refuses segments with mismatched formula lists', () => {
    const segments: CfSegment[] = [
      { ranges: [range(0, 1)], formulas: ['=B1>50', '=B1<9'] },
      { ranges: [range(2, 10)], formulas: ['=B4>50'] },
    ]
    expect(mergeCfRuleSegments(segments, { row: 0, column: 1 })).toEqual(segments)
  })

  it('passes a single segment through and tolerates empty input', () => {
    const single: CfSegment[] = [{ ranges: [range(0, 10)], formulas: ['=B1>50'] }]
    expect(mergeCfRuleSegments(single, { row: 0, column: 1 })).toEqual(single)
    expect(mergeCfRuleSegments([], { row: 0, column: 1 })).toEqual([])
  })
})

interface RuntimeFixture {
  readonly runtime: ReturnType<typeof createUniver>
  readonly service: FormulaRefRangeService
  readonly commandService: ICommandService
  readonly disposeRuntime: () => void
}

function runtimeWithFormulaEngine(id: string): RuntimeFixture {
  const runtime = createUniver({
    locale: LocaleType.EN_US,
    locales: {},
    presets: [
      { plugins: [UniverFormulaEnginePlugin] },
      { plugins: [UniverSheetsPlugin] },
      { plugins: [UniverSheetsFormulaPlugin] },
    ],
  })
  runtime.univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id,
    name: 'wb',
    styles: {},
    sheets: {
      main: { id: 'main', name: 'Main', rowCount: 40, columnCount: 8, cellData: {} },
    },
  })
  runtime.univer.__getInjector().get(IUniverInstanceService).focusUnit(id)
  return {
    runtime,
    service: runtime.univer.__getInjector().get(FormulaRefRangeService),
    commandService: runtime.univer.__getInjector().get(ICommandService),
    disposeRuntime: () => runtime.univer.dispose(),
  }
}

const insertRowInside = (fixture: RuntimeFixture, unitId: string): Promise<boolean> =>
  fixture.commandService.executeCommand('sheet.command.insert-row', {
    unitId,
    subUnitId: 'main',
    range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 7 },
  })

describe('installCfRefSegmentMerge across an insert-row', () => {
  it('collapses the ref-shift bands of one formula range into one segment', async () => {
    const fixture = runtimeWithFormulaEngine('wb-merged')
    try {
      installCfRefSegmentMerge(fixture.runtime)
      const outcomes: unknown[] = []
      fixture.service.registerRangeFormula(
        'wb-merged',
        'main',
        [range(0, 9)],
        ['B1>50'],
        (segments) => {
          outcomes.push(segments)
          return { undos: [], redos: [] }
        },
      )
      await insertRowInside(fixture, 'wb-merged')
      expect(outcomes).toEqual([[{ ranges: [range(0, 10)], formulas: ['B1>50'] }]])
    } finally {
      fixture.disposeRuntime()
    }
  })

  it('without the wrapper the same insert hands back several segments', async () => {
    const fixture = runtimeWithFormulaEngine('wb-split')
    try {
      const outcomes: unknown[] = []
      fixture.service.registerRangeFormula(
        'wb-split',
        'main',
        [range(0, 9)],
        ['B1>50'],
        (segments) => {
          outcomes.push(segments)
          return { undos: [], redos: [] }
        },
      )
      await insertRowInside(fixture, 'wb-split')
      const segments = outcomes[0] as CfSegment[] | undefined
      // Documents the upstream split the merge undoes: more than one band
      // for a single registered rule.
      expect((segments ?? []).length).toBeGreaterThan(1)
    } finally {
      fixture.disposeRuntime()
    }
  })
})
