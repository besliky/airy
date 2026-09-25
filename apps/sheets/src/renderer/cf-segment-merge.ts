/**
 * Structural row/column edits are ref-shifted by Univer's
 * FormulaRefRangeService: rules that carry formulas (expression conditional
 * formats, validations with formula sources) come back as one segment per
 * affected band, and the conditional-formatting ref-range controller
 * re-emits every extra segment as a NEW rule. Excel keeps ONE rule whose
 * range extends or shrinks — and the split also drifts semantics: inserting
 * a row inside B1:B10 (=B1>50 anchored at B1) hands back the lower band
 * anchored at B3 testing B4>50, one row off for every covered cell
 * (BUG-1717).
 *
 * This module wraps registerRangeFormula so segments that exactly tile one
 * bounding box merge back into a single result: the ranges collapse to the
 * box and the formulas come from the segment still holding the original
 * anchor (the top-left of the pre-edit ranges) — precisely the rule Excel
 * keeps. Segments that do not tile one box (a range torn by a row move)
 * pass through untouched, preserving the upstream behavior for shapes this
 * module does not model.
 */
import type { IRange } from '@univerjs/core'
import { FormulaRefRangeService } from '@univerjs/sheets-formula'

import type { UniverRuntime } from './univer-state'

/// One registerRangeFormula result entry, in the shape both the
/// conditional-formatting and data-validation controllers consume.
export interface CfSegment {
  readonly ranges: readonly IRange[]
  readonly formulas: readonly string[]
}

/** Structural surface of FormulaRefRangeService this module relies on —
 * kept explicit so tests can drive a plain mock. */
export interface FormulaRefRangeServiceLike {
  registerRangeFormula(
    unitId: string,
    subUnitId: string,
    oldRanges: readonly IRange[],
    formulas: readonly string[],
    callback: (segments: readonly CfSegment[]) => unknown,
  ): unknown
}

const rangeArea = (range: IRange): number =>
  (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1)

const isDegenerate = (range: IRange): boolean =>
  range.startRow < 0 ||
  range.startColumn < 0 ||
  range.endRow < range.startRow ||
  range.endColumn < range.startColumn

const rangesIntersect = (a: IRange, b: IRange): boolean =>
  a.startRow <= b.endRow &&
  b.startRow <= a.endRow &&
  a.startColumn <= b.endColumn &&
  b.startColumn <= a.endColumn

const topLeft = (range: IRange): number => range.startRow * 16_384 + range.startColumn

/**
 * Merges ref-shift segments that tile exactly one bounding box back into a
 * single segment. `anchor` is the pre-edit top-left the original rule's
 * formula was relative to; the segment still covering it keeps its formulas
 * (they are the ones Excel preserves — the other segments' re-anchored
 * formulas drift by the inserted/deleted band). Anything that is not a
 * clean tiling (gaps, overlaps, mismatched formula lists) passes through
 * unchanged.
 */
export function mergeCfRuleSegments(
  segments: readonly CfSegment[],
  anchor: { readonly row: number; readonly column: number } | null,
): CfSegment[] {
  if (!Array.isArray(segments) || segments.length < 2) return [...(segments ?? [])]
  // Segments whose formula lists differ in length cannot come from one
  // original rule — refuse to guess.
  const formulaCount = segments[0]?.formulas.length
  for (const segment of segments) {
    if (!Array.isArray(segment.ranges) || segment.ranges.length === 0) return [...segments]
    if (!Array.isArray(segment.formulas) || segment.formulas.length !== formulaCount) {
      return [...segments]
    }
    for (const range of segment.ranges) {
      if (!range || isDegenerate(range)) return [...segments]
    }
  }
  const flat = segments.flatMap((segment) => segment.ranges)
  for (let i = 0; i < flat.length; i += 1) {
    for (let j = i + 1; j < flat.length; j += 1) {
      if (rangesIntersect(flat[i] as IRange, flat[j] as IRange)) return [...segments]
    }
  }
  const box: IRange = {
    startRow: Math.min(...flat.map((range) => range.startRow)),
    endRow: Math.max(...flat.map((range) => range.endRow)),
    startColumn: Math.min(...flat.map((range) => range.startColumn)),
    endColumn: Math.max(...flat.map((range) => range.endColumn)),
  }
  // The segments must tile the box exactly: no gaps (a torn range stays
  // split) and no over-cover (sum of areas equals the box area only when
  // every cell is covered once).
  if (flat.reduce((sum, range) => sum + rangeArea(range), 0) !== rangeArea(box)) {
    return [...segments]
  }
  // The original anchor's segment survives verbatim upstream, so its
  // formulas are the shifted originals; any segment still covering the
  // anchor point qualifies (the first by top-left otherwise).
  const anchorSegment =
    (anchor &&
      segments.find((segment) =>
        segment.ranges.some(
          (range: IRange) =>
            range.startRow <= anchor.row &&
            anchor.row <= range.endRow &&
            range.startColumn <= anchor.column &&
            anchor.column <= range.endColumn,
        ),
      )) ||
    [...segments].sort((a, b) => topLeft(a.ranges[0] as IRange) - topLeft(b.ranges[0] as IRange))[0]
  return [{ ranges: [box], formulas: [...(anchorSegment?.formulas ?? [])] }]
}

/** Resolve the injector's FormulaRefRangeService singleton and wrap its
 * registerRangeFormula so every consumer's segments flow through the merge. */
export function installCfRefSegmentMerge(runtime: UniverRuntime): { dispose(): void } {
  let service: FormulaRefRangeServiceLike
  try {
    service = runtime.univer
      .__getInjector()
      .get(FormulaRefRangeService) as unknown as FormulaRefRangeServiceLike
  } catch {
    return { dispose() {} } // formula plugin not installed in this runtime
  }
  if (typeof service.registerRangeFormula !== 'function') return { dispose() {} }
  const original = service.registerRangeFormula.bind(service)
  service.registerRangeFormula = (unitId, subUnitId, oldRanges, formulas, callback) =>
    original(unitId, subUnitId, oldRanges, formulas, (segments) =>
      callback(
        mergeCfRuleSegments(
          segments,
          oldRanges[0] ? { row: oldRanges[0].startRow, column: oldRanges[0].startColumn } : null,
        ),
      ),
    )
  return {
    dispose() {
      service.registerRangeFormula = original
    },
  }
}
