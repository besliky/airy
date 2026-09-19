/// Outline (group) model shared by the Data-tab commands, the +/- gutter,
/// and the Outline Settings dialog. Univer has no outline model of its own
/// — levels and collapsed flags live in the streamed workbook state and
/// journal declaratively — so the group math lives here, pure and testable.

export interface OutlineEntry {
  readonly level: number
  readonly collapsed: boolean
}

/// One outline group: a maximal contiguous run of lines at or below
/// `level`, plus the summary line that carries its +/- button (the line
/// after the run when summary is below/right, the line before otherwise).
export interface OutlineGroup {
  readonly axis: 'rows' | 'cols'
  /// The group's outline level (1-7); the button sits in the level-1 lane
  /// for level-1 groups, one lane further out per deeper level.
  readonly level: number
  readonly start: number
  readonly end: number
  readonly summary: number
}

export const MAX_OUTLINE_LEVEL = 7

/// Excel's outlinePr defaults: summary lines below the detail (rows) and to
/// the right of it (columns).
export interface OutlinePlacement {
  readonly summaryBelow: boolean
  readonly summaryRight: boolean
}

export const DEFAULT_OUTLINE_PLACEMENT: OutlinePlacement = {
  summaryBelow: true,
  summaryRight: true,
}

/// The placement for one axis (summaryBelow governs rows, summaryRight
/// columns).
export function placementForAxis(placement: OutlinePlacement, axis: 'rows' | 'cols'): boolean {
  return axis === 'rows' ? placement.summaryBelow : placement.summaryRight
}

/// Computes every group of an axis from its level map. Runs of lines with
/// level >= l nest exactly like Excel's outline: each maximal run at l is a
/// group whose summary sits just outside the run (after it when
/// summary-after, before otherwise). Groups whose summary line would fall
/// outside [0, extent] are dropped — there is nowhere to draw the button.
export function computeOutlineGroups(
  entries: ReadonlyMap<number, OutlineEntry>,
  axis: 'rows' | 'cols',
  summaryAfter: boolean,
  extent: number,
): OutlineGroup[] {
  let maxLevel = 0
  for (const entry of entries.values()) maxLevel = Math.max(maxLevel, entry.level)
  const groups: OutlineGroup[] = []
  const levelOf = (index: number): number => entries.get(index)?.level ?? 0
  for (let level = 1; level <= Math.min(maxLevel, MAX_OUTLINE_LEVEL); level += 1) {
    let start: number | null = null
    for (let index = 0; index <= extent; index += 1) {
      const inside = levelOf(index) >= level
      if (inside && start === null) start = index
      if ((inside && index === extent) || (!inside && start !== null)) {
        const end = inside ? index : index - 1
        const summary = summaryAfter ? end + 1 : (start as number) - 1
        if (summary >= 0 && summary <= extent && levelOf(summary) < level) {
          groups.push({ axis, level, start: start as number, end, summary })
        }
        start = null
      }
    }
  }
  return groups
}

/// The highest outline level present on an axis (0 = no groups).
export function maxOutlineLevel(entries: ReadonlyMap<number, OutlineEntry>): number {
  let max = 0
  for (const entry of entries.values()) max = Math.max(max, entry.level)
  return max
}
