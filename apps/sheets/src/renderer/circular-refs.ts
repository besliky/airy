/**
 * Excel badges circular references in the status bar and (with iterative
 * calculation enabled) converges them over iterateCount passes. This
 * editor's engine instead resolves a reference cycle in a single pass and
 * silently shows the one-pass numbers (A1=B1+1 / B1=A1+1 → 2 and 1), even
 * when the file's calcPr asks for iterative recalculation — which the
 * engine does not perform either (BUG-1718). This module statically detects
 * reference cycles over a workbook's formula list so the editor can surface
 * them instead of staying silent.
 *
 * The analysis walks a bipartite graph — formula cells and the coordinate
 * ranges they reference — without expanding ranges to cells: a range points
 * back only at the formula cells it actually covers, so whole-column
 * references on million-row sheets stay bounded by the formula count, not
 * by the grid. Finding no cycle says nothing about evaluation-order quirks;
 * every reported hit is a cycle the engine resolves in one silent pass.
 */
import { qualifierMatches } from '../gateway/xlsx-structure'

import { type ClosureSheetInput, parseFormulaReferences } from './formula-closure'

export interface CircularRefHit {
  readonly sheetId: string
  readonly row: number
  readonly column: number
}

/// Above this the graph build would dominate open latency for a warning;
/// such workbooks are far beyond the sizes this badge targets.
const MAX_ANALYZED_FORMULAS = 150_000

interface RangeNode {
  readonly sheetIndex: number
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/** Returns every formula cell that participates in a reference cycle,
 * ordered by sheet then position. An empty result means "no cycles found"
 * (or the list was too large to analyze) — never a guarantee. */
export function findCircularFormulas(sheets: readonly ClosureSheetInput[]): CircularRefHit[] {
  const formulaCount = sheets.reduce((sum, sheet) => sum + sheet.formulas.length, 0)
  if (formulaCount === 0 || formulaCount > MAX_ANALYZED_FORMULAS) return []

  const sheetIndexById = new Map(sheets.map((sheet, index) => [sheet.id, index] as const))
  const sheetByLowerName = new Map(
    sheets.map((sheet, index) => [sheet.name.toLowerCase(), index] as const),
  )
  const resolveSheet = (qualifier: string | undefined): number | undefined => {
    if (qualifier === undefined) return undefined
    for (let index = 0; index < sheets.length; index += 1) {
      if (qualifierMatches(qualifier, sheets[index]?.name ?? '')) return index
    }
    const unquoted = qualifier.startsWith("'")
      ? qualifier.slice(1, -1).replaceAll("''", "'")
      : qualifier
    return sheetByLowerName.get(unquoted.toLowerCase())
  }

  // Nodes: 0..formulaCount-1 are formula cells, then deduplicated ranges.
  interface FormulaNode {
    readonly sheetIndex: number
    readonly row: number
    readonly column: number
  }
  const formulas: FormulaNode[] = []
  for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex += 1) {
    const sheet = sheets[sheetIndex] as ClosureSheetInput
    for (const cell of sheet.formulas) {
      formulas.push({ sheetIndex, row: cell.row, column: cell.column })
    }
  }

  const edges: number[][] = formulas.map(() => [])
  const rangeNodes: RangeNode[] = []
  const rangeNodeIndex = new Map<string, number>()
  const rangeEdges: number[][] = []
  const internRange = (node: RangeNode): number => {
    const key = `${node.sheetIndex}|${node.startRow}|${node.endRow}|${node.startColumn}|${node.endColumn}`
    const existing = rangeNodeIndex.get(key)
    if (existing !== undefined) return existing
    const index = rangeNodes.length
    rangeNodes.push(node)
    rangeEdges.push([])
    rangeNodeIndex.set(key, index)
    return index
  }

  const formulaBase: number[] = []
  let cursor = 0
  for (const sheet of sheets) {
    formulaBase.push(cursor)
    cursor += sheet.formulas.length
  }

  for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex += 1) {
    const sheet = sheets[sheetIndex] as ClosureSheetInput
    const base = formulaBase[sheetIndex] ?? 0
    for (let cellIndex = 0; cellIndex < sheet.formulas.length; cellIndex += 1) {
      const cell = sheet.formulas[cellIndex] as ClosureSheetInput['formulas'][number]
      const nodeIndex = base + cellIndex
      for (const reference of parseFormulaReferences(cell.formula)) {
        const targetIndex =
          reference.qualifier === undefined ? sheetIndex : resolveSheet(reference.qualifier)
        if (targetIndex === undefined) continue
        const target = sheets[targetIndex] as ClosureSheetInput
        const clamped: RangeNode = {
          sheetIndex: targetIndex,
          startRow: Math.max(reference.token.startRow ?? 0, 0),
          endRow: Math.min(reference.token.endRow ?? target.rowCount - 1, target.rowCount - 1),
          startColumn: Math.max(reference.token.startColumn ?? 0, 0),
          endColumn: Math.min(
            reference.token.endColumn ?? target.columnCount - 1,
            target.columnCount - 1,
          ),
        }
        if (clamped.endRow < clamped.startRow || clamped.endColumn < clamped.startColumn) {
          continue
        }
        // Range nodes live in the index space after every formula node.
        edges[nodeIndex]?.push(formulaCount + internRange(clamped))
      }
    }
  }

  // Range → formula edges: only formulas the range actually covers (per
  // sheet, formulas are position-sorted so the band scan stays linear in
  // the covered rows).
  for (let rangeIndex = 0; rangeIndex < rangeNodes.length; rangeIndex += 1) {
    const range = rangeNodes[rangeIndex] as RangeNode
    const sheet = sheets[range.sheetIndex] as ClosureSheetInput
    const base = formulaBase[range.sheetIndex] ?? 0
    const out = rangeEdges[rangeIndex] as number[]
    for (let cellIndex = 0; cellIndex < sheet.formulas.length; cellIndex += 1) {
      const cell = sheet.formulas[cellIndex] as ClosureSheetInput['formulas'][number]
      if (
        cell.row >= range.startRow &&
        cell.row <= range.endRow &&
        cell.column >= range.startColumn &&
        cell.column <= range.endColumn
      ) {
        out.push(base + cellIndex)
      }
    }
  }

  const cyclic = findCyclicNodes(formulaCount + rangeNodes.length, (nodeIndex) => {
    if (nodeIndex < formulaCount) return edges[nodeIndex] as number[]
    return rangeEdges[nodeIndex - formulaCount] as number[]
  })

  const hits: CircularRefHit[] = []
  for (const nodeIndex of cyclic) {
    if (nodeIndex >= formulaCount) continue
    const node = formulas[nodeIndex] as FormulaNode
    const sheet = sheets[node.sheetIndex] as ClosureSheetInput
    hits.push({ sheetId: sheet.id, row: node.row, column: node.column })
  }
  hits.sort(
    (a, b) =>
      sheetIndexById.get(a.sheetId)! - sheetIndexById.get(b.sheetId)! ||
      a.row - b.row ||
      a.column - b.column,
  )
  return hits
}

/** Iterative Tarjan SCC over a node-indexed graph; returns every node that
 * sits on a cycle (in an SCC of two or more nodes, or self-reachable). */
export function findCyclicNodes(
  nodeCount: number,
  edgesOf: (nodeIndex: number) => readonly number[],
): number[] {
  const UNVISITED = -1
  const indexOf = new Int32Array(nodeCount).fill(UNVISITED)
  const low = new Int32Array(nodeCount)
  const onStack = new Uint8Array(nodeCount)
  const stack: number[] = []
  const cyclic: boolean[] = new Array(nodeCount).fill(false)
  let counter = 0

  for (let root = 0; root < nodeCount; root += 1) {
    if (indexOf[root] !== UNVISITED) continue
    // Frames of [node, edgeCursor]; the root frame drives the outer loop.
    const callStack: number[][] = [[root, 0]]
    indexOf[root] = low[root] = counter++
    stack.push(root)
    onStack[root] = 1
    while (callStack.length > 0) {
      const frame = callStack[callStack.length - 1] as number[]
      const node = frame[0] as number
      const outgoing = edgesOf(node)
      if ((frame[1] as number) < outgoing.length) {
        const next = outgoing[frame[1] as number] as number
        frame[1] = (frame[1] as number) + 1
        if (next === node) cyclic[node] = true
        if (indexOf[next] !== UNVISITED) {
          if (onStack[next]) low[node] = Math.min(low[node] as number, indexOf[next] as number)
          continue
        }
        indexOf[next] = low[next] = counter++
        stack.push(next)
        onStack[next] = 1
        callStack.push([next, 0])
        continue
      }
      callStack.pop()
      const parent = callStack[callStack.length - 1]?.[0]
      if (parent !== undefined) {
        low[parent] = Math.min(low[parent] as number, low[node] as number)
      }
      if (low[node] === indexOf[node]) {
        // Root of an SCC: pop its members.
        const members: number[] = []
        for (;;) {
          const member = stack.pop() as number
          onStack[member] = 0
          members.push(member)
          if (member === node) break
        }
        if (members.length > 1) {
          for (const member of members) cyclic[member] = true
        }
      }
    }
  }
  const result: number[] = []
  for (let node = 0; node < nodeCount; node += 1) {
    if (cyclic[node]) result.push(node)
  }
  return result
}

/** A1-style address of a hit ("B1"), the form Excel's status bar uses. */
export function circularRefAddress(hit: CircularRefHit): string {
  let letters = ''
  let remaining = hit.column + 1
  while (remaining > 0) {
    remaining -= 1
    letters = String.fromCharCode(65 + (remaining % 26)) + letters
    remaining = Math.floor(remaining / 26)
  }
  return `${letters}${hit.row + 1}`
}

/** Badge list: the first addresses plus a count of the rest ("A1, B1, C1 +2"). */
export function formatCircularRefList(hits: readonly CircularRefHit[], max = 3): string {
  return formatAddressList(hits.map(circularRefAddress), max)
}

/** Same cap for already-formatted address lists (the status-bar badge only
 * carries addresses once detection has run). */
export function formatAddressList(addresses: readonly string[], max = 3): string {
  if (addresses.length <= max) return addresses.join(', ')
  return `${addresses.slice(0, max).join(', ')} +${addresses.length - max}`
}
