/**
 * BUG-1109: the visual-undo registry grows unbounded — every gutter toggle,
 * Outline Settings change, or visual op parked a closure forever while
 * Univer's own undo stacks cap at 20 entries per unit. The registry is now
 * pruned against the live stacks: a token survives exactly while its
 * mutation sits in some unit's undo or redo stack.
 */
import { LocalUndoRedoService } from '@univerjs/core'
import { beforeEach, describe, expect, it } from 'vitest'

import { VISUAL_UNDO_COMMAND_ID } from '../src/renderer/undo-carry'
import {
  nextVisualUndoToken,
  pruneVisualUndoRegistry,
  visualUndoRegistry,
} from '../src/renderer/visual-undo-registry'
import { installJournalSuppressionUndoFilter } from '../src/renderer/univer-state'

const step = () => ({ undo: () => {}, redo: () => {} })

/// One stack entry carrying a visual-undo mutation for `token`.
const visualItem = (token: number, unit = 'wb') => ({
  unitID: unit,
  undoMutations: [{ id: VISUAL_UNDO_COMMAND_ID, params: { token, direction: 'undo' } }],
  redoMutations: [{ id: VISUAL_UNDO_COMMAND_ID, params: { token, direction: 'redo' } }],
})

const plainItem = (unit = 'wb') => ({ unitID: unit, undoMutations: [], redoMutations: [] })

/// The LocalUndoRedoService internals both the real pushUndoRedo and the
/// prune scan touch (same shape the ai-undo-budget fake uses, plus the
/// stack maps the registry reads).
function fakeService() {
  const undoStacks = new Map<string, unknown[]>()
  const redoStacks = new Map<string, unknown[]>()
  return {
    _undoStacks: undoStacks,
    _redoStacks: redoStacks,
    _getUndoStack: (unitId: string) => {
      const stack = undoStacks.get(unitId) ?? []
      undoStacks.set(unitId, stack)
      return stack
    },
    _getRedoStack: (unitId: string) => {
      const stack = redoStacks.get(unitId) ?? []
      redoStacks.set(unitId, stack)
      return stack
    },
    _pitchUndoElement: (unitId: string) => undoStacks.get(unitId)?.slice(-1)[0] ?? null,
    _batchingStatus: new Map(),
    _updateStatus: () => {},
  }
}

describe('pruneVisualUndoRegistry', () => {
  beforeEach(() => {
    visualUndoRegistry.clear()
  })

  it('keeps tokens on either stack and drops unreachable ones', () => {
    const service = fakeService()
    const undoToken = nextVisualUndoToken(step())
    const redoToken = nextVisualUndoToken(step())
    nextVisualUndoToken(step()) // never attached to any stack
    service._undoStacks.set('a', [visualItem(undoToken, 'a')])
    service._redoStacks.set('b', [visualItem(redoToken, 'b')])
    pruneVisualUndoRegistry(service)
    expect([...visualUndoRegistry.keys()].sort()).toEqual([undoToken, redoToken].sort())
  })

  it('keeps everything when the service shape is not readable', () => {
    const token = nextVisualUndoToken(step())
    pruneVisualUndoRegistry(null)
    pruneVisualUndoRegistry({ _undoStacks: 42 })
    expect(visualUndoRegistry.has(token)).toBe(true)
  })

  it('releases closures once the capacity trim evicts their entry (BUG-1109)', () => {
    installJournalSuppressionUndoFilter()
    const service = fakeService()
    const push = (item: unknown) =>
      (LocalUndoRedoService.prototype.pushUndoRedo as (i: unknown) => void).call(service, item)
    // Univer caps each unit's undo stack at 20 entries; the oldest is
    // spliced off on every push past that.
    const evicted = nextVisualUndoToken(step())
    push(visualItem(evicted))
    for (let index = 0; index < 20; index += 1) push(plainItem())
    expect(service._undoStacks.get('wb')).toHaveLength(20)
    expect(visualUndoRegistry.has(evicted)).toBe(false)

    // A token still inside the capacity window stays undoable.
    const live = nextVisualUndoToken(step())
    push(visualItem(live))
    expect(visualUndoRegistry.has(live)).toBe(true)
  })

  it('releases closures whose only home was the redo stack a new push clears', () => {
    installJournalSuppressionUndoFilter()
    const service = fakeService()
    const push = (item: unknown) =>
      (LocalUndoRedoService.prototype.pushUndoRedo as (i: unknown) => void).call(service, item)
    // Simulate an undone visual step: its entry lives on the redo stack.
    const undone = nextVisualUndoToken(step())
    service._redoStacks.set('wb', [visualItem(undone)])
    push(plainItem())
    expect(service._redoStacks.get('wb')).toHaveLength(0)
    expect(visualUndoRegistry.has(undone)).toBe(false)
  })

  it('caps the registry at the reachable steps over a long session', () => {
    installJournalSuppressionUndoFilter()
    const service = fakeService()
    const push = (item: unknown) =>
      (LocalUndoRedoService.prototype.pushUndoRedo as (i: unknown) => void).call(service, item)
    // Every op registers a step; without the prune this map would hold 200
    // closures forever while only the newest 20 entries can ever undo.
    for (let index = 0; index < 200; index += 1) push(visualItem(nextVisualUndoToken(step())))
    expect(visualUndoRegistry.size).toBeLessThanOrEqual(20)
    expect(visualUndoRegistry.size).toBeGreaterThan(0)
  })
})
