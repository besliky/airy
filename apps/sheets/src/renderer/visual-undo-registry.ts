/**
 * The visual-undo step registry. Interactive visual and outline edits land
 * on Univer's own undo stack as custom mutations whose params are opaque
 * tokens resolving to closures stored here (univer-sync pushes them, the
 * mutation handler replays them). A token is only reachable while its
 * mutation sits in some unit's undo or redo stack — once Univer evicts the
 * entry (undo-stack capacity trim, the redo clear every push performs, or
 * clearUndoRedo at a session swap) the closure would otherwise be retained
 * for the whole session (BUG-1109), so the stack-touching paths prune the
 * registry against the live stacks.
 */
import { VISUAL_UNDO_COMMAND_ID } from './undo-carry'

export interface VisualUndoStep {
  undo(): void
  redo(): void
}

/// Token → step closure. Exported for the prune hook and tests; production
/// code goes through nextVisualUndoToken.
export const visualUndoRegistry = new Map<number, VisualUndoStep>()

let visualUndoSequence = 0

/// Registers one step and returns the token its stack mutations carry.
export function nextVisualUndoToken(step: VisualUndoStep): number {
  visualUndoSequence += 1
  visualUndoRegistry.set(visualUndoSequence, step)
  return visualUndoSequence
}

/// The LocalUndoRedoService internals the reachability scan reads (the same
/// private stack maps undo-carry uses).
interface UndoRedoStacksLike {
  readonly _undoStacks?: Map<string, readonly StackItemLike[] | undefined> | undefined
  readonly _redoStacks?: Map<string, readonly StackItemLike[] | undefined> | undefined
}

interface StackItemLike {
  readonly undoMutations?: readonly { id?: unknown; params?: unknown }[] | undefined
  readonly redoMutations?: readonly { id?: unknown; params?: unknown }[] | undefined
}

/// Drops registry entries whose mutation no longer sits in any unit's undo
/// or redo stack. Never throws: the internals shape is Univer's own, and a
/// prune failure must not break whatever push triggered it — on any doubt
/// every entry stays.
export function pruneVisualUndoRegistry(service: unknown): void {
  if (visualUndoRegistry.size === 0) return
  try {
    const stacks = service as UndoRedoStacksLike
    const reachable = new Set<number>()
    const collect = (items: readonly StackItemLike[] | undefined): void => {
      for (const item of items ?? []) {
        for (const mutation of [...(item.undoMutations ?? []), ...(item.redoMutations ?? [])]) {
          if (mutation?.id !== VISUAL_UNDO_COMMAND_ID) continue
          const token = (mutation.params as { token?: unknown } | undefined)?.token
          if (typeof token === 'number') reachable.add(token)
        }
      }
    }
    for (const items of stacks._undoStacks?.values() ?? []) collect(items)
    for (const items of stacks._redoStacks?.values() ?? []) collect(items)
    for (const token of [...visualUndoRegistry.keys()]) {
      if (!reachable.has(token)) visualUndoRegistry.delete(token)
    }
  } catch {
    // Keep everything rather than risk dropping a reachable step.
  }
}
