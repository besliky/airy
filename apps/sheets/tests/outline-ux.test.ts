/**
 * PAR-209 outline UX: group math over the streamed level map (summary
 * below/above), the Data-tab commands journaling declaratively with visual
 * undo, and the Outline Settings placement round-trip into sheetPr's
 * outlinePr.
 */
import { describe, expect, it, vi } from 'vitest'

import { ICommandService } from '@univerjs/core'

import { computeOutlineGroups, maxOutlineLevel, type OutlineEntry } from '../src/renderer/outline'
import {
  handleOutline,
  handleOutlineSettings,
  outlinePlacement,
  toggleOutlineGroup,
  type DataToolsContext,
} from '../src/renderer/data-tools-actions'
import { createEditJournal, type EditJournal } from '../src/renderer/edit-journal'
import { VISUAL_UNDO_COMMAND_ID } from '../src/renderer/undo-carry'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

describe('computeOutlineGroups', () => {
  const levels = (spec: Record<number, number>): Map<number, OutlineEntry> =>
    new Map(
      Object.entries(spec).map(([index, level]) => [Number(index), { level, collapsed: false }]),
    )

  it('finds each maximal run per level with the summary after it', () => {
    const rows = levels({ 2: 1, 3: 1, 4: 2, 5: 2, 8: 1, 9: 1 })
    const groups = computeOutlineGroups(rows, 'rows', true, 20)
    expect(groups).toContainEqual({ axis: 'rows', level: 1, start: 2, end: 5, summary: 6 })
    expect(groups).toContainEqual({ axis: 'rows', level: 2, start: 4, end: 5, summary: 6 })
    expect(groups).toContainEqual({ axis: 'rows', level: 1, start: 8, end: 9, summary: 10 })
    // No duplicate per-summary-and-level buttons.
    expect(groups.filter((group) => group.summary === 6 && group.level === 1)).toHaveLength(1)
  })

  it('places the summary before the run when summaryBelow is off', () => {
    const rows = levels({ 3: 1, 4: 1 })
    const groups = computeOutlineGroups(rows, 'rows', false, 20)
    expect(groups).toEqual([{ axis: 'rows', level: 1, start: 3, end: 4, summary: 2 }])
  })

  it('drops groups whose summary line falls outside the sheet', () => {
    const atStart = computeOutlineGroups(levels({ 0: 1, 1: 1 }), 'rows', false, 10)
    expect(atStart).toEqual([])
    // Sheet with rows 0..10: a run ending at the last row has no summary.
    const atEnd = computeOutlineGroups(levels({ 9: 1, 10: 1 }), 'rows', true, 10)
    expect(atEnd).toEqual([])
  })

  it('reports the highest level', () => {
    expect(maxOutlineLevel(levels({ 1: 3, 2: 7 }))).toBe(7)
    expect(maxOutlineLevel(new Map())).toBe(0)
  })
})

/// A fake runtime/worksheet pair capturing hide/show calls and undo pushes.
function makeOutlineContext() {
  const hideRows = vi.fn()
  const showRows = vi.fn()
  const hideColumns = vi.fn()
  const showColumns = vi.fn()
  const worksheet = {
    getSheetId: () => 'sheet-1',
    getSheetName: () => 'S1',
    hideRows,
    showRows,
    hideColumns,
    showColumns,
  }
  const registeredCommands: {
    id: string
    handler: (accessor: unknown, params?: { token: number; direction: 'undo' | 'redo' }) => boolean
  }[] = []
  const undoPushes: {
    undoMutations: { id: string; params: { token: number; direction: 'undo' | 'redo' } }[]
    redoMutations: { id: string; params: { token: number; direction: 'undo' | 'redo' } }[]
  }[] = []
  const injector = {
    get: (token: unknown) =>
      token === ICommandService
        ? {
            registerCommand: (command: (typeof registeredCommands)[number]) =>
              registeredCommands.push(command),
          }
        : { pushUndoRedo: (item: (typeof undoPushes)[number]) => undoPushes.push(item) },
  }
  const runtime = {
    univerAPI: {
      getActiveWorkbook: () => ({
        getId: () => 'wb-1',
        getActiveSheet: () => worksheet,
        getActiveRange: () => ({
          getRow: () => 1,
          getColumn: () => 0,
          getHeight: () => 3,
          getWidth: () => 3,
        }),
      }),
    },
    univer: { __getInjector: () => injector },
  } as unknown as UniverRuntime
  const editJournal: EditJournal = createEditJournal()
  const state = {
    file: { sheets: [], name: 'Book.xlsx', sessionId: 'session-1' },
    editJournal,
    sheetFilePageSetups: new Map(),
    outline: new Map(),
    flags: { preloadComplete: true, preloadRunning: false },
  } as unknown as LazyWorkbookState
  const messages: string[] = []
  const pendingEdits: number[] = []
  let outlineChanges = 0
  const ctx: DataToolsContext = {
    univerRef: { current: runtime },
    lazyWorkbookRef: { current: state },
    setMessage: (message: string) => messages.push(message),
    setPendingEdits: (count: number) => pendingEdits.push(count),
    setAdvancedFilterColumns: () => {},
    onOutlineChanged: () => {
      outlineChanges += 1
    },
  }
  /// Replays the newest visual-undo entry in one direction.
  const replay = (direction: 'undo' | 'redo'): void => {
    const pushed = undoPushes.at(-1)
    expect(pushed).toBeDefined()
    const command = registeredCommands.find((entry) => entry.id === VISUAL_UNDO_COMMAND_ID)
    expect(command).toBeDefined()
    const mutation = direction === 'undo' ? pushed!.undoMutations[0] : pushed!.redoMutations[0]
    expect(mutation!.id).toBe(VISUAL_UNDO_COMMAND_ID)
    expect(command!.handler(null, mutation!.params)).toBe(true)
  }
  return {
    ctx,
    state,
    editJournal,
    replay,
    counts: () => ({
      hidden: hideRows.mock.calls.length + hideColumns.mock.calls.length,
      shown: showRows.mock.calls.length + showColumns.mock.calls.length,
      pushes: undoPushes.length,
      messages,
      outlineChanges,
      hideRows,
      showRows,
    }),
  }
}

describe('handleOutline group/ungroup', () => {
  it('groups the selection one level deeper and undoes back through the journal', () => {
    const env = makeOutlineContext()
    handleOutline(env.ctx, 'group', 'rows')
    const ops = env.editJournal.structuralOps.get('sheet-1') ?? []
    expect(ops).toEqual([
      { kind: 'set-rows-outline', start: 1, end: 3, level: 1, collapsed: undefined },
    ])
    expect(env.state.outline.get('sheet-1')?.rows.get(2)?.level).toBe(1)

    env.replay('undo')
    // The undo re-journals the previous level instead of leaving the edit in.
    const afterUndo = env.editJournal.structuralOps.get('sheet-1') ?? []
    expect(afterUndo.at(-1)).toEqual({
      kind: 'set-rows-outline',
      start: 1,
      end: 3,
      level: 0,
    })
    expect(env.state.outline.get('sheet-1')?.rows.get(2)?.level).toBe(0)

    env.replay('redo')
    expect(env.state.outline.get('sheet-1')?.rows.get(2)?.level).toBe(1)
    expect(env.counts().pushes).toBe(1)
    expect(env.counts().outlineChanges).toBeGreaterThanOrEqual(2)
  })

  it('clamps ungroup at level 0 and reports nothing to undo', () => {
    const env = makeOutlineContext()
    handleOutline(env.ctx, 'ungroup', 'rows')
    expect(env.editJournal.structuralOps.get('sheet-1')).toBeUndefined()
    expect(env.counts().pushes).toBe(0)
  })
})

describe('toggleOutlineGroup (gutter click and hide/show detail)', () => {
  it('collapses through the hidden pipeline with the collapsed flag, and one undo restores both', () => {
    const env = makeOutlineContext()
    toggleOutlineGroup(env.ctx, 'rows', { start: 2, end: 5 }, 6, true)
    expect(env.counts().hideRows).toHaveBeenCalledWith(2, 4)
    const outline = env.state.outline.get('sheet-1')!
    expect(outline.rows.get(6)).toEqual({ level: 0, collapsed: true })
    expect(env.editJournal.structuralOps.get('sheet-1')).toContainEqual({
      kind: 'set-rows-outline',
      start: 6,
      end: 6,
      level: 0,
      collapsed: true,
    })

    env.replay('undo')
    expect(env.counts().showRows).toHaveBeenCalledWith(2, 4)
    expect(outline.rows.get(6)?.collapsed).toBe(false)
    expect(env.editJournal.structuralOps.get('sheet-1')?.at(-1)).toMatchObject({
      collapsed: false,
    })

    env.replay('redo')
    expect(outline.rows.get(6)?.collapsed).toBe(true)
  })

  it('hide-detail honors a summary-above placement', () => {
    const env = makeOutlineContext()
    handleOutlineSettings(env.ctx, { summaryBelow: false, summaryRight: true })
    handleOutline(env.ctx, 'hide-detail', 'rows')
    // Selection rows 2-4 (0-based 1-3) with the summary above → row 0.
    expect(env.state.outline.get('sheet-1')?.rows.get(0)?.collapsed).toBe(true)
    expect(env.counts().hideRows).toHaveBeenCalledWith(1, 3)
  })
})

describe('outlinePlacement and Outline Settings', () => {
  it('journals the placement over Excel defaults and undoes it', () => {
    const env = makeOutlineContext()
    expect(outlinePlacement(env.state, 'sheet-1')).toEqual({
      summaryBelow: true,
      summaryRight: true,
    })
    expect(handleOutlineSettings(env.ctx, { summaryBelow: false, summaryRight: false })).toBeNull()
    expect(outlinePlacement(env.state, 'sheet-1')).toEqual({
      summaryBelow: false,
      summaryRight: false,
    })
    expect(env.editJournal.pageSetup.get('sheet-1')).toEqual({
      outlineSummaryBelow: false,
      outlineSummaryRight: false,
    })
    env.replay('undo')
    expect(outlinePlacement(env.state, 'sheet-1')).toEqual({
      summaryBelow: true,
      summaryRight: true,
    })
  })

  it('seeds from the file pageSetup over the defaults', () => {
    const env = makeOutlineContext()
    env.state.sheetFilePageSetups.set('sheet-1', {
      outlineSummaryBelow: false,
    } as LazyWorkbookState['sheetFilePageSetups'] extends Map<string, infer T> ? T : never)
    expect(outlinePlacement(env.state, 'sheet-1')).toEqual({
      summaryBelow: false,
      summaryRight: true,
    })
  })
})
