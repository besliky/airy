/**
 * Regression lock for cross-save undo: ⌘S swaps the sidecar session and
 * rebuilds the Univer unit, which used to wipe the undo/redo history (unlike
 * Excel/Word). The save flow must capture the pre-save stack, stash it
 * rewritten against the saved file's unit id, and the reopen's bookkeeping
 * must install it back — dropping the load-time decoration entries so a save
 * itself is never undoable. These tests pin the handleSave ↔ undo-carry
 * wiring (the unit-level rewrite/install logic lives in undo-carry.test.ts).
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'

import {
  createEditJournal,
  recordSetRangeValues,
  recordSheetInsert,
} from '../src/renderer/edit-journal'
import { handleSave, type SaveContext } from '../src/renderer/save-actions'
import type { WorkbookSaveResult } from '../src/shared/desktop-api'
import {
  consumePendingUndoCarry,
  hasPendingUndoCarry,
  stashUndoCarry,
  type CarriedUndoItem,
} from '../src/renderer/undo-carry'
import type { UniverRuntime } from '../src/renderer/univer-state'

const OLD_UNIT = 'file-oldsha'
const NEW_SHA = 'a'.repeat(64)
const NEW_UNIT = `file-${NEW_SHA}`

function cellItem(unitID: string, value: unknown): CarriedUndoItem {
  return {
    unitID,
    undoMutations: [
      { id: 'sheet.mutation.set-range-values', params: { unitId: unitID, cellValue: value } },
    ],
    redoMutations: [
      { id: 'sheet.mutation.set-range-values', params: { unitId: unitID, cellValue: value } },
    ],
  }
}

function tableArtifact(unitID: string): CarriedUndoItem {
  return {
    unitID,
    undoMutations: [{ id: 'sheet.mutation.delete-table', params: { unitId: unitID } }],
    redoMutations: [{ id: 'sheet.mutation.add-table', params: { unitId: unitID } }],
  }
}

interface FakeService {
  _undoStacks: Map<string, CarriedUndoItem[]>
  _redoStacks: Map<string, CarriedUndoItem[]>
  pushed: unknown[]
  statusUpdates: number
  _updateStatus(): void
  clearUndoRedo(unitId: string): void
  pushUndoRedo(item: unknown): void
}

function fakeService(
  unitId: string,
  stack: CarriedUndoItem[],
  redoStack: CarriedUndoItem[] = [],
): FakeService {
  const service: FakeService = {
    _undoStacks: new Map([[unitId, stack]]),
    _redoStacks: new Map([[unitId, redoStack]]),
    pushed: [],
    statusUpdates: 0,
    _updateStatus() {
      service.statusUpdates += 1
    },
    clearUndoRedo(id: string) {
      service._undoStacks.set(id, [])
      service._redoStacks.set(id, [])
    },
    // Mirror the real service: every push empties the unit's redo stack.
    pushUndoRedo(item: unknown) {
      service.pushed.push(item)
      service._redoStacks.set((item as CarriedUndoItem).unitID, [])
    },
  }
  return service
}

function fakeRuntime(service: FakeService, unitId: string): UniverRuntime {
  return {
    univer: { __getInjector: () => ({ get: () => service }) },
    univerAPI: {
      getActiveWorkbook: () => ({
        getId: () => unitId,
        getSnapshot: () => ({ styles: {} }),
        getActiveSheet: () => ({ getSheetId: () => 'sheet-1' }),
        getSheets: () => [{ getSheetId: () => 'sheet-1' }],
      }),
    },
  } as unknown as UniverRuntime
}

const saveWorkbookEdits = vi.fn()

function savedResult(sha: string): WorkbookSaveResult {
  return {
    canceled: false,
    file: {
      sessionId: '22222222-2222-4222-8222-222222222222',
      name: 'book.xlsx',
      sha256: sha,
      entryCount: 1,
      sheets: [],
      styles: [],
      dxfStyles: [],
      visuals: [],
      definedNames: [],
      readOnly: false,
    },
    touchedEntries: [],
  } as unknown as WorkbookSaveResult
}

function buildCtx(runtime: UniverRuntime): {
  ctx: SaveContext
  reopened: unknown[]
  journal: ReturnType<typeof createEditJournal>
} {
  const journal = createEditJournal()
  recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'edited' } } })
  const reopened: unknown[] = []
  return {
    reopened,
    journal,
    ctx: {
      univerRef: { current: runtime },
      stashViewRestore: () => {},
      lazyWorkbookRef: {
        current: {
          editJournal: journal,
          recalc: {
            timer: null,
            generation: 0,
            failed: false,
            formulaCells: new Map(),
            overlay: new Map(),
          },
          file: {
            sessionId: '11111111-1111-4111-8111-111111111111',
            needsSaveAs: false,
          },
        },
      } as never,
      setMessage: () => {},
      openLazyWorkbook: (file) => {
        reopened.push(file)
      },
    },
  }
}

beforeEach(() => {
  saveWorkbookEdits.mockReset().mockResolvedValue(savedResult(NEW_SHA))
  ;(globalThis as unknown as { window: unknown }).window = {
    desktopApi: { saveWorkbookEdits },
  }
  // A previous test's unconsumed stash would make every save opt out.
  stashUndoCarry(null)
})

describe('handleSave undo carry', () => {
  it('a normal save carries the pre-save undo stack across the session swap', async () => {
    const service = fakeService(OLD_UNIT, [cellItem(OLD_UNIT, 1), cellItem(OLD_UNIT, 2)])
    const { ctx, reopened } = buildCtx(fakeRuntime(service, OLD_UNIT))

    await handleSave(ctx, 'save')

    expect(saveWorkbookEdits).toHaveBeenCalledTimes(1)
    // The reopen must see the saved file (new session, new unit id).
    expect(reopened).toHaveLength(1)
    expect((reopened[0] as { sha256: string }).sha256).toBe(NEW_SHA)
    // The carry was stashed for the reopened unit, rewritten off the old id.
    expect(hasPendingUndoCarry()).toBe(true)

    // App-side bookkeeping: the reopened unit's decoration artifacts are
    // swapped for the carried user history.
    const reopenedService = fakeService(NEW_UNIT, [tableArtifact(NEW_UNIT)])
    consumePendingUndoCarry(fakeRuntime(reopenedService, NEW_UNIT), NEW_UNIT)
    expect(reopenedService.pushed).toHaveLength(2)
    const newest = reopenedService.pushed[1] as CarriedUndoItem
    expect(newest.unitID).toBe(NEW_UNIT)
    expect((newest.undoMutations[0] as { params?: { unitId?: string } }).params?.unitId).toBe(
      NEW_UNIT,
    )
  })

  it('the redo stack survives the save too', async () => {
    const service = fakeService(OLD_UNIT, [cellItem(OLD_UNIT, 1)], [cellItem(OLD_UNIT, 2)])
    const { ctx } = buildCtx(fakeRuntime(service, OLD_UNIT))

    await handleSave(ctx, 'save')
    expect(hasPendingUndoCarry()).toBe(true)

    const reopenedService = fakeService(NEW_UNIT, [tableArtifact(NEW_UNIT)])
    consumePendingUndoCarry(fakeRuntime(reopenedService, NEW_UNIT), NEW_UNIT)
    expect(reopenedService.pushed).toHaveLength(1)
    const redos = reopenedService._redoStacks.get(NEW_UNIT) ?? []
    expect(redos).toHaveLength(1)
    expect(redos[0]!.unitID).toBe(NEW_UNIT)
  })

  it('a save itself is not undoable: load artifacts never land on the stack', async () => {
    // A clean session (no user edits to carry) still reopens; the reopen's
    // own decoration entries must be dropped, leaving an empty undo stack.
    const service = fakeService(OLD_UNIT, [])
    const { ctx, reopened } = buildCtx(fakeRuntime(service, OLD_UNIT))

    await handleSave(ctx, 'save')
    expect(reopened).toHaveLength(1)
    expect(hasPendingUndoCarry()).toBe(false)

    const reopenedService = fakeService(NEW_UNIT, [
      tableArtifact(NEW_UNIT),
      tableArtifact(NEW_UNIT),
    ])
    consumePendingUndoCarry(fakeRuntime(reopenedService, NEW_UNIT), NEW_UNIT)
    expect(reopenedService.pushed).toEqual([])
    expect(reopenedService._undoStacks.get(NEW_UNIT)).toEqual([])
  })

  it('Save As carries undo onto the new file unit', async () => {
    const service = fakeService(OLD_UNIT, [cellItem(OLD_UNIT, 1)])
    const { ctx, reopened } = buildCtx(fakeRuntime(service, OLD_UNIT))

    await handleSave(ctx, 'save-as')

    expect(saveWorkbookEdits.mock.calls[0]![0] as { mode: string }).toMatchObject({
      mode: 'save-as',
    })
    expect(reopened).toHaveLength(1)
    expect(hasPendingUndoCarry()).toBe(true)
    const reopenedService = fakeService(NEW_UNIT, [])
    consumePendingUndoCarry(fakeRuntime(reopenedService, NEW_UNIT), NEW_UNIT)
    expect(reopenedService.pushed).toHaveLength(1)
  })

  it('saves with sheet add/remove ops opt out of carrying (documented degradation)', async () => {
    const service = fakeService(OLD_UNIT, [cellItem(OLD_UNIT, 1)])
    const { ctx, journal, reopened } = buildCtx(fakeRuntime(service, OLD_UNIT))
    recordSheetInsert(journal, 'sheet-2', 'Sheet2')

    await handleSave(ctx, 'save')

    expect(reopened).toHaveLength(1)
    expect(hasPendingUndoCarry()).toBe(false)
    const reopenedService = fakeService(NEW_UNIT, [tableArtifact(NEW_UNIT)])
    consumePendingUndoCarry(fakeRuntime(reopenedService, NEW_UNIT), NEW_UNIT)
    expect(reopenedService.pushed).toEqual([])
  })

  it('a save while a previous carry is still unconsumed does not capture a half-settled stack', async () => {
    // Simulate a reopen whose bookkeeping never ran: the stash stays pending.
    stashUndoCarry({ unitId: 'file-unsettled', items: [], redoItems: [] })
    const service = fakeService(OLD_UNIT, [cellItem(OLD_UNIT, 1)])
    const { ctx } = buildCtx(fakeRuntime(service, OLD_UNIT))

    await handleSave(ctx, 'save')

    // The stale carry was replaced with nothing — the next reopen starts clean.
    expect(hasPendingUndoCarry()).toBe(false)
    const reopenedService = fakeService(NEW_UNIT, [tableArtifact(NEW_UNIT)])
    consumePendingUndoCarry(fakeRuntime(reopenedService, NEW_UNIT), NEW_UNIT)
    expect(reopenedService.pushed).toEqual([])
  })

  it('a canceled save leaves the stack untouched and stashes nothing', async () => {
    saveWorkbookEdits.mockResolvedValueOnce({ canceled: true })
    const service = fakeService(OLD_UNIT, [cellItem(OLD_UNIT, 1)])
    const { ctx, reopened } = buildCtx(fakeRuntime(service, OLD_UNIT))

    await handleSave(ctx, 'save')

    expect(reopened).toHaveLength(0)
    expect(hasPendingUndoCarry()).toBe(false)
    // The live pre-save stack is intact — nothing was cleared.
    expect(service._undoStacks.get(OLD_UNIT)).toHaveLength(1)
  })
})
