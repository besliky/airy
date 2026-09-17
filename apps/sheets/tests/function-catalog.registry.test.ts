import { LocaleType, UniverInstanceType } from '@univerjs/core'
import { IFunctionService } from '@univerjs/engine-formula'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import { describe, expect, it } from 'vitest'
import { createUniver } from '../src/renderer/create-univer'
import { buildFunctionCatalog } from '../src/renderer/function-catalog'
import { CURATED_FUNCTIONS } from '../src/renderer/InsertFunctionDialog'

describe('Insert Function catalog from the live registry', () => {
  it('derives hundreds of functions with categories and descriptions', async () => {
    const runtime = createUniver({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: UniverPresetSheetsCoreEnUS as never },
      presets: [
        { plugins: [UniverFormulaEnginePlugin] },
        { plugins: [UniverSheetsPlugin] },
        { plugins: [UniverSheetsFormulaPlugin] },
      ],
    })
    runtime.univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
      id: 'wb1',
      sheetOrder: ['main'],
      name: 'wb',
      styles: {},
      sheets: { main: { id: 'main', name: 'Main', rowCount: 5, columnCount: 5, cellData: {} } },
    })
    await new Promise((resolve) => setTimeout(resolve, 50))
    const svc = runtime.univer.__getInjector().get(IFunctionService)
    const catalog = buildFunctionCatalog(svc, UniverPresetSheetsCoreEnUS, CURATED_FUNCTIONS)
    // the engine registers hundreds of executors — far beyond the curated 61
    expect(catalog.length).toBeGreaterThan(300)
    const byName = new Map(catalog.map((entry) => [entry.name, entry]))
    // curated entries keep their hand-written help
    expect(byName.get('SUM')?.descKey).toBe('dlgFnDescSum')
    expect(byName.get('SUM')?.category).toBe('Math')
    // derived entries carry locale descriptions and enum categories
    const delta = byName.get('DELTA')
    expect(delta?.category).toBe('Engineering')
    expect(delta?.description).toBeTruthy()
    expect(delta?.syntax).toContain('DELTA(')
    expect(byName.get('ISBLANK')?.category).toBe('Information')
    expect(byName.get('DGET')?.category).toBe('Database')
    // names are unique and sorted
    for (let index = 1; index < catalog.length; index += 1) {
      expect(catalog[index - 1]!.name.localeCompare(catalog[index]!.name)).toBeLessThanOrEqual(0)
    }
    expect(new Set(catalog.map((entry) => entry.name)).size).toBe(catalog.length)
  })
})
