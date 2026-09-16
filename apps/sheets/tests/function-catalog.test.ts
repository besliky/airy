/**
 * The Insert Function catalog derives from the live Univer registry
 * (executors for names, FUNCTION_NAMES_* enums for categories, the merged
 * locale's functionList for descriptions/signatures) with the curated 61
 * as an override layer. These tests pin the derivation and merge logic.
 */
import { describe, expect, it } from 'vitest'

import {
  buildFunctionCatalog,
  catalogCategories,
  FALLBACK_CATEGORY,
  type CuratedFunction,
  type FunctionListLocale,
} from '../src/renderer/function-catalog'

const CURATED: readonly CuratedFunction[] = [
  {
    name: 'SUM',
    category: 'Math',
    syntax: 'SUM(number1, [number2], …)',
    descKey: 'dlgFnDescSum' as never,
  },
  {
    name: 'VLOOKUP',
    category: 'Lookup',
    syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
    descKey: 'dlgFnDescVlookup' as never,
  },
]

const LOCALE: FunctionListLocale = {
  'sheets-formula': {
    functionList: {
      SUM: { description: 'Adds numbers.', abstract: 'Adds numbers.' },
      DELTA: {
        abstract: 'Tests whether two values are equal.',
        functionParameter: { number1: { name: 'number1' }, number2: { name: '[number2]' } },
      },
      ISEVEN: { description: 'Even check.', functionParameter: {} },
    },
  },
}

function registryOf(...names: string[]): { getExecutors(): Map<string, object> } {
  return { getExecutors: () => new Map(names.map((name) => [name, {}])) }
}

describe('buildFunctionCatalog', () => {
  it('derives entries from the registry executors, sorted by name', () => {
    const catalog = buildFunctionCatalog(registryOf('SUM', 'DELTA', 'ACCRINT'), LOCALE, CURATED)
    expect(catalog.map((entry) => entry.name)).toEqual(['ACCRINT', 'DELTA', 'SUM'])
  })

  it('curated entries keep their hand-written help and category', () => {
    const catalog = buildFunctionCatalog(registryOf('SUM'), LOCALE, CURATED)
    const sum = catalog.find((entry) => entry.name === 'SUM')
    expect(sum).toEqual({
      name: 'SUM',
      category: 'Math',
      syntax: 'SUM(number1, [number2], …)',
      descKey: 'dlgFnDescSum',
      description: undefined,
    })
  })

  it('derives category, description, and syntax for non-curated functions', () => {
    const catalog = buildFunctionCatalog(registryOf('DELTA', 'ISEVEN', 'ACCRINT'), LOCALE, CURATED)
    const delta = catalog.find((entry) => entry.name === 'DELTA')
    expect(delta?.category).toBe('Engineering')
    expect(delta?.description).toBe('Tests whether two values are equal.')
    expect(delta?.syntax).toBe('DELTA(number1, [number2])')
    // no parameters listed: no derived syntax
    expect(catalog.find((entry) => entry.name === 'ISEVEN')?.syntax).toBe('')
    expect(catalog.find((entry) => entry.name === 'ACCRINT')?.category).toBe('Financial')
  })

  it('maps the builtin enums to the curated category ids', () => {
    const catalog = buildFunctionCatalog(
      registryOf('IF', 'COUNTIF', 'INDEX', 'TEXT', 'TODAY', 'PMT', 'DGET', 'ISBLANK', 'ERROR.TYPE'),
      LOCALE,
      CURATED,
    )
    const categoryOf = (name: string) => catalog.find((entry) => entry.name === name)?.category
    expect(categoryOf('IF')).toBe('Logical')
    expect(categoryOf('COUNTIF')).toBe('Statistical')
    expect(categoryOf('INDEX')).toBe('Lookup')
    expect(categoryOf('TEXT')).toBe('Text')
    expect(categoryOf('TODAY')).toBe('Date & Time')
    expect(categoryOf('PMT')).toBe('Financial')
    expect(categoryOf('DGET')).toBe('Database')
    expect(categoryOf('ISBLANK')).toBe('Information')
    expect(categoryOf('ERROR.TYPE')).toBe('Information')
  })

  it('unknown executor names land in the fallback category', () => {
    const catalog = buildFunctionCatalog(registryOf('MY_CUSTOM_FN'), LOCALE, CURATED)
    expect(catalog[0]?.category).toBe(FALLBACK_CATEGORY)
  })

  it('falls back to the curated list when the registry has not landed yet', () => {
    const empty = buildFunctionCatalog({ getExecutors: () => new Map() }, LOCALE, CURATED)
    expect(empty.map((entry) => entry.name)).toEqual(['SUM', 'VLOOKUP'])
    const nullRegistry = buildFunctionCatalog(null, LOCALE, CURATED)
    expect(nullRegistry).toHaveLength(2)
  })

  it('deduplicates executor keys and ignores blanks', () => {
    const catalog = buildFunctionCatalog(registryOf('SUM', 'SUM', ''), LOCALE, CURATED)
    expect(catalog.filter((entry) => entry.name === 'SUM')).toHaveLength(1)
  })
})

describe('catalogCategories', () => {
  it('lists curated categories first, derived-only ones alphabetically', () => {
    const catalog = buildFunctionCatalog(
      registryOf('SUM', 'DELTA', 'DGET', 'MY_CUSTOM_FN'),
      LOCALE,
      CURATED,
    )
    expect(catalogCategories(catalog, ['Math', 'Lookup', 'Financial'])).toEqual([
      'Math',
      'Database',
      'Engineering',
      'More',
    ])
  })

  it('drops curated categories with no matching entries', () => {
    const catalog = buildFunctionCatalog(registryOf('DELTA'), LOCALE, CURATED)
    expect(catalogCategories(catalog, ['Math', 'Lookup', 'Financial'])).toEqual(['Engineering'])
  })
})
