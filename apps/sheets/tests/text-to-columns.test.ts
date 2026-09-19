/**
 * PAR-208 Text to Columns: multi-delimiter and fixed-width splitting,
 * per-column typing (General / Text / Date with DMY-MDY-YMD orders),
 * break-position and destination parsing, and the apply path writing
 * journaled values through the worksheet range.
 */
import { beforeEach, describe, expect, it } from 'vitest'

import {
  activeDelimiterChars,
  coerceFieldValue,
  excelDateSerial,
  parseBreakPositions,
  parseDestinationCell,
  splitDelimited,
  splitFixedWidth,
  type TextToColumnsConfig,
} from '../src/renderer/text-to-columns'
import { handleTextToColumns, type DataToolsContext } from '../src/renderer/data-tools-actions'

describe('splitDelimited', () => {
  it('splits by any checked delimiter character', () => {
    expect(splitDelimited('a,b;c', [',', ';'], false)).toEqual(['a', 'b', 'c'])
    expect(splitDelimited('a b\tc', [' ', '\t'], false)).toEqual(['a', 'b', 'c'])
  })

  it('keeps empty fields between consecutive delimiters unless collapsed', () => {
    expect(splitDelimited('a,,b', [','], false)).toEqual(['a', '', 'b'])
    expect(splitDelimited('a,,b', [','], true)).toEqual(['a', 'b'])
    expect(splitDelimited('a,;b', [',', ';'], true)).toEqual(['a', 'b'])
  })

  it('treats every character of the custom box as a delimiter, regex-special ones included', () => {
    expect(
      activeDelimiterChars({
        tab: false,
        semicolon: false,
        comma: false,
        space: false,
        custom: '|',
        consecutiveAsOne: false,
      }),
    ).toEqual(['|'])
    expect(splitDelimited('a|b', ['|'], false)).toEqual(['a', 'b'])
    expect(splitDelimited('a.b', ['.'], false)).toEqual(['a', 'b'])
  })

  it('returns the whole text when no delimiter is active', () => {
    expect(splitDelimited('a,b', [], false)).toEqual(['a,b'])
  })
})

describe('fixed width', () => {
  it('parses, sorts, and dedupes break positions', () => {
    expect(parseBreakPositions('12, 5 5;20')).toEqual([5, 12, 20])
    expect(parseBreakPositions('')).toEqual([])
  })

  it('rejects non-numeric or out-of-range positions', () => {
    expect(parseBreakPositions('5,x')).toBeNull()
    expect(parseBreakPositions('-3')).toBeNull()
    expect(parseBreakPositions('0')).toBeNull()
    expect(parseBreakPositions('99999')).toBeNull()
  })

  it('splits at the break positions with the tail kept', () => {
    expect(splitFixedWidth('abcdefgh', [3, 5])).toEqual(['abc', 'de', 'fgh'])
    expect(splitFixedWidth('abc', [5])).toEqual(['abc'])
    expect(splitFixedWidth('abc', [])).toEqual(['abc'])
  })
})

describe('coerceFieldValue', () => {
  it('General types numbers, thousands separators, percents, and booleans', () => {
    expect(coerceFieldValue('42', 'general')).toEqual({ v: 42 })
    expect(coerceFieldValue('-3.5', 'general')).toEqual({ v: -3.5 })
    expect(coerceFieldValue('1,234.5', 'general')).toEqual({ v: 1234.5 })
    expect(coerceFieldValue('12%', 'general')).toEqual({ v: 0.12 })
    expect(coerceFieldValue('TRUE', 'general')).toEqual({ v: true })
    expect(coerceFieldValue('false', 'general')).toEqual({ v: false })
    expect(coerceFieldValue(' 7 ', 'general')).toEqual({ v: 7 })
    expect(coerceFieldValue('x1', 'general')).toEqual({ v: 'x1' })
    expect(coerceFieldValue('', 'general')).toEqual({ v: null })
  })

  it('Text keeps every field verbatim', () => {
    expect(coerceFieldValue('007', 'text')).toEqual({ v: '007' })
    expect(coerceFieldValue(' 1,2 ', 'text')).toEqual({ v: ' 1,2 ' })
  })

  it('parses dates in the chosen field order with Excel two-digit years', () => {
    expect(coerceFieldValue('31/12/2024', 'date-dmy')).toEqual({ v: excelDateSerial(2024, 12, 31) })
    expect(coerceFieldValue('12/31/2024', 'date-mdy')).toEqual({ v: excelDateSerial(2024, 12, 31) })
    expect(coerceFieldValue('2024-12-31', 'date-ymd')).toEqual({ v: excelDateSerial(2024, 12, 31) })
    expect(coerceFieldValue('1.2.24', 'date-dmy')).toEqual({ v: excelDateSerial(2024, 2, 1) })
    expect(coerceFieldValue('1.2.99', 'date-dmy')).toEqual({ v: excelDateSerial(1999, 2, 1) })
  })

  it('keeps unparsable dates as text', () => {
    expect(coerceFieldValue('31/02/2024', 'date-dmy')).toEqual({ v: '31/02/2024' })
    expect(coerceFieldValue('13/13/2024', 'date-mdy')).toEqual({ v: '13/13/2024' })
    expect(coerceFieldValue('hello', 'date-dmy')).toEqual({ v: 'hello' })
  })

  it('maps known dates to Excel serials', () => {
    expect(excelDateSerial(2024, 2, 29)).toBe(45351)
    expect(excelDateSerial(1900, 1, 1)).toBe(1)
    expect(excelDateSerial(2023, 12, 31)).toBe(45291)
  })
})

describe('parseDestinationCell', () => {
  it('accepts plain same-sheet A1 references', () => {
    expect(parseDestinationCell('B2')).toEqual({ row: 1, column: 1 })
    expect(parseDestinationCell('$a$1')).toEqual({ row: 0, column: 0 })
  })

  it('rejects ranges, sheet qualifiers, and out-of-sheet rows', () => {
    expect(parseDestinationCell('A1:B2')).toBeNull()
    expect(parseDestinationCell('Sheet2!A1')).toBeNull()
    expect(parseDestinationCell('A0')).toBeNull()
    expect(parseDestinationCell('A1048577')).toBeNull()
  })
})

/// A fake active sheet capturing the written matrix and number formats.
function makeContext(displayGrid: string[][]) {
  const written: { row: number; column: number; rows: number; columns: number; values: unknown }[] =
    []
  const formats: { row: number; column: number; rows: number; columns: number; pattern: string }[] =
    []
  const worksheet = {
    getSheetId: () => 'sheet-1',
    getRange: (row: number, column: number, rows: number, columns: number) => ({
      getDisplayValues: () =>
        displayGrid.slice(row, row + rows).map((cells) => cells.slice(column, column + columns)),
      setValues: (values: unknown) => {
        written.push({ row, column, rows, columns, values })
      },
      setNumberFormat: (pattern: string) => {
        formats.push({ row, column, rows, columns, pattern })
      },
    }),
  }
  const workbook = {
    getActiveSheet: () => worksheet,
    getActiveRange: () => ({
      getRow: () => 0,
      getColumn: () => 0,
      getWidth: () => 1,
      getHeight: () => 3,
    }),
  }
  return {
    ctx: {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: { current: null },
      setMessage: () => {},
      setPendingEdits: () => {},
      setAdvancedFilterColumns: () => {},
    } as unknown as DataToolsContext,
    written,
    formats,
  }
}

const BASE_CONFIG: TextToColumnsConfig = {
  mode: 'delimited',
  delimiters: {
    tab: false,
    semicolon: false,
    comma: true,
    space: false,
    custom: '',
    consecutiveAsOne: false,
  },
  breaks: [],
  columnTypes: [],
  destination: null,
}

describe('handleTextToColumns', () => {
  let messages: string[]
  beforeEach(() => {
    messages = []
  })

  it('writes the parsed matrix over the source column with per-column types', () => {
    const { ctx, written, formats } = makeContext([
      ['a,1,2024-01-31'],
      ['b,2,2024-02-29'],
      ['c,3,x'],
    ])
    ctx.setMessage = (message: string) => messages.push(message)
    const error = handleTextToColumns(ctx, {
      ...BASE_CONFIG,
      columnTypes: ['text', 'general', 'date-ymd'],
    })
    expect(error).toBeNull()
    expect(written).toHaveLength(1)
    expect(written[0]).toMatchObject({ row: 0, column: 0, rows: 3, columns: 3 })
    expect(written[0]?.values).toEqual([
      [{ v: 'a' }, { v: 1 }, { v: excelDateSerial(2024, 1, 31) }],
      [{ v: 'b' }, { v: 2 }, { v: excelDateSerial(2024, 2, 29) }],
      [{ v: 'c' }, { v: 3 }, { v: 'x' }],
    ])
    // Only the date-typed column gains the date number format.
    expect(formats).toEqual([{ row: 0, column: 2, rows: 3, columns: 1, pattern: 'yyyy-mm-dd' }])
  })

  it('writes at a destination reference instead of the source column', () => {
    const { ctx, written } = makeContext([['a;b'], ['c;d']])
    const error = handleTextToColumns(ctx, {
      ...BASE_CONFIG,
      delimiters: { ...BASE_CONFIG.delimiters, comma: false, semicolon: true },
      destination: 'D4',
    })
    expect(error).toBeNull()
    expect(written[0]).toMatchObject({ row: 3, column: 3, rows: 2, columns: 2 })
  })

  it('reports a bad destination without writing', () => {
    const { ctx, written } = makeContext([['a,b']])
    const error = handleTextToColumns(ctx, { ...BASE_CONFIG, destination: 'not-a-cell' })
    expect(error).toBeTruthy()
    expect(written).toHaveLength(0)
  })

  it('requires a delimiter in delimited mode', () => {
    const { ctx } = makeContext([['a,b']])
    const error = handleTextToColumns(ctx, {
      ...BASE_CONFIG,
      delimiters: { ...BASE_CONFIG.delimiters, comma: false },
    })
    expect(error).toBeTruthy()
  })

  it('splits fixed-width at the configured positions', () => {
    const { ctx, written } = makeContext([['aabbb'], ['ccddd']])
    const error = handleTextToColumns(ctx, {
      ...BASE_CONFIG,
      mode: 'fixed-width',
      breaks: [2],
    })
    expect(error).toBeNull()
    expect(written[0]?.values).toEqual([
      [{ v: 'aa' }, { v: 'bbb' }],
      [{ v: 'cc' }, { v: 'ddd' }],
    ])
  })

  it('refuses a multi-column selection', () => {
    const { ctx } = makeContext([['a', 'b']])
    const wideRange = {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => ({
              getActiveSheet: () => ({}),
              getActiveRange: () => ({
                getRow: () => 0,
                getColumn: () => 0,
                getWidth: () => 2,
                getHeight: () => 1,
              }),
            }),
          },
        },
      },
      lazyWorkbookRef: { current: null },
      setMessage: () => {},
      setPendingEdits: () => {},
      setAdvancedFilterColumns: () => {},
    } as unknown as DataToolsContext
    const error = handleTextToColumns(wideRange, BASE_CONFIG)
    expect(error).toBeTruthy()
    expect(ctx).toBeDefined()
  })
})
