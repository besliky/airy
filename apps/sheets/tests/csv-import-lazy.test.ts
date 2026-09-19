/**
 * Data → From Text/CSV rides a lazily imported chunk (PERF-902): the file
 * read, the dynamic import, and the legacy-charset decode all run before
 * importCsvText's own try/catch. BUG-1201: any failure there must surface as
 * a localized message instead of an unhandled rejection that leaves the
 * picked file doing nothing.
 */
import { beforeAll, describe, expect, it, vi } from 'vitest'

import { importCsvFile } from '../src/renderer/data-tools-actions'
import { loadLocale } from '../src/renderer/i18n/locale'

const { decodeCsvBuffer, parseCsv, isNumericCell } = vi.hoisted(() => ({
  decodeCsvBuffer: vi.fn(),
  parseCsv: vi.fn(),
  isNumericCell: vi.fn(),
}))

vi.mock('../src/gateway/csv-import', () => ({ decodeCsvBuffer, parseCsv, isNumericCell }))

// t() must return real zh strings: without loadLocale the translator falls
// back to raw keys, which is exactly the regression this suite guards.
beforeAll(() => loadLocale('zh'))

interface Ctx {
  messages: string[]
  univerRef: { current: unknown }
  setMessage: (message: string) => void
  setPendingEdits: (count: number) => void
  setAdvancedFilterColumns: (columns: unknown) => void
}

function makeCtx(setValues = vi.fn()): Ctx {
  const messages: string[] = []
  const worksheet = { getRange: vi.fn(() => ({ setValues })) }
  const workbook = {
    getActiveSheet: () => worksheet,
    getActiveRange: () => ({ getRow: () => 2, getColumn: () => 1 }),
  }
  return {
    messages,
    univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
    setMessage: (message) => messages.push(message),
    setPendingEdits: () => {},
    setAdvancedFilterColumns: () => {},
  }
}

describe('importCsvFile', () => {
  it('streams the file through the lazy chunk into the selection', async () => {
    decodeCsvBuffer.mockReturnValue('a,b\n1,x')
    parseCsv.mockImplementation((text: string) => text.split('\n').map((line) => line.split(',')))
    isNumericCell.mockImplementation((cell: string) => /^\d+$/.test(cell))
    const setValues = vi.fn()
    const ctx = makeCtx(setValues)

    await importCsvFile(ctx, new File(['a,b\n1,x'], 'in.csv', { type: 'text/csv' }))

    // the raw bytes reach the decoder, and the grid receives typed values
    expect(decodeCsvBuffer.mock.calls[0]?.[0]).toEqual(
      Uint8Array.from(['a', ',', 'b', '\n', '1', ',', 'x'].map((ch) => ch.charCodeAt(0))),
    )
    expect(setValues).toHaveBeenCalledWith([
      [{ v: 'a' }, { v: 'b' }],
      [{ v: 1 }, { v: 'x' }],
    ])
    expect(ctx.messages).toHaveLength(1)
    // localized success line, not a raw i18n key
    expect(ctx.messages[0]).toMatch(/[\u4e00-\u9fff]/)
    expect(ctx.messages[0]).not.toBe('appCsvImported')
  })

  it('reports a localized failure when the decode step throws', async () => {
    decodeCsvBuffer.mockImplementation(() => {
      throw new Error('unsupported legacy charset')
    })
    const ctx = makeCtx()

    // must resolve (no unhandled rejection) and leave a visible message
    await expect(importCsvFile(ctx, new File(['x'], 'in.csv'))).resolves.toBeUndefined()

    expect(ctx.messages).toEqual(['CSV 导入失败。'])
  })

  it('reports the same failure when the lazy chunk cannot load', async () => {
    decodeCsvBuffer.mockImplementation(() => {
      throw new Error('Failed to fetch dynamically imported module')
    })
    const ctx = makeCtx()

    await expect(importCsvFile(ctx, new File(['x'], 'in.csv'))).resolves.toBeUndefined()

    expect(ctx.messages).toEqual(['CSV 导入失败。'])
  })
})
