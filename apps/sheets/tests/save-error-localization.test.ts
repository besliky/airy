import { beforeAll, describe, expect, it } from 'vitest'

import { localizeSaveError, stripIpcErrorWrapper } from '../src/renderer/save-actions'
import { loadLocale } from '../src/renderer/i18n/locale'

// BUG-1202: since PERF-901 the dictionaries load lazily per locale. Without
// loadLocale('zh'), t() returns the raw English-ish key, which still passed
// the old not-null/not-same asserts — the file stopped checking localization
// altogether. Loading zh makes every assert below verify a real translation.
beforeAll(() => loadLocale('zh'))

describe('localizeSaveError', () => {
  it('maps user-reachable gateway errors to localized messages', () => {
    const localized = [
      'Charts or tables on "Sheet1" cannot shift here.',
      'This sheet has extended (x14) data validation — editing its rules is not supported yet.',
      'Multi-select list rules are Univer-only and cannot be saved to xlsx — delete the rule before saving.',
      'This range has extended conditional formatting (x14) that cannot be modified yet',
      'This range has a data-bar extension format (x14) that cannot be modified yet',
      'A new pivot cannot be saved together with sheet management changes — save the pivot first.',
      'A new pivot cannot be saved together with row/column changes on its sheets — save the pivot first.',
      'A new table cannot be saved together with row/column changes on its sheet — save the table first.',
      'Defined-name edits cannot be saved together with row/column or sheet changes.',
      'The workbook changed on disk while saving — aborted.',
      'The workbook stylesheet is missing fonts, fills, or cellXfs — style edits cannot be saved.',
      'Saving would change the workbook package structure — aborted.',
      'Only pie and doughnut charts support explosion.',
      'Converting a scatter is not supported.',
      'Replacing series on a pie is not supported.',
    ]
    for (const message of localized) {
      const mapped = localizeSaveError(message)
      expect(mapped, message).not.toBeNull()
      // the mapped string is a real zh translation — not the raw key that an
      // unloaded dictionary would hand back
      expect(mapped, message).toMatch(/[\u4e00-\u9fff]/)
    }
  })

  it('passes unknown messages through as null', () => {
    expect(localizeSaveError('Relationship rId9 not found.')).toBeNull()
    expect(localizeSaveError('')).toBeNull()
  })

  it('passes the pre-save disk-changed error through: main already localized it, and its Save As advice must survive', () => {
    expect(
      localizeSaveError('The workbook changed on disk after it was opened — use Save As instead.'),
    ).toBeNull()
    expect(localizeSaveError('工作簿在打开后被磁盘上的改动覆盖——请改用另存为。')).toBeNull()
  })
})

describe('stripIpcErrorWrapper', () => {
  it('unwraps Electron remote-method errors', () => {
    expect(
      stripIpcErrorWrapper(
        "Error invoking remote method 'workbook:save': Error: 工作簿在打开后被磁盘上的改动覆盖——请改用另存为。",
      ),
    ).toBe('工作簿在打开后被磁盘上的改动覆盖——请改用另存为。')
    expect(stripIpcErrorWrapper("Error invoking remote method 'workbook:save': boom")).toBe('boom')
  })

  it('leaves plain messages alone', () => {
    expect(stripIpcErrorWrapper('The workbook changed on disk while saving — aborted.')).toBe(
      'The workbook changed on disk while saving — aborted.',
    )
  })
})
