import { describe, expect, it } from 'vitest'

import {
  tableNameError,
  renameTableInFormulaText,
  tableRefsToA1InFormulaText,
  type TableA1Rewrite,
} from '../src/domain/table-refs'

describe('tableNameError', () => {
  it('accepts ordinary Excel table names', () => {
    expect(tableNameError('Users')).toBeNull()
    expect(tableNameError('Sales_2024')).toBeNull()
    expect(tableNameError('Таблица1')).toBeNull()
    expect(tableNameError('_hidden')).toBeNull()
    expect(tableNameError('A.B.C')).toBeNull()
  })

  it('rejects names Excel refuses', () => {
    expect(tableNameError('')).toMatch(/empty/)
    expect(tableNameError('My Table')).toMatch(/no spaces|only letters/)
    expect(tableNameError('1Table')).toMatch(/must start/)
    expect(tableNameError('A1')).toMatch(/cell reference/)
    expect(tableNameError('XFD1048576')).toMatch(/cell reference/)
    expect(tableNameError('R1C1')).toMatch(/R1C1/)
    expect(tableNameError('C')).toMatch(/reserved/)
    expect(tableNameError('r')).toMatch(/reserved/)
  })
})

describe('renameTableInFormulaText', () => {
  it('rewrites structured references and whole-table tokens', () => {
    expect(renameTableInFormulaText('SUM(Users[Country])', 'Users', 'Clients')).toBe(
      'SUM(Clients[Country])',
    )
    expect(renameTableInFormulaText('SUM(Users)', 'Users', 'Clients')).toBe('SUM(Clients)')
    expect(renameTableInFormulaText('SUM(users[Country])', 'Users', 'Clients')).toBe(
      'SUM(Clients[Country])',
    )
    expect(renameTableInFormulaText('COUNTA(Users[[Country]:[Users]])', 'Users', 'Clients')).toBe(
      'COUNTA(Clients[[Country]:[Users]])',
    )
    expect(renameTableInFormulaText('SUM(Users[#All])', 'Users', 'Clients')).toBe(
      'SUM(Clients[#All])',
    )
    expect(renameTableInFormulaText('SUM(Users[@Country])', 'Users', 'Clients')).toBe(
      'SUM(Clients[@Country])',
    )
  })

  it('rewrites the quoted form and quotes names that need it', () => {
    expect(renameTableInFormulaText("SUM('Old Table'[Col])", 'Old Table', 'New_Table')).toBe(
      'SUM(New_Table[Col])',
    )
    expect(renameTableInFormulaText('SUM(Users[Col])', 'Users', 'New Table')).toBe(
      "SUM('New Table'[Col])",
    )
  })

  it('leaves string literals, sheet names, lookalikes, and calls alone', () => {
    expect(renameTableInFormulaText('IF(A1="Users",Users,0)', 'Users', 'Clients')).toBe(
      'IF(A1="Users",Clients,0)',
    )
    expect(renameTableInFormulaText("'My Sheet'!A1+Users[Col]", 'Users', 'Clients')).toBe(
      "'My Sheet'!A1+Clients[Col]",
    )
    expect(renameTableInFormulaText('SUM(MyUsers)', 'Users', 'Clients')).toBe('SUM(MyUsers)')
    expect(renameTableInFormulaText('Users.Old+1', 'Users', 'Clients')).toBe('Users.Old+1')
    expect(renameTableInFormulaText('Users(A1)', 'Users', 'Clients')).toBe('Users(A1)')
  })
})

const TABLE: TableA1Rewrite = {
  name: 'Users',
  geometry: {
    startRow: 0,
    endRow: 3,
    startColumn: 0,
    endColumn: 1,
    headerRowCount: 1,
    totalsRowCount: 0,
  },
  columns: ['Country', 'Users'],
  sheetName: 'Data',
}

describe('tableRefsToA1InFormulaText', () => {
  it('converts column, span, and selector references to A1 ranges', () => {
    expect(tableRefsToA1InFormulaText('SUM(Users[Users])', TABLE, 0, 'Data')).toBe('SUM(B2:B4)')
    expect(tableRefsToA1InFormulaText('SUM(Users[Country])', TABLE, 0, 'Data')).toBe('SUM(A2:A4)')
    expect(tableRefsToA1InFormulaText('COUNTA(Users[[Country]:[Users]])', TABLE, 0, 'Data')).toBe(
      'COUNTA(A2:B4)',
    )
    expect(tableRefsToA1InFormulaText('SUM(Users[#All])', TABLE, 0, 'Data')).toBe('SUM(A1:B4)')
    expect(tableRefsToA1InFormulaText('SUM(Users[#Headers])', TABLE, 0, 'Data')).toBe('SUM(A1:B1)')
    // The bare whole-table form covers the data body across all columns.
    expect(tableRefsToA1InFormulaText('COUNTA(Users)', TABLE, 0, 'Data')).toBe('COUNTA(A2:B4)')
  })

  it('resolves this-row references against the formula cell and qualifies other sheets', () => {
    expect(tableRefsToA1InFormulaText('Users[@Users]', TABLE, 2, 'Data')).toBe('B3')
    expect(tableRefsToA1InFormulaText('Users[@Users]', TABLE, 8, 'Data')).toBe('#REF!')
    expect(tableRefsToA1InFormulaText('SUM(Users[Users])', TABLE, 0, 'Other')).toBe(
      'SUM(Data!B2:B4)',
    )
    // Structured references are never sheet-qualified; the rewritten A1
    // reference takes the table's sheet since the formula sits elsewhere.
    expect(
      tableRefsToA1InFormulaText("SUM('My Sheet'!A1+Users[Users])", TABLE, 0, 'My Sheet'),
    ).toBe("SUM('My Sheet'!A1+Data!B2:B4)")
  })

  it('maps unknown columns to #REF! and skips string literals', () => {
    expect(tableRefsToA1InFormulaText('SUM(Users[Missing])', TABLE, 0, 'Data')).toBe('SUM(#REF!)')
    expect(
      tableRefsToA1InFormulaText('IF(A1="Users[Users]",1,Users[Users])', TABLE, 0, 'Data'),
    ).toBe('IF(A1="Users[Users]",1,B2:B4)')
  })

  it('never rewrites table-name lookalikes', () => {
    expect(tableRefsToA1InFormulaText('SUM(MyUsers)', TABLE, 0, 'Data')).toBe('SUM(MyUsers)')
  })
})
